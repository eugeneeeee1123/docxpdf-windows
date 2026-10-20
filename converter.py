from __future__ import annotations

import gc
import hashlib
import errno
import os
import platform
import posixpath
import shutil
import threading
import time
import uuid
from collections.abc import Callable, Iterable
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path
from types import ModuleType
from typing import Any
from xml.etree import ElementTree
from zipfile import BadZipFile, ZipFile

from i18n import DEFAULT_LANGUAGE, normalize_language, tr


WD_ALERTS_NONE = 0
WD_DO_NOT_SAVE_CHANGES = 0
WD_EXPORT_FORMAT_PDF = 17
WD_EXPORT_OPTIMIZE_FOR_PRINT = 0
WD_EXPORT_ALL_DOCUMENT = 0
WD_EXPORT_DOCUMENT_CONTENT = 0
WD_EXPORT_CREATE_HEADING_BOOKMARKS = 1
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3
# Upper bound for one document's open + export inside Word when the caller does
# not pass a timeout. Past it, the watchdog ends that session's WINWORD.EXE.
DEFAULT_WORD_TIMEOUT = 600.0


class ConversionError(RuntimeError):
    """A user-facing conversion failure."""

    def __init__(self, message: str, *, abort_batch: bool = False) -> None:
        super().__init__(message)
        self.abort_batch = abort_batch


class MergeCancelled(ConversionError):
    """The merge was cancelled before publishing its output."""

    def __init__(self) -> None:
        super().__init__("PDF merge cancelled")


class ConversionCancelled(ConversionError):
    """The conversion was cancelled after Word's export, before publishing its output."""

    def __init__(self) -> None:
        super().__init__("Conversion cancelled")


@dataclass(frozen=True)
class ConversionResult:
    source: Path
    output: Path
    size_bytes: int
    elapsed_seconds: float
    images_restored: int = 0
    images_examined: int = 0
    output_relocated: bool = False

    @property
    def images_unrestored(self) -> int:
        """PDF images that still carry Word's (possibly downsampled) pixels."""
        return max(0, self.images_examined - self.images_restored)


@dataclass(frozen=True)
class MergeResult:
    inputs: tuple[Path, ...]
    output: Path
    page_count: int
    size_bytes: int
    elapsed_seconds: float
    output_relocated: bool = False


@dataclass(frozen=True)
class ImageRestorationReport:
    images_examined: int
    images_restored: int
    images_unmatched: int


@dataclass
class _DocxImage:
    name: str
    image: Any
    digest: str
    comparison_sample: Any
    # Pixel box of a picture cropped inside Word. Word exports only this region,
    # so it is matched and restored instead of the whole media file.
    crop_box: tuple[int, int, int, int] | None = None
    # Lazily built samples of the source composited onto a PDF /Matte colour.
    matte_samples: dict[tuple[int, int, int], Any] = field(default_factory=dict)

    @property
    def size(self) -> tuple[int, int]:
        if self.crop_box is None:
            return self.image.size
        left, top, right, bottom = self.crop_box
        return right - left, bottom - top


def _candidate_from_override(raw: str) -> Path | None:
    candidate = Path(raw).expanduser()
    if candidate.is_dir():
        candidate = candidate / "WINWORD.EXE"
    return candidate.resolve() if candidate.is_file() else None


def locate_word_app() -> Path | None:
    """Locate the registered desktop Microsoft Word executable without launching it."""
    override = os.environ.get("DOCX_PDF_WORD_APP")
    if override:
        return _candidate_from_override(override)

    if platform.system() != "Windows":
        return None

    try:
        import winreg
    except ImportError:
        return None

    app_path_key = r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE"
    access_modes = [winreg.KEY_READ]
    for flag_name in ("KEY_WOW64_64KEY", "KEY_WOW64_32KEY"):
        flag = getattr(winreg, flag_name, 0)
        if flag:
            access_modes.append(winreg.KEY_READ | flag)

    for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
        for access in access_modes:
            try:
                with winreg.OpenKey(hive, app_path_key, 0, access) as key:
                    raw, _kind = winreg.QueryValueEx(key, None)
            except OSError:
                continue
            candidate = Path(str(raw).strip().strip('"')).expanduser()
            if candidate.is_file():
                return candidate.resolve()

    command = shutil.which("WINWORD.EXE")
    return Path(command).resolve() if command else None


def word_version(word_app: Path | None = None) -> str | None:
    """Read the installed Click-to-Run Word version without starting Word."""
    if platform.system() != "Windows" or (word_app or locate_word_app()) is None:
        return None
    try:
        import winreg
    except ImportError:
        return None

    version_key = r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
    access_modes = [winreg.KEY_READ]
    for flag_name in ("KEY_WOW64_64KEY", "KEY_WOW64_32KEY"):
        flag = getattr(winreg, flag_name, 0)
        if flag:
            access_modes.append(winreg.KEY_READ | flag)

    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        for access in access_modes:
            try:
                with winreg.OpenKey(hive, version_key, 0, access) as key:
                    for value_name in ("ClientVersionToReport", "VersionToReport"):
                        try:
                            value, _kind = winreg.QueryValueEx(key, value_name)
                        except OSError:
                            continue
                        version = str(value).strip()
                        if version:
                            return version
            except OSError:
                continue
    return None


def validate_source(source: str | Path, *, language: str = DEFAULT_LANGUAGE) -> Path:
    path = Path(source).expanduser()
    if not path.is_file():
        raise ConversionError(tr(language, "source_not_found", path=path))
    if path.suffix.lower() != ".docx":
        raise ConversionError(tr(language, "source_extension", name=path.name))
    if path.name.startswith("~$"):
        raise ConversionError(tr(language, "source_lock", name=path.name))
    return path.resolve()


def ensure_output_dir(
    output_dir: str | Path,
    *,
    language: str = DEFAULT_LANGUAGE,
) -> Path:
    folder = Path(output_dir).expanduser().resolve()
    try:
        folder.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise ConversionError(tr(language, "output_dir_create", error=exc)) from exc
    if not folder.is_dir():
        raise ConversionError(tr(language, "output_not_dir", path=folder))
    return folder


def next_output_path(
    source: str | Path,
    output_dir: str | Path,
    overwrite: bool = False,
    *,
    language: str = DEFAULT_LANGUAGE,
    reserved_paths: Iterable[str | Path] | None = None,
) -> Path:
    source_path = Path(source)
    folder = Path(output_dir).expanduser()
    reserved = {
        Path(path).expanduser().resolve()
        for path in (reserved_paths or ())
    }
    target = folder / f"{source_path.stem}.pdf"
    if (overwrite or not target.exists()) and target.resolve() not in reserved:
        return target

    for number in range(1, 10_000):
        candidate = folder / f"{source_path.stem}-{number}.pdf"
        if not candidate.exists() and candidate.resolve() not in reserved:
            return candidate
    raise ConversionError(tr(language, "too_many_names", path=folder))


def next_merge_path(
    output_dir: str | Path,
    overwrite: bool = False,
    *,
    language: str = DEFAULT_LANGUAGE,
) -> Path:
    """Choose a safe name for the combined PDF without deleting an existing file."""
    source_name = "合并结果.docx" if normalize_language(language) == "zh" else "Merged.docx"
    return next_output_path(source_name, output_dir, overwrite, language=language)


def _check_pdf(path: Path, *, language: str = DEFAULT_LANGUAGE) -> int:
    try:
        size = path.stat().st_size
        with path.open("rb") as stream:
            signature = stream.read(5)
    except OSError as exc:
        raise ConversionError(tr(language, "pdf_read", error=exc)) from exc
    if size < 8 or signature != b"%PDF-":
        raise ConversionError(tr(language, "pdf_invalid"))
    return size


def _wait_for_pdf(path: Path, timeout: float = 8.0) -> bool:
    """Wait for Word or a synced folder to finish publishing a stable PDF."""
    deadline = time.monotonic() + max(0.2, timeout)
    previous_size = -1
    stable_reads = 0
    while time.monotonic() < deadline:
        try:
            size = path.stat().st_size
        except OSError:
            size = -1
        if size >= 8 and size == previous_size:
            stable_reads += 1
            if stable_reads >= 2:
                return True
        else:
            stable_reads = 0
        previous_size = size
        time.sleep(0.1)
    return path.is_file()


def _replace_with_retry(source: Path, destination: Path, *, timeout: float = 10.0) -> None:
    """Atomically replace a PDF, tolerating brief sync-client file locks."""
    deadline = time.monotonic() + max(0.0, timeout)
    delay = 0.08
    retryable_errors = {errno.EACCES, errno.EPERM, 5, 32, 33}
    while True:
        try:
            os.replace(source, destination)
            return
        except OSError as exc:
            error_code = exc.winerror if getattr(exc, "winerror", None) else exc.errno
            if error_code not in retryable_errors or time.monotonic() >= deadline:
                raise
            time.sleep(min(delay, max(0.01, deadline - time.monotonic())))
            delay = min(delay * 2, 0.8)


def _publish_pdf(
    source: Path,
    destination: Path,
    *,
    language: str = DEFAULT_LANGUAGE,
    retry_timeout: float = 10.0,
) -> tuple[Path, bool]:
    """Publish a staged PDF without losing work to an occupied destination.

    OneDrive and PDF viewers can keep a just-synced output file open. First
    wait for that transient lock. If it stays occupied, preserve the result
    under the next available filename instead of failing the completed job.
    """
    try:
        _replace_with_retry(source, destination, timeout=retry_timeout)
        return destination, False
    except OSError as original_error:
        error_code = (
            original_error.winerror
            if getattr(original_error, "winerror", None)
            else original_error.errno
        )
        if error_code not in {errno.EACCES, errno.EPERM, 5, 32, 33}:
            raise

    alternate = next_output_path(destination, destination.parent, language=language)
    _replace_with_retry(source, alternate, timeout=retry_timeout)
    return alternate, True


def _load_com_modules(language: str = DEFAULT_LANGUAGE) -> tuple[ModuleType, Any]:
    try:
        import pythoncom
        from win32com import client as win32_client
    except ImportError as exc:
        raise ConversionError(
            tr(language, "pywin32_missing"),
            abort_batch=True,
        ) from exc
    return pythoncom, win32_client


def _com_error_detail(exc: Exception, language: str = DEFAULT_LANGUAGE) -> str:
    candidates: list[str] = []

    def collect(value: Any) -> None:
        if isinstance(value, str):
            cleaned = " ".join(value.split())
            if cleaned and cleaned not in candidates:
                candidates.append(cleaned)
        elif isinstance(value, (tuple, list)):
            for item in value:
                collect(item)

    collect(getattr(exc, "args", ()))
    collect(str(exc))
    if not candidates:
        return tr(language, "word_no_detail")
    detail = max(candidates, key=len)
    return detail[:600]


def _friendly_com_error(
    exc: Exception,
    *,
    during_startup: bool = False,
    language: str = DEFAULT_LANGUAGE,
) -> ConversionError:
    detail = _com_error_detail(exc, language)
    lowered = detail.lower()
    hresult = getattr(exc, "hresult", None)
    if hresult is None and getattr(exc, "args", None):
        hresult = exc.args[0] if isinstance(exc.args[0], int) else None

    if during_startup or hresult in {-2147221005, -2146959355} or any(
        marker in lowered
        for marker in ("class not registered", "invalid class string", "server execution failed")
    ):
        return ConversionError(tr(language, "word_start_failed"), abort_batch=True)
    if hresult in {-2147418111, -2147417848, -2147023174} or any(
        marker in lowered
        for marker in (
            "call was rejected by callee",
            "rpc server is unavailable",
            "object invoked has disconnected",
        )
    ):
        return ConversionError(tr(language, "word_busy"), abort_batch=True)
    if any(marker in lowered for marker in ("password", "encrypted", "protected")):
        return ConversionError(tr(language, "document_protected"))
    if hresult == -2147024891 or any(
        marker in lowered for marker in ("access is denied", "permission denied", "read-only location")
    ):
        return ConversionError(tr(language, "word_permission"))
    return ConversionError(tr(language, "word_export_failed", detail=detail))


COMPARISON_SAMPLE_SIZE = (64, 64)
IMAGE_MATCH_THRESHOLD = 0.985
IMAGE_MATCH_AMBIGUITY_MARGIN = 0.002
WHITE = (255, 255, 255)

_PACKAGE_RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_OFFICE_RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_DRAWINGML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
# DrawingML crop edges are stored in thousandths of a percent.
_DRAWINGML_FULL_EXTENT = 100_000


def _has_alpha(image: Any) -> bool:
    return "A" in image.getbands() or "transparency" in image.info


def _flatten_for_comparison(
    image: Any,
    image_module: Any,
    background: tuple[int, int, int] = WHITE,
) -> Any:
    if _has_alpha(image):
        rgba = image.convert("RGBA")
        backdrop = image_module.new("RGBA", rgba.size, (*background, 255))
        try:
            return image_module.alpha_composite(backdrop, rgba).convert("RGB")
        finally:
            rgba.close()
            backdrop.close()
    return image.convert("RGB")


def _comparison_sample(
    image: Any,
    image_module: Any,
    background: tuple[int, int, int] = WHITE,
) -> Any:
    flattened = _flatten_for_comparison(image, image_module, background)
    try:
        return flattened.resize(COMPARISON_SAMPLE_SIZE, image_module.Resampling.LANCZOS)
    finally:
        flattened.close()


def _color_channel_sample(image: Any, image_module: Any) -> Any:
    """Sample only the stored colour of a soft-masked PDF image, ignoring alpha."""
    color = image.convert("RGB")
    try:
        return color.resize(COMPARISON_SAMPLE_SIZE, image_module.Resampling.LANCZOS)
    finally:
        color.close()


def _image_matte(image_object: Any) -> tuple[int, int, int] | None:
    """Return the RGB /Matte colour a soft-masked PDF image was pre-blended with.

    Word exports semi-transparent pictures this way. pypdf decodes the colour
    channels as stored, so they only resemble the source after the source is
    composited onto the same matte colour.
    """
    try:
        soft_mask = image_object.get("/SMask")
        if soft_mask is None:
            return None
        matte = soft_mask.get_object().get("/Matte")
        if matte is None:
            return None
        values = [
            min(255, max(0, round(float(value) * 255)))
            for value in matte.get_object()
        ]
    except Exception:
        return None
    if len(values) == 1:
        return values[0], values[0], values[0]
    if len(values) == 3:
        return values[0], values[1], values[2]
    return None


def _drawingml_percentage(value: str | None) -> int:
    if not value:
        return 0
    text = value.strip()
    if text.endswith("%"):
        return round(float(text[:-1]) * 1000)
    return int(text)


def _docx_picture_crops(archive: ZipFile) -> dict[str, set[tuple[int, int, int, int]]]:
    """Map media parts to the a:srcRect crops of the pictures that display them.

    Crop discovery is best effort: an unreadable part only means its pictures are
    matched as whole images, exactly as before crop support existed.
    """
    names = set(archive.namelist())
    crops: dict[str, set[tuple[int, int, int, int]]] = {}
    for part in sorted(names):
        if not part.startswith("word/") or not part.endswith(".xml") or "/_rels/" in part:
            continue
        folder, base = posixpath.split(part)
        rels_name = f"{folder}/_rels/{base}.rels"
        if rels_name not in names:
            continue
        try:
            raw = archive.read(part)
            if b"srcRect" not in raw:
                continue
            targets: dict[str, str] = {}
            for relationship in ElementTree.fromstring(archive.read(rels_name)):
                if relationship.tag != f"{{{_PACKAGE_RELATIONSHIPS_NS}}}Relationship":
                    continue
                if relationship.get("TargetMode") == "External":
                    continue
                rel_id = relationship.get("Id")
                target = relationship.get("Target")
                if not rel_id or not target:
                    continue
                if target.startswith("/"):
                    targets[rel_id] = posixpath.normpath(target.lstrip("/"))
                else:
                    targets[rel_id] = posixpath.normpath(posixpath.join(folder, target))
            root = ElementTree.fromstring(raw)
        except Exception:
            continue
        for blip_fill in root.iter():
            if not isinstance(blip_fill.tag, str) or not blip_fill.tag.endswith("}blipFill"):
                continue
            blip = blip_fill.find(f"{{{_DRAWINGML_NS}}}blip")
            source_rect = blip_fill.find(f"{{{_DRAWINGML_NS}}}srcRect")
            if blip is None or source_rect is None:
                continue
            media = targets.get(blip.get(f"{{{_OFFICE_RELATIONSHIPS_NS}}}embed", ""))
            if media is None:
                continue
            try:
                rect = tuple(
                    _drawingml_percentage(source_rect.get(side))
                    for side in ("l", "t", "r", "b")
                )
            except (ValueError, OverflowError):
                continue
            if any(rect):
                crops.setdefault(media, set()).add(rect)
    return crops


def _crop_box(
    size: tuple[int, int],
    rect: tuple[int, int, int, int],
) -> tuple[int, int, int, int] | None:
    """Convert DrawingML crop edges into a Pillow box, or None for no real crop."""
    width, height = size
    # Negative edges pad the picture frame; they do not remove source pixels.
    left_edge, top_edge, right_edge, bottom_edge = (max(0, value) for value in rect)
    left = round(width * left_edge / _DRAWINGML_FULL_EXTENT)
    top = round(height * top_edge / _DRAWINGML_FULL_EXTENT)
    right = round(width * (_DRAWINGML_FULL_EXTENT - right_edge) / _DRAWINGML_FULL_EXTENT)
    bottom = round(height * (_DRAWINGML_FULL_EXTENT - bottom_edge) / _DRAWINGML_FULL_EXTENT)
    if right - left < 1 or bottom - top < 1:
        return None
    box = (left, top, right, bottom)
    return None if box == (0, 0, width, height) else box


def _candidate_pixels(candidate: _DocxImage) -> Any:
    """Return the source pixels Word placed: the crop region for cropped pictures."""
    if candidate.crop_box is None:
        return candidate.image
    return candidate.image.crop(candidate.crop_box)


def _docx_candidate(
    name: str,
    image: Any,
    crop_box: tuple[int, int, int, int] | None,
    image_module: Any,
) -> _DocxImage | None:
    region = image.crop(crop_box) if crop_box else image
    flattened = None
    try:
        flattened = _flatten_for_comparison(region, image_module)
        digest = hashlib.sha256(
            f"{flattened.width}x{flattened.height}:".encode("ascii")
            + flattened.tobytes()
        ).hexdigest()
        comparison_sample = flattened.resize(
            COMPARISON_SAMPLE_SIZE,
            image_module.Resampling.LANCZOS,
        )
    except Exception:
        return None
    finally:
        if flattened is not None:
            flattened.close()
        if region is not image:
            region.close()
    return _DocxImage(
        name=name,
        image=image,
        digest=digest,
        comparison_sample=comparison_sample,
        crop_box=crop_box,
    )


def _matte_sample(
    candidate: _DocxImage,
    matte: tuple[int, int, int],
    image_module: Any,
) -> Any:
    if not _has_alpha(candidate.image):
        return candidate.comparison_sample
    sample = candidate.matte_samples.get(matte)
    if sample is None:
        pixels = _candidate_pixels(candidate)
        try:
            sample = _comparison_sample(pixels, image_module, background=matte)
        finally:
            if pixels is not candidate.image:
                pixels.close()
        candidate.matte_samples[matte] = sample
    return sample


def _close_candidates(candidates: list[_DocxImage]) -> None:
    closed_images: set[int] = set()
    for candidate in candidates:
        # Crop candidates share the decoded media image with the whole-image one.
        if id(candidate.image) not in closed_images:
            closed_images.add(id(candidate.image))
            candidate.image.close()
        candidate.comparison_sample.close()
        for sample in candidate.matte_samples.values():
            sample.close()
        candidate.matte_samples.clear()


def _sample_similarity(
    source_sample: Any,
    exported_sample: Any,
    image_chops: Any,
    image_stat: Any,
) -> float:
    difference = None
    try:
        difference = image_chops.difference(source_sample, exported_sample)
        rms = image_stat.Stat(difference).rms
        normalized = (sum(value * value for value in rms) / len(rms)) ** 0.5 / 255.0
        return max(0.0, 1.0 - normalized)
    finally:
        if difference is not None:
            difference.close()


def _load_docx_images(
    source_path: Path,
    *,
    language: str = DEFAULT_LANGUAGE,
    cancel_requested: Callable[[], bool] | None = None,
) -> list[_DocxImage]:
    try:
        from PIL import Image, ImageOps
    except ImportError as exc:
        raise ConversionError(tr(language, "pillow_restore_install")) from exc

    candidates: list[_DocxImage] = []
    try:
        with ZipFile(source_path) as archive:
            crops = _docx_picture_crops(archive)
            media_names = sorted(
                name
                for name in archive.namelist()
                if name.lower().startswith("word/media/") and not name.endswith("/")
            )
            for name in media_names:
                if cancel_requested and cancel_requested():
                    _close_candidates(candidates)
                    raise ConversionCancelled()
                try:
                    with Image.open(BytesIO(archive.read(name))) as opened:
                        image = ImageOps.exif_transpose(opened).copy()
                        image.load()
                except Exception:
                    continue
                # PDF image streams used here are 1- or 8-bit per component.
                # Skip uncommon high-bit-depth formats rather than silently reducing them.
                if image.mode in {"I", "F", "I;16", "I;16B", "I;16L"}:
                    image.close()
                    continue
                whole = _docx_candidate(name, image, None, Image)
                if whole is None:
                    image.close()
                    continue
                candidates.append(whole)
                boxes = {_crop_box(image.size, rect) for rect in crops.get(name, ())}
                boxes.discard(None)
                for box in sorted(boxes):
                    cropped = _docx_candidate(name, image, box, Image)
                    if cropped is not None:
                        candidates.append(cropped)
    except (BadZipFile, KeyError, OSError):
        _close_candidates(candidates)
        return []
    return candidates


def _visual_similarity(
    source_image: Any,
    exported_image: Any,
    *,
    language: str = DEFAULT_LANGUAGE,
) -> float:
    try:
        from PIL import Image, ImageChops, ImageStat
    except ImportError as exc:
        raise ConversionError(tr(language, "pillow_restore_required")) from exc

    if not source_image.width or not source_image.height:
        return -1.0
    if not exported_image.width or not exported_image.height:
        return -1.0
    source_ratio = source_image.width / source_image.height
    exported_ratio = exported_image.width / exported_image.height
    if abs(source_ratio - exported_ratio) / max(source_ratio, exported_ratio) > 0.01:
        return -1.0

    source_sample = _comparison_sample(source_image, Image)
    exported_sample = _comparison_sample(exported_image, Image)
    try:
        return _sample_similarity(
            source_sample,
            exported_sample,
            ImageChops,
            ImageStat,
        )
    finally:
        source_sample.close()
        exported_sample.close()


def _best_source_image(
    exported_image: Any,
    candidates: list[_DocxImage],
    *,
    matte: tuple[int, int, int] | None = None,
    language: str = DEFAULT_LANGUAGE,
) -> _DocxImage | None:
    try:
        from PIL import Image, ImageChops, ImageStat
    except ImportError as exc:
        raise ConversionError(tr(language, "pillow_restore_required")) from exc

    if not exported_image.width or not exported_image.height:
        return None
    exported_ratio = exported_image.width / exported_image.height
    exported_sample = _comparison_sample(exported_image, Image)
    # A /Matte image stores colour already blended with the matte. Compare those
    # channels with the source composited onto the same matte; keep the plain
    # comparison too, in case a decoder un-premultiplies the colour itself.
    color_sample = _color_channel_sample(exported_image, Image) if matte else None
    best_score = -1.0
    second_score = -1.0
    best_candidate: _DocxImage | None = None
    second_candidate: _DocxImage | None = None
    try:
        for candidate in candidates:
            width, height = candidate.size
            candidate_ratio = width / height
            if abs(candidate_ratio - exported_ratio) / max(
                candidate_ratio,
                exported_ratio,
            ) > 0.01:
                continue
            score = _sample_similarity(
                candidate.comparison_sample,
                exported_sample,
                ImageChops,
                ImageStat,
            )
            if matte and color_sample is not None:
                score = max(
                    score,
                    _sample_similarity(
                        _matte_sample(candidate, matte, Image),
                        color_sample,
                        ImageChops,
                        ImageStat,
                    ),
                )
            if score > best_score:
                second_score = best_score
                second_candidate = best_candidate
                best_score = score
                best_candidate = candidate
            elif score > second_score:
                second_score = score
                second_candidate = candidate
    finally:
        exported_sample.close()
        if color_sample is not None:
            color_sample.close()

    if best_candidate is None or best_score < IMAGE_MATCH_THRESHOLD:
        return None
    if (
        second_candidate is not None
        and best_score - second_score < IMAGE_MATCH_AMBIGUITY_MARGIN
        and best_candidate.digest != second_candidate.digest
    ):
        return None
    return best_candidate


def _flate_image_object(image: Any, writer: Any) -> Any:
    from pypdf.generic import BooleanObject, DecodedStreamObject, NameObject, NumberObject

    bands = image.getbands()
    alpha = None
    if "A" in bands or "transparency" in image.info:
        rgba = image.convert("RGBA")
        color = rgba.convert("RGB")
        alpha = rgba.getchannel("A")
        rgba.close()
        color_space = "/DeviceRGB"
    elif image.mode == "CMYK":
        color = image.convert("CMYK")
        color_space = "/DeviceCMYK"
    elif image.mode in {"1", "L"}:
        color = image.convert("L")
        color_space = "/DeviceGray"
    else:
        color = image.convert("RGB")
        color_space = "/DeviceRGB"

    stream = DecodedStreamObject()
    stream.set_data(color.tobytes())
    stream.update(
        {
            NameObject("/Type"): NameObject("/XObject"),
            NameObject("/Subtype"): NameObject("/Image"),
            NameObject("/Width"): NumberObject(color.width),
            NameObject("/Height"): NumberObject(color.height),
            NameObject("/ColorSpace"): NameObject(color_space),
            NameObject("/BitsPerComponent"): NumberObject(8),
            # Keep viewer-side scaling crisp. The source pixels are already restored
            # at their native resolution, so requesting interpolation can make text
            # and screenshots look slightly blurred at common zoom levels.
            NameObject("/Interpolate"): BooleanObject(False),
        }
    )
    color.close()
    encoded = stream.flate_encode()

    if alpha is not None:
        mask = DecodedStreamObject()
        mask.set_data(alpha.tobytes())
        mask.update(
            {
                NameObject("/Type"): NameObject("/XObject"),
                NameObject("/Subtype"): NameObject("/Image"),
                NameObject("/Width"): NumberObject(alpha.width),
                NameObject("/Height"): NumberObject(alpha.height),
                NameObject("/ColorSpace"): NameObject("/DeviceGray"),
                NameObject("/BitsPerComponent"): NumberObject(8),
            }
        )
        alpha.close()
        encoded[NameObject("/SMask")] = writer._add_object(mask.flate_encode())
    return encoded


def restore_original_images(
    source_docx: str | Path,
    exported_pdf: str | Path,
    *,
    language: str = DEFAULT_LANGUAGE,
    cancel_requested: Callable[[], bool] | None = None,
) -> ImageRestorationReport:
    """Restore high-confidence DOCX source pixels without changing PDF page geometry.

    `cancel_requested` is polled between images and before the restored PDF
    replaces `exported_pdf`. On cancel, ConversionCancelled is raised and
    `exported_pdf` is left as Word wrote it.
    """
    source_path = validate_source(source_docx, language=language)
    pdf_path = Path(exported_pdf).expanduser().resolve()
    candidates = _load_docx_images(
        source_path,
        language=language,
        cancel_requested=cancel_requested,
    )
    if not candidates:
        return ImageRestorationReport(0, 0, 0)

    try:
        from pypdf import PdfWriter
    except ImportError as exc:
        _close_candidates(candidates)
        raise ConversionError(tr(language, "pypdf_restore_required")) from exc

    writer = None
    stage = pdf_path.with_name(f".docxpdf-quality-{uuid.uuid4().hex}.pdf")
    images_examined = 0
    images_restored = 0
    seen_references: set[tuple[int, int]] = set()
    try:
        writer = PdfWriter(clone_from=str(pdf_path))
        for page in writer.pages:
            for exported in page.images:
                if cancel_requested and cancel_requested():
                    raise ConversionCancelled()
                reference = exported.indirect_reference
                if reference is None:
                    continue
                reference_key = (reference.idnum, reference.generation)
                if reference_key in seen_references:
                    continue
                seen_references.add(reference_key)
                images_examined += 1
                match = _best_source_image(
                    exported.image,
                    candidates,
                    matte=_image_matte(reference.get_object()),
                    language=language,
                )
                if match is None:
                    continue
                pixels = _candidate_pixels(match)
                try:
                    replacement = _flate_image_object(pixels, writer)
                finally:
                    if pixels is not match.image:
                        pixels.close()
                replacement.indirect_reference = reference
                writer._objects[reference.idnum - 1] = replacement
                images_restored += 1

        if cancel_requested and cancel_requested():
            raise ConversionCancelled()
        if images_restored:
            with stage.open("wb") as stream:
                writer.write(stream)
            close = getattr(writer, "close", None)
            if close:
                close()
            writer = None
            _check_pdf(stage, language=language)
            if cancel_requested and cancel_requested():
                raise ConversionCancelled()
            _replace_with_retry(stage, pdf_path)
    except ConversionError:
        raise
    except Exception as exc:
        raise ConversionError(tr(language, "restore_failed", error=exc)) from exc
    finally:
        stage.unlink(missing_ok=True)
        if writer is not None:
            close = getattr(writer, "close", None)
            if close:
                close()
        _close_candidates(candidates)

    return ImageRestorationReport(
        images_examined=images_examined,
        images_restored=images_restored,
        images_unmatched=max(0, images_examined - images_restored),
    )


def _open_word_process(word: Any) -> Any:
    """Open a terminate handle to the WINWORD.EXE behind a DispatchEx instance.

    A temporary unique caption identifies this instance's hidden main window even
    when other lanes start Word at the same moment. Holding the handle pins the
    process identity, so a later abort can never hit a reused process id.
    Returns None when the process cannot be identified.
    """
    try:
        import win32api
        import win32con
        import win32gui
        import win32process
    except ImportError:
        return None
    try:
        previous_caption = word.Caption
    except Exception:
        return None
    marker = f"DocxPDF-{uuid.uuid4().hex}"
    try:
        word.Caption = marker
        window = win32gui.FindWindow("OpusApp", marker)
        if not window:
            return None
        _thread_id, process_id = win32process.GetWindowThreadProcessId(window)
        return win32api.OpenProcess(
            win32con.PROCESS_TERMINATE | win32con.SYNCHRONIZE,
            False,
            process_id,
        )
    except Exception:
        return None
    finally:
        try:
            word.Caption = previous_caption
        except Exception:
            pass


def _terminate_process(handle: Any) -> None:
    import win32api

    win32api.TerminateProcess(handle, 1)


class WordSession:
    """An isolated, hidden Word COM instance reusable across one batch."""

    def __init__(self, *, language: str = DEFAULT_LANGUAGE) -> None:
        self.language = normalize_language(language)
        self._pythoncom: ModuleType | None = None
        self._word: Any = None
        self._previous_security: Any = None
        self._com_initialized = False
        self._process_handle: Any = None
        # Guards the process handle and watchdog state shared with timer threads.
        self._process_lock = threading.Lock()
        self._watchdog_armed = False
        self._timed_out = False

    def __enter__(self) -> WordSession:
        self.start()
        return self

    def __exit__(self, _exc_type, _exc, _traceback) -> None:
        self.close()

    def start(self) -> None:
        if self._word is not None:
            return
        if platform.system() != "Windows":
            raise ConversionError(tr(self.language, "windows_only"), abort_batch=True)

        pythoncom, win32_client = _load_com_modules(self.language)
        self._pythoncom = pythoncom
        try:
            pythoncom.CoInitialize()
            self._com_initialized = True
            self._word = win32_client.DispatchEx("Word.Application")
            self._process_handle = _open_word_process(self._word)
            self._word.Visible = False
            self._word.DisplayAlerts = WD_ALERTS_NONE
            try:
                self._word.ScreenUpdating = False
            except Exception:
                pass
            try:
                self._previous_security = self._word.AutomationSecurity
                self._word.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
            except Exception:
                self._previous_security = None
        except ConversionError:
            self.close()
            raise
        except Exception as exc:
            self.close()
            raise _friendly_com_error(
                exc,
                during_startup=True,
                language=self.language,
            ) from exc

    def close(self) -> None:
        word = self._word
        self._word = None
        if word is not None:
            try:
                documents = word.Documents
                for index in range(int(documents.Count), 0, -1):
                    try:
                        documents.Item(index).Close(WD_DO_NOT_SAVE_CHANGES)
                    except Exception:
                        pass
            except Exception:
                pass
            if self._previous_security is not None:
                try:
                    word.AutomationSecurity = self._previous_security
                except Exception:
                    pass
            try:
                word.Quit(WD_DO_NOT_SAVE_CHANGES)
            except Exception:
                pass

        with self._process_lock:
            handle = self._process_handle
            self._process_handle = None
            self._watchdog_armed = False
        if handle is not None:
            try:
                handle.Close()
            except Exception:
                pass
        self._previous_security = None
        gc.collect()
        if self._com_initialized and self._pythoncom is not None:
            try:
                self._pythoncom.CoUninitialize()
            except Exception:
                pass
        self._com_initialized = False
        self._pythoncom = None

    def abort(self) -> bool:
        """Terminate this session's own WINWORD.EXE; safe to call from any thread.

        A COM call blocked by a hidden Word dialog never returns on its own.
        Ending the process makes that call fail so the lane can recover.
        """
        with self._process_lock:
            handle = self._process_handle
            if handle is None:
                return False
            try:
                _terminate_process(handle)
            except Exception:
                return False
        return True

    def _arm_watchdog(self, seconds: float) -> threading.Timer | None:
        with self._process_lock:
            self._timed_out = False
            if self._process_handle is None:
                return None
            self._watchdog_armed = True
        timer = threading.Timer(seconds, self._watchdog_expired)
        timer.daemon = True
        timer.start()
        return timer

    def _watchdog_expired(self) -> None:
        with self._process_lock:
            if not self._watchdog_armed:
                return
            self._watchdog_armed = False
            self._timed_out = True
        self.abort()

    def _disarm_watchdog(self, timer: threading.Timer | None) -> None:
        with self._process_lock:
            self._watchdog_armed = False
        if timer is not None:
            timer.cancel()

    def _timeout_error(self, seconds: float) -> ConversionError:
        return ConversionError(
            tr(self.language, "word_timed_out", seconds=round(seconds))
        )

    def convert_docx(
        self,
        source: str | Path,
        output: str | Path,
        *,
        overwrite: bool = False,
        timeout: float | None = None,
        cancel_requested: Callable[[], bool] | None = None,
    ) -> ConversionResult:
        """Export one DOCX with this session's Word, then restore its images.

        `cancel_requested` is polled after Word's export: between images during
        restoration and before publishing. Word's export itself is not
        interrupted. On cancel the exported PDF is deleted, nothing is published
        and ConversionCancelled is raised.
        """
        if self._word is None or self._pythoncom is None:
            raise ConversionError(
                tr(self.language, "session_not_started"),
                abort_batch=True,
            )

        source_path = validate_source(source, language=self.language)
        output_path = Path(output).expanduser().resolve()
        if output_path.suffix.lower() != ".pdf":
            raise ConversionError(tr(self.language, "output_pdf_extension"))
        if source_path == output_path:
            raise ConversionError(tr(self.language, "same_input_output"))
        ensure_output_dir(output_path.parent, language=self.language)
        if output_path.exists() and not overwrite:
            raise ConversionError(
                tr(self.language, "output_exists", name=output_path.name)
            )

        # Word is given a visible, unique filename in the final folder. This is
        # more reliable for OneDrive and controlled-folder access than a hidden file.
        temporary_output = output_path.with_name(f"DocxPDF-{uuid.uuid4().hex}.pdf")
        started = time.monotonic()
        document: Any = None
        # `timeout` also bounds Word's open + export: a hidden dialog would
        # otherwise block this lane, and with it cancel and close, forever.
        word_timeout = DEFAULT_WORD_TIMEOUT if timeout is None else max(float(timeout), 1.0)
        watchdog = self._arm_watchdog(word_timeout)
        try:
            missing = self._pythoncom.Missing
            document = self._word.Documents.Open(
                str(source_path),
                False,  # ConfirmConversions
                True,  # ReadOnly
                False,  # AddToRecentFiles
                missing,
                missing,
                False,  # Revert
                missing,
                missing,
                missing,
                missing,
                False,  # Visible
                False,  # OpenAndRepair
                missing,  # DocumentDirection
                True,  # NoEncodingDialog
            )
            export_args = (
                str(temporary_output),
                WD_EXPORT_FORMAT_PDF,
                False,  # OpenAfterExport
                WD_EXPORT_OPTIMIZE_FOR_PRINT,
                WD_EXPORT_ALL_DOCUMENT,
                1,
                1,
                WD_EXPORT_DOCUMENT_CONTENT,
                True,  # IncludeDocProps
                True,  # KeepIRM and sensitivity labels
                WD_EXPORT_CREATE_HEADING_BOOKMARKS,
                True,  # DocStructureTags
                True,  # BitmapMissingFonts preserves appearance
                False,  # Do not force PDF/A; it can introduce visual artifacts
            )
            try:
                high_quality_export = document.ExportAsFixedFormat2
            except Exception:
                high_quality_export = None
            if high_quality_export is not None:
                # The final True is OptimizeForImageQuality. Microsoft documents
                # this as retaining original image quality instead of downsampling.
                high_quality_export(*export_args, True)
            else:
                # Older Word builds expose the equivalent document switch but not
                # ExportAsFixedFormat2. Keep the change in memory only; never save it.
                try:
                    document.OptimizeForImageQuality = True
                except Exception as exc:
                    raise ConversionError(
                        tr(self.language, "word_quality_unsupported"),
                        abort_batch=True,
                    ) from exc
                document.ExportAsFixedFormat(*export_args)
            settle_timeout = 8.0 if timeout is None else min(max(float(timeout), 1.0), 15.0)
            if not _wait_for_pdf(temporary_output, settle_timeout):
                raise ConversionError(
                    tr(self.language, "word_pdf_not_finished"),
                    abort_batch=True,
                )
            size = _check_pdf(temporary_output, language=self.language)
        except ConversionError as exc:
            temporary_output.unlink(missing_ok=True)
            if self._timed_out:
                raise self._timeout_error(word_timeout) from exc
            raise
        except OSError as exc:
            temporary_output.unlink(missing_ok=True)
            if self._timed_out:
                raise self._timeout_error(word_timeout) from exc
            raise ConversionError(
                tr(self.language, "output_write_failed", error=exc)
            ) from exc
        except Exception as exc:
            temporary_output.unlink(missing_ok=True)
            if self._timed_out:
                raise self._timeout_error(word_timeout) from exc
            raise _friendly_com_error(exc, language=self.language) from exc
        finally:
            if document is not None:
                try:
                    document.Close(WD_DO_NOT_SAVE_CHANGES)
                except Exception:
                    pass
                document = None
                gc.collect()
            self._disarm_watchdog(watchdog)
        if self._timed_out:
            # The export returned just as the watchdog ended Word. The PDF may be
            # fine, but this session is gone; report it so the lane restarts Word.
            temporary_output.unlink(missing_ok=True)
            raise self._timeout_error(word_timeout)

        quality_report = ImageRestorationReport(0, 0, 0)
        try:
            quality_report = restore_original_images(
                source_path,
                temporary_output,
                language=self.language,
                cancel_requested=cancel_requested,
            )
            size = _check_pdf(temporary_output, language=self.language)
            if cancel_requested and cancel_requested():
                raise ConversionCancelled()
            output_path, output_relocated = _publish_pdf(
                temporary_output, output_path, language=self.language
            )
        except ConversionError:
            raise
        except OSError as exc:
            raise ConversionError(
                tr(self.language, "final_save_failed", error=exc)
            ) from exc
        finally:
            temporary_output.unlink(missing_ok=True)

        return ConversionResult(
            source=source_path,
            output=output_path,
            size_bytes=size,
            elapsed_seconds=time.monotonic() - started,
            images_restored=quality_report.images_restored,
            images_examined=quality_report.images_examined,
            output_relocated=output_relocated,
        )


def convert_docx(
    source: str | Path,
    output: str | Path,
    *,
    overwrite: bool = False,
    timeout: float | None = None,
    language: str = DEFAULT_LANGUAGE,
    cancel_requested: Callable[[], bool] | None = None,
) -> ConversionResult:
    """Export one DOCX through an isolated Word instance without modifying it."""
    with WordSession(language=language) as session:
        return session.convert_docx(
            source,
            output,
            overwrite=overwrite,
            timeout=timeout,
            cancel_requested=cancel_requested,
        )


def merge_pdfs(
    pdfs: list[str | Path] | tuple[str | Path, ...],
    output: str | Path,
    *,
    overwrite: bool = False,
    language: str = DEFAULT_LANGUAGE,
    cancel_requested: Callable[[], bool] | None = None,
    on_file_merged: Callable[[Path, int, int], None] | None = None,
) -> MergeResult:
    """Concatenate PDF pages without rasterizing or re-encoding page images."""
    if len(pdfs) < 2:
        raise ConversionError(tr(language, "merge_two_required"))
    output_path = Path(output).expanduser().resolve()
    if output_path.suffix.lower() != ".pdf":
        raise ConversionError(tr(language, "merge_pdf_extension"))
    ensure_output_dir(output_path.parent, language=language)
    if output_path.exists() and not overwrite:
        raise ConversionError(
            tr(language, "merge_output_exists", name=output_path.name)
        )

    try:
        from pypdf import PdfReader, PdfWriter
    except ImportError as exc:
        raise ConversionError(tr(language, "merge_pypdf_required")) from exc

    input_paths = tuple(Path(pdf).expanduser().resolve() for pdf in pdfs)
    if output_path in input_paths:
        raise ConversionError(tr(language, "merge_output_input"))
    for path in input_paths:
        if not path.is_file() or path.suffix.lower() != ".pdf":
            raise ConversionError(tr(language, "merge_input_missing", path=path))
        _check_pdf(path, language=language)

    # A regular staging name avoids problems some sync clients have with
    # hidden files in a OneDrive-controlled output folder.
    temporary_output = output_path.with_name(f"DocxPDF-merge-{uuid.uuid4().hex}.pdf")
    started = time.monotonic()
    writer = PdfWriter()
    readers = []
    page_count = 0
    try:
        for index, path in enumerate(input_paths, start=1):
            if cancel_requested and cancel_requested():
                raise MergeCancelled()
            reader = PdfReader(str(path), strict=False)
            readers.append(reader)
            if reader.is_encrypted:
                raise ConversionError(tr(language, "merge_encrypted", name=path.name))
            for page in reader.pages:
                if cancel_requested and cancel_requested():
                    raise MergeCancelled()
                writer.add_page(page)
                page_count += 1
            if on_file_merged:
                on_file_merged(path, index, len(input_paths))
        if cancel_requested and cancel_requested():
            raise MergeCancelled()
        with temporary_output.open("wb") as stream:
            writer.write(stream)
        size = _check_pdf(temporary_output, language=language)
        if cancel_requested and cancel_requested():
            raise MergeCancelled()
        output_path, output_relocated = _publish_pdf(
            temporary_output, output_path, language=language
        )
    except ConversionError:
        raise
    except Exception as exc:
        raise ConversionError(tr(language, "merge_failed", error=exc)) from exc
    finally:
        temporary_output.unlink(missing_ok=True)
        for reader in readers:
            close = getattr(reader, "close", None)
            if close:
                close()

    return MergeResult(
        inputs=input_paths,
        output=output_path,
        page_count=page_count,
        size_bytes=size,
        elapsed_seconds=time.monotonic() - started,
        output_relocated=output_relocated,
    )
