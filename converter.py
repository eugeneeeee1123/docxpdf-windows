from __future__ import annotations

import gc
import hashlib
import os
import platform
import shutil
import time
import uuid
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from types import ModuleType
from typing import Any
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


class ConversionError(RuntimeError):
    """A user-facing conversion failure."""

    def __init__(self, message: str, *, abort_batch: bool = False) -> None:
        super().__init__(message)
        self.abort_batch = abort_batch


@dataclass(frozen=True)
class ConversionResult:
    source: Path
    output: Path
    size_bytes: int
    elapsed_seconds: float
    images_restored: int = 0
    images_examined: int = 0


@dataclass(frozen=True)
class MergeResult:
    inputs: tuple[Path, ...]
    output: Path
    page_count: int
    size_bytes: int
    elapsed_seconds: float


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
) -> Path:
    source_path = Path(source)
    folder = Path(output_dir).expanduser()
    target = folder / f"{source_path.stem}.pdf"
    if overwrite or not target.exists():
        return target

    for number in range(1, 10_000):
        candidate = folder / f"{source_path.stem}-{number}.pdf"
        if not candidate.exists():
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


def _flatten_for_comparison(image: Any, image_module: Any) -> Any:
    if "A" in image.getbands() or "transparency" in image.info:
        rgba = image.convert("RGBA")
        background = image_module.new("RGBA", rgba.size, (255, 255, 255, 255))
        return image_module.alpha_composite(background, rgba).convert("RGB")
    return image.convert("RGB")


def _load_docx_images(
    source_path: Path,
    *,
    language: str = DEFAULT_LANGUAGE,
) -> list[_DocxImage]:
    try:
        from PIL import Image, ImageOps
    except ImportError as exc:
        raise ConversionError(tr(language, "pillow_restore_install")) from exc

    candidates: list[_DocxImage] = []
    try:
        with ZipFile(source_path) as archive:
            media_names = sorted(
                name
                for name in archive.namelist()
                if name.lower().startswith("word/media/") and not name.endswith("/")
            )
            for name in media_names:
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
                flattened = _flatten_for_comparison(image, Image)
                digest = hashlib.sha256(flattened.tobytes()).hexdigest()
                flattened.close()
                candidates.append(_DocxImage(name=name, image=image, digest=digest))
    except (BadZipFile, KeyError, OSError):
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

    sample_size = (64, 64)
    source_sample = _flatten_for_comparison(source_image, Image).resize(
        sample_size, Image.Resampling.LANCZOS
    )
    exported_sample = _flatten_for_comparison(exported_image, Image).resize(
        sample_size, Image.Resampling.LANCZOS
    )
    try:
        difference = ImageChops.difference(source_sample, exported_sample)
        rms = ImageStat.Stat(difference).rms
        normalized = (sum(value * value for value in rms) / len(rms)) ** 0.5 / 255.0
        return max(0.0, 1.0 - normalized)
    finally:
        source_sample.close()
        exported_sample.close()


def _best_source_image(
    exported_image: Any,
    candidates: list[_DocxImage],
    *,
    language: str = DEFAULT_LANGUAGE,
) -> _DocxImage | None:
    scored = sorted(
        (
            (
                _visual_similarity(
                    candidate.image,
                    exported_image,
                    language=language,
                ),
                candidate,
            )
            for candidate in candidates
        ),
        key=lambda item: item[0],
        reverse=True,
    )
    if not scored or scored[0][0] < 0.985:
        return None
    if (
        len(scored) > 1
        and scored[0][0] - scored[1][0] < 0.002
        and scored[0][1].digest != scored[1][1].digest
    ):
        return None
    return scored[0][1]


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
            NameObject("/Interpolate"): BooleanObject(True),
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
) -> ImageRestorationReport:
    """Restore high-confidence DOCX source pixels without changing PDF page geometry."""
    source_path = validate_source(source_docx, language=language)
    pdf_path = Path(exported_pdf).expanduser().resolve()
    candidates = _load_docx_images(source_path, language=language)
    if not candidates:
        return ImageRestorationReport(0, 0, 0)

    try:
        from pypdf import PdfWriter
    except ImportError as exc:
        for candidate in candidates:
            candidate.image.close()
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
                    language=language,
                )
                if match is None:
                    continue
                replacement = _flate_image_object(match.image, writer)
                replacement.indirect_reference = reference
                writer._objects[reference.idnum - 1] = replacement
                images_restored += 1

        if images_restored:
            with stage.open("wb") as stream:
                writer.write(stream)
            close = getattr(writer, "close", None)
            if close:
                close()
            writer = None
            _check_pdf(stage, language=language)
            os.replace(stage, pdf_path)
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
        for candidate in candidates:
            candidate.image.close()

    return ImageRestorationReport(
        images_examined=images_examined,
        images_restored=images_restored,
        images_unmatched=max(0, images_examined - images_restored),
    )


class WordSession:
    """An isolated, hidden Word COM instance reusable across one batch."""

    def __init__(self, *, language: str = DEFAULT_LANGUAGE) -> None:
        self.language = normalize_language(language)
        self._pythoncom: ModuleType | None = None
        self._word: Any = None
        self._previous_security: Any = None
        self._com_initialized = False

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

        self._previous_security = None
        gc.collect()
        if self._com_initialized and self._pythoncom is not None:
            try:
                self._pythoncom.CoUninitialize()
            except Exception:
                pass
        self._com_initialized = False
        self._pythoncom = None

    def convert_docx(
        self,
        source: str | Path,
        output: str | Path,
        *,
        overwrite: bool = False,
        timeout: float | None = None,
    ) -> ConversionResult:
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
        except ConversionError:
            temporary_output.unlink(missing_ok=True)
            raise
        except OSError as exc:
            temporary_output.unlink(missing_ok=True)
            raise ConversionError(
                tr(self.language, "output_write_failed", error=exc)
            ) from exc
        except Exception as exc:
            temporary_output.unlink(missing_ok=True)
            raise _friendly_com_error(exc, language=self.language) from exc
        finally:
            if document is not None:
                try:
                    document.Close(WD_DO_NOT_SAVE_CHANGES)
                except Exception:
                    pass
                document = None
                gc.collect()

        quality_report = ImageRestorationReport(0, 0, 0)
        try:
            quality_report = restore_original_images(
                source_path,
                temporary_output,
                language=self.language,
            )
            size = _check_pdf(temporary_output, language=self.language)
            os.replace(temporary_output, output_path)
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
        )


def convert_docx(
    source: str | Path,
    output: str | Path,
    *,
    overwrite: bool = False,
    timeout: float | None = None,
    language: str = DEFAULT_LANGUAGE,
) -> ConversionResult:
    """Export one DOCX through an isolated Word instance without modifying it."""
    with WordSession(language=language) as session:
        return session.convert_docx(
            source,
            output,
            overwrite=overwrite,
            timeout=timeout,
        )


def merge_pdfs(
    pdfs: list[str | Path] | tuple[str | Path, ...],
    output: str | Path,
    *,
    overwrite: bool = False,
    language: str = DEFAULT_LANGUAGE,
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
    for path in input_paths:
        if not path.is_file() or path.suffix.lower() != ".pdf":
            raise ConversionError(tr(language, "merge_input_missing", path=path))
        _check_pdf(path, language=language)

    temporary_output = output_path.with_name(f".docxpdf-stage-{uuid.uuid4().hex}.pdf")
    started = time.monotonic()
    writer = PdfWriter()
    readers = []
    page_count = 0
    try:
        for path in input_paths:
            reader = PdfReader(str(path), strict=False)
            if reader.is_encrypted:
                raise ConversionError(tr(language, "merge_encrypted", name=path.name))
            readers.append(reader)
            for page in reader.pages:
                writer.add_page(page)
                page_count += 1
        with temporary_output.open("wb") as stream:
            writer.write(stream)
        size = _check_pdf(temporary_output, language=language)
        os.replace(temporary_output, output_path)
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
    )
