from __future__ import annotations

import errno
import os
import tempfile
import unittest
from io import BytesIO
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile

import converter


class FakePythonCom:
    Missing = object()

    def __init__(self) -> None:
        self.initialized = 0
        self.uninitialized = 0

    def CoInitialize(self) -> None:
        self.initialized += 1

    def CoUninitialize(self) -> None:
        self.uninitialized += 1


class FakeDocument:
    def __init__(self, *, export_error: Exception | None = None) -> None:
        self.export_error = export_error
        self.export_args = None
        self.export_method = None
        self.close_calls: list[int] = []

    def ExportAsFixedFormat2(self, *args) -> None:
        self.export_method = "ExportAsFixedFormat2"
        self.export_args = args
        if self.export_error:
            raise self.export_error
        Path(args[0]).write_bytes(b"%PDF-1.7\nmock")

    def ExportAsFixedFormat(self, *args) -> None:
        self.export_method = "ExportAsFixedFormat"
        self.export_args = args
        if self.export_error:
            raise self.export_error
        Path(args[0]).write_bytes(b"%PDF-1.7\nmock")

    def Close(self, save_changes: int) -> None:
        self.close_calls.append(save_changes)


class FakeDocuments:
    def __init__(self, document: FakeDocument) -> None:
        self.document = document
        self.open_args = None

    @property
    def Count(self) -> int:
        return 0

    def Open(self, *args):
        self.open_args = args
        return self.document

    def Item(self, _index: int):
        return self.document


class FakeWord:
    def __init__(self, document: FakeDocument) -> None:
        self.Documents = FakeDocuments(document)
        self.Visible = True
        self.DisplayAlerts = 1
        self.ScreenUpdating = True
        self.AutomationSecurity = 1
        self.quit_calls: list[int] = []

    def Quit(self, save_changes: int) -> None:
        self.quit_calls.append(save_changes)


def write_image_pdf(path: Path, color, alpha=None, matte=None) -> None:
    """Write a one-page PDF shaped like Word's export of a single picture."""
    from pypdf import PdfWriter
    from pypdf.generic import (
        ArrayObject,
        DecodedStreamObject,
        DictionaryObject,
        FloatObject,
        NameObject,
        NumberObject,
    )

    def image_stream(image, color_space: str):
        stream = DecodedStreamObject()
        stream.set_data(image.tobytes())
        stream.update(
            {
                NameObject("/Type"): NameObject("/XObject"),
                NameObject("/Subtype"): NameObject("/Image"),
                NameObject("/Width"): NumberObject(image.width),
                NameObject("/Height"): NumberObject(image.height),
                NameObject("/ColorSpace"): NameObject(color_space),
                NameObject("/BitsPerComponent"): NumberObject(8),
            }
        )
        return stream

    writer = PdfWriter()
    page = writer.add_blank_page(width=400, height=300)
    image = image_stream(color.convert("RGB"), "/DeviceRGB").flate_encode()
    if alpha is not None:
        mask = image_stream(alpha.convert("L"), "/DeviceGray")
        if matte is not None:
            mask[NameObject("/Matte")] = ArrayObject(
                FloatObject(value / 255) for value in matte
            )
        image[NameObject("/SMask")] = writer._add_object(mask.flate_encode())
    page[NameObject("/Resources")] = DictionaryObject(
        {
            NameObject("/XObject"): DictionaryObject(
                {NameObject("/Im0"): writer._add_object(image)}
            )
        }
    )
    content = DecodedStreamObject()
    content.set_data(b"q 400 0 0 300 0 0 cm /Im0 Do Q")
    page[NameObject("/Contents")] = writer._add_object(content)
    with path.open("wb") as stream:
        writer.write(stream)


def png_bytes(image) -> bytes:
    buffer = BytesIO()
    image.save(buffer, "PNG")
    return buffer.getvalue()


RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DOCUMENT_NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
)


def picture_part(root: str, pictures: list[tuple[str, str]]) -> str:
    """A minimal document/header part with (relationship id, srcRect attrs) pictures."""
    blips = "".join(
        f'<pic:pic><pic:blipFill><a:blip r:embed="{rel_id}"/>'
        f"<a:srcRect {crop}/><a:stretch><a:fillRect/></a:stretch>"
        "</pic:blipFill></pic:pic>"
        for rel_id, crop in pictures
    )
    return f"<w:{root} {DOCUMENT_NS}><w:body>{blips}</w:body></w:{root}>"


def relationships_part(targets: dict[str, str], external: set[str] = frozenset()) -> str:
    rows = "".join(
        f'<Relationship Id="{rel_id}" Type="image" Target="{target}"'
        + (' TargetMode="External"' if rel_id in external else "")
        + "/>"
        for rel_id, target in targets.items()
    )
    return f'<Relationships xmlns="{RELS_NS}">{rows}</Relationships>'


class FakeClient:
    def __init__(self, word: FakeWord) -> None:
        self.word = word
        self.dispatch_calls: list[str] = []

    def DispatchEx(self, prog_id: str) -> FakeWord:
        self.dispatch_calls.append(prog_id)
        return self.word


class ConverterTests(unittest.TestCase):
    def test_replace_with_retry_handles_transient_access_denied(self) -> None:
        source = Path("source.pdf")
        destination = Path("destination.pdf")
        denied = PermissionError(errno.EACCES, "Access is denied")
        denied.winerror = 5
        with patch.object(converter.os, "replace", side_effect=[denied, None]) as replace:
            with patch.object(converter.time, "sleep") as sleep:
                converter._replace_with_retry(source, destination, timeout=1)

        self.assertEqual(replace.call_count, 2)
        sleep.assert_called_once()

    def test_next_output_path_is_non_destructive(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            source = folder / "说明.docx"
            source.touch()
            (folder / "说明.pdf").write_bytes(b"existing")
            self.assertEqual(converter.next_output_path(source, folder).name, "说明-1.pdf")
            self.assertEqual(converter.next_merge_path(folder).name, "合并结果.pdf")

    @unittest.skipUnless(os.name == "nt", "Requires Windows file sharing semantics")
    def test_publish_pdf_uses_safe_name_when_existing_output_stays_locked(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            source = folder / "DocxPDF-merge-stage.pdf"
            destination = folder / "Merged.pdf"
            source.write_bytes(b"%PDF-1.7\nstaged")
            destination.write_bytes(b"%PDF-1.7\nopen")
            # A real open Windows file handle denies replacement. This exposes
            # WinError 5, which a mock containing only errno.EACCES misses.
            with destination.open("rb"):
                published, relocated = converter._publish_pdf(
                    source, destination, language="en", retry_timeout=0
                )

            self.assertEqual(published, folder / "Merged-1.pdf")
            self.assertTrue(relocated)
            self.assertEqual(destination.read_bytes(), b"%PDF-1.7\nopen")
            self.assertEqual(published.read_bytes(), b"%PDF-1.7\nstaged")
            self.assertFalse(source.exists())

    def test_english_errors_and_merge_filename(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            with self.assertRaises(converter.ConversionError) as raised:
                converter.validate_source(folder / "missing.docx", language="en")
            self.assertIn("File not found", str(raised.exception))
            self.assertEqual(
                converter.next_merge_path(folder, language="en").name,
                "Merged.pdf",
            )
            busy = converter._friendly_com_error(
                RuntimeError("Call was rejected by callee"),
                language="en",
            )
            self.assertIn("Microsoft Word is busy", str(busy))
            self.assertTrue(busy.abort_batch)

    def test_locate_word_app_accepts_file_or_folder_override(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            executable = folder / "WINWORD.EXE"
            executable.touch()
            with patch.dict(os.environ, {"DOCX_PDF_WORD_APP": str(folder)}):
                self.assertEqual(converter.locate_word_app(), executable.resolve())
            with patch.dict(os.environ, {"DOCX_PDF_WORD_APP": str(executable)}):
                self.assertEqual(converter.locate_word_app(), executable.resolve())

    def test_word_session_uses_read_only_print_quality_export_and_is_atomic(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            source = folder / "带空格 与 中文.docx"
            source.write_bytes(b"mock docx")
            output = folder / "result.pdf"
            document = FakeDocument()
            word = FakeWord(document)
            client = FakeClient(word)
            pythoncom = FakePythonCom()

            with patch.object(converter.platform, "system", return_value="Windows"):
                with patch.object(
                    converter,
                    "_load_com_modules",
                    return_value=(pythoncom, client),
                ):
                    with patch.object(converter, "_wait_for_pdf", return_value=True):
                        with converter.WordSession() as session:
                            result = session.convert_docx(source, output)

            self.assertEqual(client.dispatch_calls, ["Word.Application"])
            self.assertEqual(pythoncom.initialized, 1)
            self.assertEqual(pythoncom.uninitialized, 1)
            self.assertFalse(word.Visible)
            self.assertEqual(word.DisplayAlerts, converter.WD_ALERTS_NONE)
            self.assertEqual(word.AutomationSecurity, 1)
            self.assertEqual(word.quit_calls, [converter.WD_DO_NOT_SAVE_CHANGES])

            open_args = word.Documents.open_args
            self.assertEqual(open_args[0], str(source.resolve()))
            self.assertFalse(open_args[1])  # ConfirmConversions
            self.assertTrue(open_args[2])  # ReadOnly
            self.assertFalse(open_args[3])  # AddToRecentFiles
            self.assertFalse(open_args[11])  # Visible
            self.assertFalse(open_args[12])  # OpenAndRepair
            self.assertIs(open_args[13], pythoncom.Missing)  # DocumentDirection
            self.assertTrue(open_args[14])  # NoEncodingDialog

            export_args = document.export_args
            self.assertEqual(document.export_method, "ExportAsFixedFormat2")
            self.assertTrue(Path(export_args[0]).name.startswith("DocxPDF-"))
            self.assertEqual(export_args[1], converter.WD_EXPORT_FORMAT_PDF)
            self.assertEqual(export_args[3], converter.WD_EXPORT_OPTIMIZE_FOR_PRINT)
            self.assertEqual(export_args[4], converter.WD_EXPORT_ALL_DOCUMENT)
            self.assertEqual(export_args[7], converter.WD_EXPORT_DOCUMENT_CONTENT)
            self.assertTrue(export_args[12])  # BitmapMissingFonts
            self.assertFalse(export_args[13])  # PDF/A disabled
            self.assertTrue(export_args[14])  # OptimizeForImageQuality
            self.assertEqual(document.close_calls, [converter.WD_DO_NOT_SAVE_CHANGES])
            self.assertEqual(result.output, output.resolve())
            self.assertEqual(output.read_bytes(), b"%PDF-1.7\nmock")
            self.assertFalse(any(path.name.startswith("DocxPDF-") for path in folder.iterdir()))

    def test_com_failure_closes_document_and_aborts_batch_when_word_disconnects(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            source = folder / "source.docx"
            source.write_bytes(b"mock")
            output = folder / "output.pdf"
            document = FakeDocument(export_error=RuntimeError("Call was rejected by callee"))
            word = FakeWord(document)
            pythoncom = FakePythonCom()

            with patch.object(converter.platform, "system", return_value="Windows"):
                with patch.object(
                    converter,
                    "_load_com_modules",
                    return_value=(pythoncom, FakeClient(word)),
                ):
                    with converter.WordSession() as session:
                        with self.assertRaises(converter.ConversionError) as raised:
                            session.convert_docx(source, output)

            self.assertTrue(raised.exception.abort_batch)
            self.assertIn("Word 正忙", str(raised.exception))
            self.assertEqual(document.close_calls, [converter.WD_DO_NOT_SAVE_CHANGES])
            self.assertFalse(output.exists())

    def test_watchdog_closes_word_hung_during_export_and_keeps_batch_going(self) -> None:
        import threading
        import time

        class FakeProcessHandle:
            closed = False

            def Close(self) -> None:
                self.closed = True

        class HungDocument(FakeDocument):
            def __init__(self, killed: threading.Event) -> None:
                super().__init__()
                self.killed = killed

            def ExportAsFixedFormat2(self, *args) -> None:
                # Like a hidden Word dialog: the call only returns once the
                # process is gone, and then fails as a broken RPC call.
                Path(args[0]).write_bytes(b"%PDF-partial")
                if not self.killed.wait(10):
                    raise AssertionError("watchdog did not end Word")
                raise RuntimeError("The remote procedure call failed.")

        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            source = folder / "source.docx"
            source.write_bytes(b"mock")
            output = folder / "output.pdf"
            killed = threading.Event()
            handle = FakeProcessHandle()
            terminated = []

            def terminate(process_handle) -> None:
                terminated.append(process_handle)
                killed.set()

            document = HungDocument(killed)
            with patch.object(converter.platform, "system", return_value="Windows"):
                with patch.object(
                    converter,
                    "_load_com_modules",
                    return_value=(FakePythonCom(), FakeClient(FakeWord(document))),
                ):
                    with patch.object(converter, "_open_word_process", return_value=handle):
                        with patch.object(converter, "_terminate_process", side_effect=terminate):
                            with converter.WordSession(language="en") as session:
                                started = time.monotonic()
                                with self.assertRaises(converter.ConversionError) as raised:
                                    session.convert_docx(source, output, timeout=1)
                                elapsed = time.monotonic() - started

            self.assertEqual(terminated, [handle])
            self.assertLess(elapsed, 5)
            self.assertIn("did not finish the export within 1 seconds", str(raised.exception))
            # Only this document failed; the lane restarts Word for the next one.
            self.assertFalse(raised.exception.abort_batch)
            self.assertEqual(document.close_calls, [converter.WD_DO_NOT_SAVE_CHANGES])
            self.assertTrue(handle.closed)
            self.assertFalse(output.exists())
            self.assertFalse(list(folder.glob("DocxPDF-*")))

    def test_watchdog_leaves_word_running_when_export_finishes_in_time(self) -> None:
        import time

        class FakeProcessHandle:
            closed = False

            def Close(self) -> None:
                self.closed = True

        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            source = folder / "source.docx"
            source.write_bytes(b"mock")
            handle = FakeProcessHandle()
            with patch.object(converter.platform, "system", return_value="Windows"):
                with patch.object(
                    converter,
                    "_load_com_modules",
                    return_value=(FakePythonCom(), FakeClient(FakeWord(FakeDocument()))),
                ):
                    with patch.object(converter, "_open_word_process", return_value=handle):
                        with patch.object(converter, "_terminate_process") as terminate:
                            with patch.object(converter, "_wait_for_pdf", return_value=True):
                                with converter.WordSession() as session:
                                    first = session.convert_docx(source, folder / "one.pdf", timeout=1)
                                    second = session.convert_docx(source, folder / "two.pdf", timeout=1)
                                    time.sleep(1.3)  # past the watchdog deadline
                                    self.assertFalse(handle.closed)

            terminate.assert_not_called()
            self.assertTrue(handle.closed)
            self.assertEqual(first.output.read_bytes(), b"%PDF-1.7\nmock")
            self.assertEqual(second.output.read_bytes(), b"%PDF-1.7\nmock")

    def test_non_windows_session_is_rejected(self) -> None:
        with patch.object(converter.platform, "system", return_value="Darwin"):
            with self.assertRaises(converter.ConversionError) as raised:
                converter.WordSession().start()
        self.assertTrue(raised.exception.abort_batch)
        self.assertIn("Windows", str(raised.exception))

    def test_merge_pdfs_keeps_page_order(self) -> None:
        try:
            from pypdf import PdfReader, PdfWriter
        except ImportError:
            self.skipTest("pypdf is not installed")

        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            pdfs = []
            for index, count in enumerate((1, 2), start=1):
                path = folder / f"part-{index}.pdf"
                writer = PdfWriter()
                for _ in range(count):
                    writer.add_blank_page(width=595, height=842)
                with path.open("wb") as stream:
                    writer.write(stream)
                pdfs.append(path)

            result = converter.merge_pdfs(pdfs, folder / "合并结果.pdf")
            self.assertEqual(result.page_count, 3)
            self.assertEqual(len(PdfReader(str(result.output)).pages), 3)
            self.assertTrue(result.output.read_bytes().startswith(b"%PDF-"))

    def test_restore_original_images_preserves_page_geometry_and_source_pixels(self) -> None:
        try:
            from PIL import Image, ImageChops, ImageDraw
            from pypdf import PdfReader
        except ImportError:
            self.skipTest("pypdf and Pillow are required")

        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            source_docx = folder / "source.docx"
            exported_pdf = folder / "exported.pdf"

            original = Image.new("RGB", (320, 180), (232, 240, 250))
            drawing = ImageDraw.Draw(original)
            drawing.rectangle((15, 15, 305, 165), outline=(20, 70, 120), width=5)
            drawing.line((20, 140, 300, 35), fill=(190, 45, 45), width=4)
            image_bytes = BytesIO()
            original.save(image_bytes, "PNG")
            with ZipFile(source_docx, "w") as archive:
                archive.writestr("word/media/image1.png", image_bytes.getvalue())

            downsampled = original.resize((160, 90), Image.Resampling.LANCZOS)
            # Keep the synthetic export visually faithful, like Word's Flate-encoded
            # downsampled image, so this test exercises resolution restoration rather
            # than JPEG chroma-subsampling artifacts.
            downsampled.save(exported_pdf, "PDF", quality=95, subsampling=0)
            before_reader = PdfReader(str(exported_pdf))
            before_content = before_reader.pages[0].get_contents().get_data()
            before_box = tuple(before_reader.pages[0].mediabox)

            report = converter.restore_original_images(source_docx, exported_pdf)

            after_reader = PdfReader(str(exported_pdf))
            restored_item = after_reader.pages[0].images[0]
            restored = restored_item.image.convert("RGB")
            self.assertEqual(report.images_examined, 1)
            self.assertEqual(report.images_restored, 1)
            self.assertEqual(restored.size, original.size)
            self.assertIsNone(ImageChops.difference(original, restored).getbbox())
            self.assertFalse(
                restored_item.indirect_reference.get_object()["/Interpolate"].value
            )
            self.assertEqual(after_reader.pages[0].get_contents().get_data(), before_content)
            self.assertEqual(tuple(after_reader.pages[0].mediabox), before_box)

    def test_restore_semi_transparent_image_exported_with_matte(self) -> None:
        try:
            from PIL import Image, ImageChops, ImageDraw
            from pypdf import PdfReader
        except ImportError:
            self.skipTest("pypdf and Pillow are required")

        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            source_docx = folder / "source.docx"
            exported_pdf = folder / "exported.pdf"

            size = (320, 160)
            horizontal = Image.linear_gradient("L").rotate(90).resize(size)
            vertical = Image.linear_gradient("L").resize(size)
            original = Image.merge(
                "RGB", (horizontal, vertical, Image.new("L", size, 210))
            ).convert("RGBA")
            drawing = ImageDraw.Draw(original)
            drawing.ellipse((30, 20, 150, 140), fill=(230, 40, 40, 255))
            drawing.rectangle((180, 30, 300, 130), fill=(250, 250, 40, 160))
            alpha = ImageChops.multiply(
                original.getchannel("A"), horizontal.point(lambda v: max(v, 8))
            )
            original.putalpha(alpha)
            with ZipFile(source_docx, "w") as archive:
                archive.writestr("word/media/image1.png", png_bytes(original))

            # Word pre-blends the colour with the /Matte (black) and downsamples.
            black = Image.new("RGBA", size, (0, 0, 0, 255))
            preblended = Image.alpha_composite(black, original).convert("RGB")
            half = (size[0] // 2, size[1] // 2)
            write_image_pdf(
                exported_pdf,
                preblended.resize(half, Image.Resampling.LANCZOS),
                alpha.resize(half, Image.Resampling.LANCZOS),
                matte=(0, 0, 0),
            )

            exported = PdfReader(str(exported_pdf)).pages[0].images[0].image
            candidates = converter._load_docx_images(source_docx)
            try:
                # Without Matte awareness this picture stays at Word's resolution.
                self.assertIsNone(converter._best_source_image(exported, candidates))
                self.assertIs(
                    converter._best_source_image(exported, candidates, matte=(0, 0, 0)),
                    candidates[0],
                )
                self.assertIn((0, 0, 0), candidates[0].matte_samples)
            finally:
                converter._close_candidates(candidates)

            report = converter.restore_original_images(source_docx, exported_pdf)

            self.assertEqual((report.images_examined, report.images_restored), (1, 1))
            restored_item = PdfReader(str(exported_pdf)).pages[0].images[0]
            restored = restored_item.image.convert("RGBA")
            self.assertEqual(restored.size, original.size)
            self.assertIsNone(ImageChops.difference(original, restored).getbbox())
            soft_mask = restored_item.indirect_reference.get_object()["/SMask"].get_object()
            self.assertNotIn("/Matte", soft_mask)

    def test_restore_picture_cropped_in_word_uses_native_crop_region(self) -> None:
        try:
            from PIL import Image, ImageChops, ImageDraw
            from pypdf import PdfReader
        except ImportError:
            self.skipTest("pypdf and Pillow are required")

        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            source_docx = folder / "source.docx"
            exported_pdf = folder / "exported.pdf"

            original = Image.new("RGB", (400, 240), "white")
            drawing = ImageDraw.Draw(original)
            for y in range(0, 240, 12):
                drawing.text((4, y), f"Cropped line {y:03d} 0123456789", fill=(0, 0, 0))
            drawing.rectangle((260, 40, 380, 200), fill=(30, 110, 200))
            with ZipFile(source_docx, "w") as archive:
                archive.writestr("word/media/image1.png", png_bytes(original))
                archive.writestr(
                    "word/document.xml",
                    picture_part("document", [("rId7", 'l="10000" t="5000" r="25000" b="20000"')]),
                )
                archive.writestr(
                    "word/_rels/document.xml.rels",
                    relationships_part({"rId7": "media/image1.png"}),
                )

            crop_box = (40, 12, 300, 192)
            crop = original.crop(crop_box)
            # Word exports only the cropped region, downsampled.
            write_image_pdf(exported_pdf, crop.resize((130, 90), Image.Resampling.LANCZOS))

            report = converter.restore_original_images(source_docx, exported_pdf)

            self.assertEqual((report.images_examined, report.images_restored), (1, 1))
            restored = PdfReader(str(exported_pdf)).pages[0].images[0].image.convert("RGB")
            self.assertEqual(restored.size, crop.size)
            self.assertIsNone(ImageChops.difference(crop, restored).getbbox())

    def test_restore_cancel_stops_between_images_and_keeps_word_export(self) -> None:
        try:
            from PIL import Image, ImageDraw
            import pypdf  # noqa: F401
        except ImportError:
            self.skipTest("pypdf and Pillow are required")
        from threading import Event

        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            source_docx = folder / "source.docx"
            exported_pdf = folder / "exported.pdf"
            originals = []
            for index, color in enumerate(((200, 40, 40), (40, 90, 200))):
                original = Image.new("RGB", (320, 180), (240, 240, 235))
                ImageDraw.Draw(original).rectangle(
                    (20 + 60 * index, 20, 200 + 60 * index, 150), fill=color
                )
                originals.append(original)
            with ZipFile(source_docx, "w") as archive:
                for index, original in enumerate(originals, start=1):
                    archive.writestr(f"word/media/image{index}.png", png_bytes(original))
            # Two pages with one downsampled picture each, like Word's export.
            exported = [image.resize((160, 90), Image.Resampling.LANCZOS) for image in originals]
            exported[0].save(
                exported_pdf,
                "PDF",
                save_all=True,
                append_images=exported[1:],
                quality=95,
                subsampling=0,
            )
            word_export = exported_pdf.read_bytes()
            matched = []
            best_source_image = converter._best_source_image

            def count_matches(*args, **kwargs):
                matched.append(args[0].size)
                return best_source_image(*args, **kwargs)

            # 0: while loading the DOCX media; 1: before the second PDF image;
            # 2: after the last image, before the restored PDF is written.
            for cancel_after in (0, 1, 2):
                matched.clear()
                with self.subTest(cancel_after=cancel_after):
                    with patch.object(converter, "_best_source_image", side_effect=count_matches):
                        with self.assertRaises(converter.ConversionCancelled):
                            converter.restore_original_images(
                                source_docx,
                                exported_pdf,
                                cancel_requested=lambda: len(matched) >= cancel_after,
                            )
                    self.assertEqual(len(matched), cancel_after)
                    self.assertEqual(exported_pdf.read_bytes(), word_export)
                    self.assertEqual(
                        sorted(path.name for path in folder.iterdir()),
                        ["exported.pdf", "source.docx"],
                    )

            # Cancel after the restored PDF was staged, just before it replaces
            # Word's export.
            staged = Event()
            check_pdf = converter._check_pdf

            def cancel_after_staged_file(path, **kwargs):
                size = check_pdf(path, **kwargs)
                if path.name.startswith(".docxpdf-quality-"):
                    staged.set()
                return size

            with patch.object(converter, "_check_pdf", side_effect=cancel_after_staged_file):
                with self.assertRaises(converter.ConversionCancelled):
                    converter.restore_original_images(
                        source_docx, exported_pdf, cancel_requested=staged.is_set
                    )
            self.assertTrue(staged.is_set())
            self.assertEqual(exported_pdf.read_bytes(), word_export)
            self.assertFalse(list(folder.glob(".docxpdf-quality-*")))

            report = converter.restore_original_images(
                source_docx, exported_pdf, cancel_requested=lambda: False
            )
            self.assertEqual((report.images_examined, report.images_restored), (2, 2))
            self.assertNotEqual(exported_pdf.read_bytes(), word_export)

    def test_word_session_cancel_after_export_discards_pdf_and_publishes_nothing(self) -> None:
        from threading import Event

        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            source = folder / "source.docx"
            source.write_bytes(b"mock")
            output = folder / "result.pdf"
            for cancel_at in ("during restoration", "before publishing", None):
                cancelled = Event()
                restore_calls = []

                def restore(source_docx, exported_pdf, *, language, cancel_requested):
                    restore_calls.append(cancel_requested)
                    # Word's export is staged in the output folder meanwhile.
                    self.assertTrue(Path(exported_pdf).name.startswith("DocxPDF-"))
                    self.assertTrue(Path(exported_pdf).exists())
                    if cancel_at == "during restoration":
                        cancelled.set()
                        if cancel_requested():
                            raise converter.ConversionCancelled()
                    if cancel_at == "before publishing":
                        cancelled.set()
                    return converter.ImageRestorationReport(1, 1, 0)

                with self.subTest(cancel_at=cancel_at):
                    with patch.object(converter.platform, "system", return_value="Windows"), patch.object(
                        converter,
                        "_load_com_modules",
                        return_value=(FakePythonCom(), FakeClient(FakeWord(FakeDocument()))),
                    ), patch.object(converter, "_wait_for_pdf", return_value=True), patch.object(
                        converter, "restore_original_images", side_effect=restore
                    ):
                        with converter.WordSession() as session:
                            if cancel_at is None:
                                result = session.convert_docx(
                                    source, output, cancel_requested=cancelled.is_set
                                )
                            else:
                                with self.assertRaises(converter.ConversionCancelled):
                                    session.convert_docx(
                                        source, output, cancel_requested=cancelled.is_set
                                    )

                    self.assertEqual(restore_calls, [cancelled.is_set])
                    self.assertFalse(list(folder.glob("DocxPDF-*")))
                    if cancel_at is None:
                        self.assertEqual(result.output, output.resolve())
                        self.assertEqual(result.images_restored, 1)
                        self.assertEqual(output.read_bytes(), b"%PDF-1.7\nmock")
                    else:
                        self.assertFalse(output.exists())

    def test_picture_crops_are_read_from_every_part_with_relationships(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            source_docx = Path(tmp) / "source.docx"
            with ZipFile(source_docx, "w") as archive:
                archive.writestr(
                    "word/document.xml",
                    picture_part(
                        "document",
                        [
                            ("rId1", 'r="25000"'),
                            ("rId1", 'r="25000"'),
                            ("rId2", 'l="1000"'),
                            ("rId3", ""),
                        ],
                    ),
                )
                archive.writestr(
                    "word/_rels/document.xml.rels",
                    relationships_part(
                        {
                            "rId1": "media/image1.png",
                            "rId2": "https://example.invalid/linked.png",
                            "rId3": "media/image3.png",
                        },
                        external={"rId2"},
                    ),
                )
                archive.writestr(
                    "word/header1.xml",
                    picture_part("hdr", [("rId1", 't="-5000" b="10%"')]),
                )
                archive.writestr(
                    "word/_rels/header1.xml.rels",
                    relationships_part({"rId1": "/word/media/image2.jpeg"}),
                )
                archive.writestr("word/footer1.xml", "<not-xml srcRect")
                archive.writestr("word/_rels/footer1.xml.rels", relationships_part({}))

            with ZipFile(source_docx) as archive:
                crops = converter._docx_picture_crops(archive)

        self.assertEqual(
            crops,
            {
                "word/media/image1.png": {(0, 0, 25000, 0)},
                "word/media/image2.jpeg": {(0, -5000, 0, 10000)},
            },
        )
        self.assertEqual(converter._crop_box((100, 80), (0, 0, 25000, 0)), (0, 0, 75, 80))
        # Negative edges only pad the frame, so they do not cut source pixels.
        self.assertEqual(converter._crop_box((100, 80), (0, -5000, 0, 10000)), (0, 0, 100, 72))
        self.assertIsNone(converter._crop_box((100, 80), (-5000, 0, 0, 0)))
        self.assertIsNone(converter._crop_box((100, 80), (60000, 0, 50000, 0)))

    def test_conversion_result_counts_images_left_at_word_quality(self) -> None:
        result = converter.ConversionResult(
            Path("a.docx"), Path("a.pdf"), 1, 0.1, images_restored=2, images_examined=5
        )
        self.assertEqual(result.images_unrestored, 3)
        self.assertEqual(
            converter.ConversionResult(Path("a.docx"), Path("a.pdf"), 1, 0.1).images_unrestored,
            0,
        )

    def test_merge_pdfs_does_not_reencode_embedded_image_streams(self) -> None:
        try:
            from PIL import Image
            from pypdf import PdfReader
        except ImportError:
            self.skipTest("PDF/image QA dependencies are not installed")

        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            parts = []
            for index in range(2):
                part = folder / f"image-{index}.pdf"
                Image.new("RGB", (32, 32), (10, 20, 30)).save(
                    part,
                    "PDF",
                    quality=95,
                    subsampling=0,
                )
                parts.append(part)

            before = [
                PdfReader(str(part)).pages[0].images[0].indirect_reference.get_object()._data
                for part in parts
            ]
            result = converter.merge_pdfs(parts, folder / "merged.pdf")
            after_reader = PdfReader(str(result.output))
            after = [
                after_reader.pages[index].images[0].indirect_reference.get_object()._data
                for index in range(2)
            ]
            self.assertEqual(before, after)

    def test_merge_cancellation_after_write_preserves_existing_output(self) -> None:
        from pypdf import PdfWriter
        from threading import Event

        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            paths = [folder / "one.pdf", folder / "two.pdf", folder / "merged.pdf"]
            for path in paths:
                writer = PdfWriter()
                writer.add_blank_page(width=100, height=100)
                writer.write(path)
            original = paths[-1].read_bytes()
            cancelled = Event()
            check_pdf = converter._check_pdf

            def cancel_after_staged_file(path, **kwargs):
                size = check_pdf(path, **kwargs)
                if path.name.startswith("DocxPDF-merge-"):
                    cancelled.set()
                return size

            with patch.object(converter, "_check_pdf", side_effect=cancel_after_staged_file):
                with self.assertRaises(converter.MergeCancelled):
                    converter.merge_pdfs(paths[:2], paths[-1], overwrite=True, cancel_requested=cancelled.is_set)
            self.assertEqual(paths[-1].read_bytes(), original)
            self.assertFalse(list(folder.glob("DocxPDF-merge-*")))

    def test_merge_output_cannot_replace_an_input(self) -> None:
        from pypdf import PdfWriter

        with tempfile.TemporaryDirectory() as tmp:
            paths = [Path(tmp) / "one.pdf", Path(tmp) / "two.pdf"]
            for path in paths:
                writer = PdfWriter()
                writer.add_blank_page(width=100, height=100)
                writer.write(path)
            originals = [path.read_bytes() for path in paths]
            with self.assertRaisesRegex(converter.ConversionError, "cannot overwrite an input"):
                converter.merge_pdfs(paths, paths[0], overwrite=True, language="en")
            self.assertEqual([path.read_bytes() for path in paths], originals)


if __name__ == "__main__":
    unittest.main()
