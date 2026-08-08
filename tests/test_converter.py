from __future__ import annotations

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


class FakeClient:
    def __init__(self, word: FakeWord) -> None:
        self.word = word
        self.dispatch_calls: list[str] = []

    def DispatchEx(self, prog_id: str) -> FakeWord:
        self.dispatch_calls.append(prog_id)
        return self.word


class ConverterTests(unittest.TestCase):
    def test_next_output_path_is_non_destructive(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            source = folder / "说明.docx"
            source.touch()
            (folder / "说明.pdf").write_bytes(b"existing")
            self.assertEqual(converter.next_output_path(source, folder).name, "说明-1.pdf")
            self.assertEqual(converter.next_merge_path(folder).name, "合并结果.pdf")

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
            restored = after_reader.pages[0].images[0].image.convert("RGB")
            self.assertEqual(report.images_examined, 1)
            self.assertEqual(report.images_restored, 1)
            self.assertEqual(restored.size, original.size)
            self.assertIsNone(ImageChops.difference(original, restored).getbbox())
            self.assertEqual(after_reader.pages[0].get_contents().get_data(), before_content)
            self.assertEqual(tuple(after_reader.pages[0].mediabox), before_box)

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


if __name__ == "__main__":
    unittest.main()
