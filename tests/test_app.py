from __future__ import annotations

import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import app
from PyQt6.QtWidgets import QApplication
from converter import ConversionResult


class AppTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.qt_app = QApplication.instance() or QApplication([])
        app.configure_application(cls.qt_app)

    def test_diagnostic_convert_writes_a_report(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            source = folder / "source.docx"
            output = folder / "output.pdf"
            report = folder / "report.txt"
            result = ConversionResult(source, output, 123, 0.1)

            with patch.object(app, "convert_docx", return_value=result):
                code = app.diagnostic_convert([str(source), str(output), str(report)])

            self.assertEqual(code, 0)
            self.assertEqual(report.read_text(encoding="utf-8"), f"OK\n{output}\n123\n")

    def test_fatal_conversion_error_stops_the_batch(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            sources = [folder / "one.docx", folder / "two.docx"]
            worker = app.ConversionWorker(sources, folder, False, False)
            error = app.ConversionError("Word 全局状态错误", abort_batch=True)
            session = MagicMock()
            session.__enter__.return_value = session
            session.__exit__.return_value = None
            session.convert_docx.side_effect = error

            with patch.object(app, "WordSession", return_value=session):
                worker.run()

            self.assertEqual(session.convert_docx.call_count, 1)

    def test_cancel_before_first_file_skips_conversion(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            worker = app.ConversionWorker([folder / "one.docx"], folder, False, False)
            session = MagicMock()
            session.__enter__.return_value = session
            session.__exit__.return_value = None
            finished_payloads = []
            worker.finished.connect(lambda *payload: finished_payloads.append(payload))
            worker.request_cancel()

            with patch.object(app, "WordSession", return_value=session):
                worker.run()

            session.convert_docx.assert_not_called()
            self.assertTrue(finished_payloads[-1][3])

    def test_language_switch_retranslates_controls_and_file_state(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "sample.docx"
            source.touch()
            settings = MagicMock()
            settings.value.return_value = ""

            with patch.object(app, "QSettings", return_value=settings):
                with patch.object(app, "locate_word_app", return_value=None):
                    window = app.MainWindow(language="zh")
                    window.add_files([source])
                    self.assertEqual(window.title_label.text(), "DOCX → PDF，保持原始版式")
                    self.assertIn("等待", window.file_list.item(0).text())
                    window.on_file_succeeded(
                        ConversionResult(
                            source=source.resolve(),
                            output=Path(tmp) / "sample.pdf",
                            size_bytes=1024,
                            elapsed_seconds=0.1,
                            images_restored=2,
                            images_examined=2,
                        ),
                        1,
                        1,
                    )
                    self.assertIn("原图保真 2 张", window.status_label.text())

                    english_index = window.language_combo.findData("en")
                    window.language_combo.setCurrentIndex(english_index)

                    self.assertEqual(
                        window.title_label.text(),
                        "DOCX → PDF, faithful to the original",
                    )
                    self.assertIn("2 original image(s) restored", window.file_list.item(0).text())
                    self.assertIn("2 original image(s) restored", window.status_label.text())
                    self.assertEqual(window.output_button.text(), "Browse…")
                    settings.setValue.assert_called_with("ui/language", "en")
                    window.close()


if __name__ == "__main__":
    unittest.main()
