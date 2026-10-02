from __future__ import annotations

import os
import tempfile
import threading
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import app
from PyQt6.QtWidgets import QApplication, QBoxLayout
from PyQt6.QtCore import QEventLoop, QTimer
from converter import ConversionResult, MergeResult


class AppTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.qt_app = QApplication.instance() or QApplication([])
        app.configure_application(cls.qt_app)

    def test_two_step_pdf_merge_preserves_images_order_and_inputs_without_word(self) -> None:
        from PIL import Image
        from pypdf import PdfReader

        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            docs = [folder / "one.docx", folder / "two.docx"]
            pdfs = [folder / "Merged.pdf", folder / "第二份.pdf"]
            for index, (doc, pdf) in enumerate(zip(docs, pdfs)):
                doc.touch()
                Image.new("RGB", (32 + index, 40), (10 + index, 20, 30)).save(pdf, "PDF", quality=95)
            original_files = [pdf.read_bytes() for pdf in pdfs]
            original_streams = [
                PdfReader(pdf).pages[0].images[0].indirect_reference.get_object()._data
                for pdf in pdfs
            ]
            settings = MagicMock()
            settings.value.return_value = ""
            with patch.object(app, "QSettings", return_value=settings), patch.object(app, "locate_word_app", return_value=None), patch.object(app, "WordSession") as word, patch.object(app.os, "startfile") as open_folder:
                window = app.MainWindow(language="en")
                window.reveal_box.setChecked(False)
                window.add_files(docs)
                results = [ConversionResult(doc.resolve(), pdf.resolve(), pdf.stat().st_size, 0.1) for doc, pdf in zip(docs, pdfs)]
                window.on_conversion_finished(results, [], None, False)
                window.mode_combo.setCurrentIndex(window.mode_combo.findData("pdf"))
                self.assertEqual(window.sources, [pdf.resolve() for pdf in pdfs])
                self.assertTrue(window.merge_button.isEnabled())
                self.assertTrue(window.convert_button.isHidden())
                window.overwrite_box.setChecked(True)
                window.reveal_box.setChecked(True)
                first = window.file_list.takeItem(0)
                window.file_list.addItem(first)
                window._sync_source_order()
                window.language_combo.setCurrentIndex(window.language_combo.findData("zh"))
                self.assertEqual(window.merge_button.text(), "合并 PDF")
                window.language_combo.setCurrentIndex(window.language_combo.findData("en"))
                loop = QEventLoop()
                window.start_merge()
                thread = window.thread
                timer = QTimer()
                timer.setSingleShot(True)
                timer.timeout.connect(loop.quit)
                timer.start(5000)
                thread.finished.connect(loop.quit)
                loop.exec()
                timer.stop()
                if window.running:
                    window.worker.request_cancel()
                    thread.quit()
                    thread.wait(5000)
                self.assertFalse(window.running)
                self.assertIn("Merge complete: 2 PDFs", window.status_label.text())
                merged = PdfReader(folder / "Merged-1.pdf")
                self.assertEqual([page.images[0].indirect_reference.get_object()._data for page in merged.pages], list(reversed(original_streams)))
                self.assertEqual([pdf.read_bytes() for pdf in pdfs], original_files)
                open_folder.assert_called_once_with(str(folder.resolve()), "explore")
                word.assert_not_called()
                window.mode_combo.setCurrentIndex(window.mode_combo.findData("docx"))
                self.assertEqual(window.sources, [doc.resolve() for doc in docs])
                window.mode_combo.setCurrentIndex(window.mode_combo.findData("pdf"))
                self.assertEqual(window.sources, [pdf.resolve() for pdf in reversed(pdfs)])
                window.close()

    def test_finished_opens_actual_output_folder_and_reports_open_failure(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            input_folder = folder / "Documents"
            output_folder = folder / "输出 location"
            input_folder.mkdir()
            output_folder.mkdir()
            sources = [input_folder / "one.pdf", input_folder / "two.pdf"]
            for source in sources:
                source.touch()
            output = output_folder / "合并 result.pdf"
            output.touch()
            result = MergeResult(tuple(sources), output, 2, 100, 0.1)
            with patch.object(app, "QSettings", return_value=MagicMock()), patch.object(app, "locate_word_app", return_value=None):
                window = app.MainWindow(language="en")
                window.mode_combo.setCurrentIndex(window.mode_combo.findData("pdf"))
                window.add_files(sources)
                window.job_merge = window.job_pdf_merge = True
                # The folder must come from the saved result, even if the edit
                # still displays Documents; never pass a PDF to ShellExecute.
                with patch.object(app.os, "startfile") as open_folder:
                    window.on_conversion_finished([], [], result, False)
                    open_folder.assert_called_once_with(str(output_folder.resolve()), "explore")
                    self.assertEqual(window.reveal_box.text(), "Open output folder when finished")
                    window.reveal_box.setChecked(False)
                    window.on_conversion_finished([], [], result, False)
                    self.assertEqual(open_folder.call_count, 1)
                window.reveal_box.setChecked(True)
                with patch.object(app.os, "startfile", side_effect=OSError("Explorer unavailable")), patch.object(window, "_show_warning") as warning:
                    window.on_conversion_finished([], [], result, False)
                    self.assertEqual(warning.call_args.args[0], "open_output_failed_title")
                    self.assertIn(str(output_folder.resolve()), warning.call_args.args[1])
                    self.assertIn("Merge complete", window.status_label.text())
                window.close()

    def test_pdf_mode_filters_picker_and_drop_files(self) -> None:
        from PyQt6.QtCore import QMimeData, QUrl

        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            pdf, doc = folder / "one.PDF", folder / "one.docx"
            pdf.touch()
            doc.touch()
            with patch.object(app, "QSettings", return_value=MagicMock()), patch.object(app, "locate_word_app", return_value=None):
                window = app.MainWindow(language="en")
                window.mode_combo.setCurrentIndex(window.mode_combo.findData("pdf"))
                event = MagicMock()
                mime = QMimeData()
                mime.setUrls([QUrl.fromLocalFile(str(pdf)), QUrl.fromLocalFile(str(doc))])
                event.mimeData.return_value = mime
                self.assertEqual(window.drop_panel._file_paths(event), [pdf])
                with patch.object(app.QFileDialog, "getOpenFileNames", return_value=([str(pdf)], "")) as picker:
                    window.choose_files()
                self.assertEqual(picker.call_args.args[-1], "PDF documents (*.pdf)")
                window.add_files([pdf, doc])
                self.assertEqual(window.sources, [pdf.resolve()])
                self.assertFalse(window.merge_button.isEnabled())
                window.close()

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

    def test_nonfatal_word_failure_restarts_session_before_next_file(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            sources = [folder / "one.docx", folder / "two.docx"]
            first_session = MagicMock()
            first_session.__enter__.return_value = first_session
            first_session.__exit__.return_value = None
            first_session.convert_docx.side_effect = app.ConversionError(
                "Microsoft Word 导出失败：Command failed"
            )
            second_session = MagicMock()
            second_session.__enter__.return_value = second_session
            second_session.__exit__.return_value = None
            second_session.convert_docx.return_value = ConversionResult(
                sources[1].resolve(), folder / "two.pdf", 10, 0.1
            )

            worker = app.ConversionWorker(sources, folder, False, False)
            with patch.object(
                app,
                "WordSession",
                side_effect=[first_session, second_session],
            ) as session_factory:
                worker.run()

            self.assertEqual(session_factory.call_count, 2)
            first_session.convert_docx.assert_called_once()
            second_session.convert_docx.assert_called_once()
            first_session.__exit__.assert_called_once()
            second_session.__exit__.assert_called_once()

    def test_long_batch_rotates_word_session(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            sources = [folder / f"file-{index:02d}.docx" for index in range(33)]
            sessions = []
            for source in sources:
                session = MagicMock()
                session.__enter__.return_value = session
                session.__exit__.return_value = None
                session.convert_docx.return_value = ConversionResult(
                    source.resolve(), folder / f"{source.stem}.pdf", 10, 0.1
                )
                sessions.append(session)

            worker = app.ConversionWorker(sources, folder, False, False)
            with patch.object(app, "WordSession", side_effect=sessions) as session_factory:
                worker.run()

            self.assertEqual(session_factory.call_count, 2)
            self.assertEqual(sum(s.convert_docx.call_count for s in sessions), 33)

    def test_parallel_lanes_overlap_and_keep_result_order(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            first = folder / "first" / "same-name.docx"
            second = folder / "second" / "same-name.docx"
            first.parent.mkdir()
            second.parent.mkdir()
            first.touch()
            second.touch()
            sources = [first, second]
            barrier = threading.Barrier(2)
            active = 0
            max_active = 0
            active_lock = threading.Lock()

            def convert(source, output, **_kwargs):
                nonlocal active, max_active
                with active_lock:
                    active += 1
                    max_active = max(max_active, active)
                try:
                    barrier.wait(timeout=3)
                finally:
                    with active_lock:
                        active -= 1
                return ConversionResult(source.resolve(), Path(output), 10, 0.1)

            sessions = []
            for _ in sources:
                session = MagicMock()
                session.__enter__.return_value = session
                session.__exit__.return_value = None
                session.convert_docx.side_effect = convert
                sessions.append(session)

            worker = app.ConversionWorker(
                sources,
                folder,
                False,
                False,
                parallel_workers=2,
            )
            finished_payloads = []
            worker.finished.connect(lambda *payload: finished_payloads.append(payload))
            with patch.object(app, "WordSession", side_effect=sessions):
                worker.run()

            self.assertEqual(max_active, 2)
            self.assertEqual(
                [result.source for result in finished_payloads[-1][0]],
                [source.resolve() for source in sources],
            )
            self.assertEqual(
                [result.output.name for result in finished_payloads[-1][0]],
                ["same-name.pdf", "same-name-1.pdf"],
            )
            self.assertFalse(finished_payloads[-1][1])

    def test_language_switch_retranslates_controls_and_file_state(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "sample.docx"
            source.touch()
            settings = MagicMock()
            settings.value.return_value = ""

            with patch.object(app, "QSettings", return_value=settings):
                with patch.object(app, "locate_word_app", return_value=None):
                    window = app.MainWindow(language="zh")
                    self.assertEqual(window.language_combo.view().objectName(), "languagePopup")
                    self.assertEqual(window.language_combo.itemText(0), "中文")
                    self.assertEqual(window.language_combo.itemText(1), "English")
                    window.add_files([source])
                    self.assertEqual(window.title_label.text(), "DOCX → PDF，保持原始版式")
                    self.assertIn("等待", window.file_list.item(0).text())
                    output_path = str(source.resolve().parent)
                    self.assertEqual(window.output_edit.text(), output_path)
                    self.assertEqual(window.output_edit.cursorPosition(), 0)
                    self.assertIn(output_path, window.output_path_preview.text())
                    self.assertEqual(window.output_edit.toolTip(), output_path)
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
                    self.assertIn("Current folder:", window.output_path_preview.text())
                    settings.setValue.assert_called_with("ui/language", "en")
                    window.close()

    def test_theme_switch_persists_and_layout_adapts_to_narrow_window(self) -> None:
        settings = MagicMock()
        settings.value.return_value = ""

        with patch.object(app, "QSettings", return_value=settings):
            with patch.object(app, "locate_word_app", return_value=None):
                window = app.MainWindow(language="en")
                self.assertEqual(window.theme_combo.itemText(0), "Light")
                self.assertEqual(window.theme_combo.itemText(1), "Dark")
                self.assertTrue(window.scroll_area.widgetResizable())

                dark_index = window.theme_combo.findData("dark")
                window.theme_combo.setCurrentIndex(dark_index)
                self.assertEqual(window.theme, "dark")
                self.assertIn("#000000", self.qt_app.styleSheet())
                settings.setValue.assert_any_call("ui/theme", "dark")

                window.resize(600, 540)
                window._apply_responsive_layout()
                for row in (
                    window.header_row,
                    window.list_header,
                    window.output_row,
                    window.option_row,
                    window.footer,
                ):
                    self.assertEqual(row.direction(), QBoxLayout.Direction.LeftToRight)
                self.assertEqual(window.root_layout.contentsMargins().left(), 12)
                self.assertEqual(window.convert_button.minimumWidth(), 116)
                self.assertEqual(
                    window.workspace_layout.direction(),
                    QBoxLayout.Direction.TopToBottom,
                )

                window.resize(1120, 740)
                window._apply_responsive_layout()
                for row in (
                    window.header_row,
                    window.list_header,
                    window.output_row,
                    window.option_row,
                    window.footer,
                ):
                    self.assertEqual(row.direction(), QBoxLayout.Direction.LeftToRight)
                self.assertEqual(window.root_layout.contentsMargins().left(), 34)
                self.assertEqual(window.convert_button.minimumWidth(), 144)
                self.assertEqual(
                    window.workspace_layout.direction(),
                    QBoxLayout.Direction.LeftToRight,
                )
                self.assertEqual(window.workspace_layout.stretch(0), 3)
                self.assertEqual(window.workspace_layout.stretch(1), 2)
                self.assertEqual(window.file_list.minimumHeight(), 220)

                light_index = window.theme_combo.findData("light")
                window.theme_combo.setCurrentIndex(light_index)
                self.assertEqual(window.theme, "light")
                window.close()


if __name__ == "__main__":
    unittest.main()
