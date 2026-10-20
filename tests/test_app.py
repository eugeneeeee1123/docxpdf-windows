from __future__ import annotations

import os
import tempfile
import threading
import time
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import app
from PyQt6.QtWidgets import QApplication, QBoxLayout
from PyQt6.QtCore import QEventLoop, Qt, QTimer
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

    @staticmethod
    def _stuck_word_session(started: threading.Event | None = None) -> MagicMock:
        """A session whose export hangs, like Word on a hidden dialog.

        The export only fails once abort() ends the (fake) Word process.
        """
        aborted = threading.Event()
        session = MagicMock()
        session.__enter__.return_value = session
        session.__exit__.return_value = None
        session.aborted_at = None

        def abort():
            session.aborted_at = time.monotonic()
            aborted.set()
            return True

        def stuck_export(_source, _output, **_kwargs):
            if started is not None:
                started.set()
            if not aborted.wait(10):
                raise AssertionError("the stuck export was never aborted")
            raise app.ConversionError("The RPC server is unavailable.", abort_batch=True)

        session.abort.side_effect = abort
        session.convert_docx.side_effect = stuck_export
        return session

    @staticmethod
    def _restoring_word_session(started: threading.Event) -> MagicMock:
        """A session whose Word export is done and whose image restoration is slow.

        Like converter.restore_original_images, restoration polls
        cancel_requested between images and stops with ConversionCancelled.
        """
        session = MagicMock()
        session.__enter__.return_value = session
        session.__exit__.return_value = None
        session.cancel_seen_at = None

        def restoring_export(_source, _output, *, cancel_requested=None, **_kwargs):
            started.set()
            deadline = time.monotonic() + 10
            while time.monotonic() < deadline:
                if cancel_requested is not None and cancel_requested():
                    session.cancel_seen_at = time.monotonic()
                    raise app.ConversionCancelled()
                time.sleep(0.01)
            raise AssertionError("image restoration never saw the cancel request")

        session.convert_docx.side_effect = restoring_export
        return session

    def _close_choosing_stop_remaining(self, window: app.MainWindow) -> tuple[float, float]:
        """Close a running window, choose "Stop remaining files" and wait for the job.

        Returns when the close was requested and when the job thread finished.
        """
        dialog = MagicMock()
        buttons = []
        dialog.addButton.side_effect = lambda *_args: buttons.append(MagicMock()) or buttons[-1]
        # The first button added is "Stop remaining files".
        dialog.clickedButton.side_effect = lambda: buttons[0]
        loop = QEventLoop()
        guard = QTimer()
        guard.setSingleShot(True)
        guard.timeout.connect(loop.quit)
        window.thread.finished.connect(loop.quit)
        with patch.object(app, "QMessageBox", return_value=dialog) as message_box, patch.object(window, "_show_warning") as warning, patch.object(window, "close") as deferred_close:
            closing_at = time.monotonic()
            app.MainWindow.close(window)
            self.assertTrue(window.close_after_cancel)
            guard.start(5000)
            loop.exec()
            finished_at = time.monotonic()
            guard.stop()
            QApplication.processEvents()

            self.assertEqual(message_box.call_count, 1)
            warning.assert_not_called()
            deferred_close.assert_called_once()
        self.assertFalse(window.running)
        self.assertFalse(window.close_after_cancel)
        return closing_at, finished_at

    def test_cancel_ends_stuck_word_export_after_grace_period(self) -> None:
        grace = 0.4
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            sources = [folder / "stuck.docx", folder / "next.docx"]
            started = threading.Event()
            session = self._stuck_word_session(started)
            worker = app.ConversionWorker(sources, folder, False, False, parallel_workers=1)
            finished_payloads = []
            failed = []
            worker.finished.connect(lambda *payload: finished_payloads.append(payload))
            # Lanes emit from pool threads; a direct connection records it here.
            worker.file_failed.connect(
                lambda *payload: failed.append(payload),
                type=Qt.ConnectionType.DirectConnection,
            )
            observed = {}

            def cancel_once_exporting():
                started.wait(5)
                observed["cancelled_at"] = time.monotonic()
                worker.request_cancel()
                time.sleep(grace / 4)
                observed["aborted_early"] = session.abort.called

            canceller = threading.Thread(target=cancel_once_exporting)
            with patch.object(app, "CANCEL_ABORT_GRACE_SECONDS", grace), patch.object(app, "WordSession", return_value=session):
                canceller.start()
                worker.run()
                finished_at = time.monotonic()
            canceller.join(5)

            self.assertFalse(observed["aborted_early"])
            session.abort.assert_called_once()
            self.assertGreaterEqual(session.aborted_at - observed["cancelled_at"], grace * 0.9)
            self.assertLess(finished_at - observed["cancelled_at"], grace + 2)
            session.convert_docx.assert_called_once()
            session.__exit__.assert_called_once()
            results, errors, merged_result, cancelled = finished_payloads[-1]
            self.assertEqual(results, [])
            self.assertEqual(errors, [])
            self.assertIsNone(merged_result)
            self.assertTrue(cancelled)
            self.assertEqual(failed, [])

    def test_cancel_keeps_export_that_finishes_within_grace_period(self) -> None:
        grace = 0.4
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            sources = [folder / "slow.docx", folder / "next.docx"]
            started = threading.Event()
            release = threading.Event()
            session = MagicMock()
            session.__enter__.return_value = session
            session.__exit__.return_value = None

            def slow_export(source, output, **_kwargs):
                started.set()
                release.wait(5)
                return ConversionResult(source.resolve(), Path(output), 10, 0.1)

            session.convert_docx.side_effect = slow_export
            worker = app.ConversionWorker(sources, folder, False, False, parallel_workers=1)
            finished_payloads = []
            worker.finished.connect(lambda *payload: finished_payloads.append(payload))

            def cancel_then_finish_export():
                started.wait(5)
                worker.request_cancel()
                time.sleep(grace / 4)
                release.set()

            canceller = threading.Thread(target=cancel_then_finish_export)
            with patch.object(app, "CANCEL_ABORT_GRACE_SECONDS", grace), patch.object(app, "WordSession", return_value=session):
                canceller.start()
                worker.run()
            canceller.join(5)

            # Finishing the job cancels the grace timer before it can fire.
            self.assertTrue(worker._cancel_timer.finished.is_set())
            time.sleep(grace + 0.1)
            session.abort.assert_not_called()
            session.convert_docx.assert_called_once()
            results, errors, _merged_result, cancelled = finished_payloads[-1]
            self.assertEqual([result.source for result in results], [sources[0].resolve()])
            self.assertEqual(errors, [])
            self.assertTrue(cancelled)

    def test_stop_remaining_on_close_ends_stuck_exports_as_cancelled(self) -> None:
        grace = 0.3
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            docs = [folder / "one.docx", folder / "two.docx"]
            for doc in docs:
                doc.touch()
            started = [threading.Event(), threading.Event()]
            sessions = [self._stuck_word_session(event) for event in started]
            settings = MagicMock()
            settings.value.return_value = ""
            with patch.object(app, "QSettings", return_value=settings), patch.object(app, "locate_word_app", return_value=None), patch.object(app, "CANCEL_ABORT_GRACE_SECONDS", grace), patch.object(app, "WordSession", side_effect=sessions), patch.dict(os.environ, {"DOCXPDF_WORD_WORKERS": "2"}):
                window = app.MainWindow(language="en")
                window.reveal_box.setChecked(False)
                window.add_files(docs)
                window.start_conversion()
                for event in started:
                    self.assertTrue(event.wait(5))

                closing_at, finished_at = self._close_choosing_stop_remaining(window)
                self.assertLess(finished_at - closing_at, grace + 2)
                for session in sessions:
                    session.abort.assert_called_once()
                    self.assertGreaterEqual(session.aborted_at - closing_at, grace * 0.9)
                self.assertEqual(
                    [window.item_states[str(source)][0] for source in window.sources],
                    ["cancelled", "cancelled"],
                )
                self.assertIn("Task cancelled; 0 file(s)", window.status_label.text())
                window.close()

    def test_cancel_interrupts_image_restoration_without_waiting_for_grace_period(self) -> None:
        grace = 5
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            sources = [folder / "heavy.docx", folder / "next.docx"]
            started = threading.Event()
            session = self._restoring_word_session(started)
            worker = app.ConversionWorker(sources, folder, False, False, parallel_workers=1)
            finished_payloads = []
            failed = []
            worker.finished.connect(lambda *payload: finished_payloads.append(payload))
            worker.file_failed.connect(
                lambda *payload: failed.append(payload),
                type=Qt.ConnectionType.DirectConnection,
            )
            observed = {}

            def cancel_during_restoration():
                started.wait(5)
                observed["cancelled_at"] = time.monotonic()
                worker.request_cancel()

            canceller = threading.Thread(target=cancel_during_restoration)
            with patch.object(app, "CANCEL_ABORT_GRACE_SECONDS", grace), patch.object(app, "WordSession", return_value=session):
                canceller.start()
                worker.run()
                finished_at = time.monotonic()
            canceller.join(5)

            # The lane passes the cancel event itself, so restoration stops at
            # once rather than when the cancel timer would end Word.
            self.assertEqual(
                session.convert_docx.call_args.kwargs["cancel_requested"],
                worker._cancel_event.is_set,
            )
            self.assertLess(finished_at - observed["cancelled_at"], 1)
            self.assertTrue(worker._cancel_timer.finished.is_set())
            session.abort.assert_not_called()
            session.convert_docx.assert_called_once()
            session.__exit__.assert_called_once()
            results, errors, merged_result, cancelled = finished_payloads[-1]
            self.assertEqual(results, [])
            self.assertEqual(errors, [])
            self.assertIsNone(merged_result)
            self.assertTrue(cancelled)
            self.assertEqual(failed, [])

    def test_stop_remaining_on_close_interrupts_image_restoration(self) -> None:
        grace = 5
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            docs = [folder / "one.docx", folder / "two.docx"]
            for doc in docs:
                doc.touch()
            started = [threading.Event(), threading.Event()]
            sessions = [self._restoring_word_session(event) for event in started]
            settings = MagicMock()
            settings.value.return_value = ""
            with patch.object(app, "QSettings", return_value=settings), patch.object(app, "locate_word_app", return_value=None), patch.object(app, "CANCEL_ABORT_GRACE_SECONDS", grace), patch.object(app, "WordSession", side_effect=sessions), patch.dict(os.environ, {"DOCXPDF_WORD_WORKERS": "2"}):
                window = app.MainWindow(language="en")
                window.reveal_box.setChecked(False)
                window.add_files(docs)
                window.start_conversion()
                for event in started:
                    self.assertTrue(event.wait(5))

                closing_at, finished_at = self._close_choosing_stop_remaining(window)
                self.assertLess(finished_at - closing_at, 2)
                for session in sessions:
                    self.assertIsNotNone(session.cancel_seen_at)
                    session.abort.assert_not_called()
                    session.__exit__.assert_called_once()
                self.assertEqual(
                    [window.item_states[str(source)][0] for source in window.sources],
                    ["cancelled", "cancelled"],
                )
                self.assertIn("Task cancelled; 0 file(s)", window.status_label.text())
                window.close()

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

    def test_images_left_at_word_quality_are_reported_in_both_languages(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            sources = [folder / "one.docx", folder / "two.docx", folder / "three.docx"]
            for source in sources:
                source.touch()
            results = [
                ConversionResult(
                    source.resolve(),
                    folder / f"{source.stem}.pdf",
                    1024,
                    0.1,
                    images_restored=restored,
                    images_examined=examined,
                )
                for source, (restored, examined) in zip(sources, ((2, 3), (0, 1), (4, 4)))
            ]
            settings = MagicMock()
            settings.value.return_value = ""

            with patch.object(app, "QSettings", return_value=settings):
                with patch.object(app, "locate_word_app", return_value=None):
                    window = app.MainWindow(language="zh")
                    window.add_files(sources)
                    window.on_file_succeeded(results[0], 1, 3)
                    self.assertIn("原图保真 2 张", window.file_list.item(0).text())
                    self.assertIn("1 张图片未能还原（仍为 Word 版本）", window.file_list.item(0).text())
                    self.assertIn("1 张图片未能还原（仍为 Word 版本）", window.status_label.text())
                    window.on_file_succeeded(results[1], 2, 3)
                    self.assertIn("1 张图片未能还原（仍为 Word 版本）", window.file_list.item(1).text())
                    window.on_file_succeeded(results[2], 3, 3)
                    self.assertNotIn("未能还原", window.file_list.item(2).text())
                    self.assertNotIn("未能还原", window.status_label.text())

                    window.on_conversion_finished(results, [], None, False)
                    self.assertEqual(
                        window.status_label.text(),
                        "全部完成：已生成 3 个 PDF，原图保真 6 张；2 张图片未能还原（仍为 Word 版本）。",
                    )

                    window.language_combo.setCurrentIndex(window.language_combo.findData("en"))
                    self.assertEqual(
                        window.status_label.text(),
                        "All done: created 3 PDF file(s), with 6 original image(s) restored; "
                        "2 image(s) could not be restored (still Word's version).",
                    )
                    self.assertIn(
                        "Done · 2 original image(s) restored · 1 image(s) could not be restored "
                        "(still Word's version)",
                        window.file_list.item(0).text(),
                    )
                    self.assertIn(
                        "Done · 1 image(s) could not be restored (still Word's version)",
                        window.file_list.item(1).text(),
                    )
                    self.assertNotIn("could not be restored", window.file_list.item(2).text())
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
