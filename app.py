from __future__ import annotations

import os
import subprocess
import sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from queue import Empty, Queue
from threading import Event, Lock
from typing import Any

from PyQt6.QtCore import QLocale, QObject, QSettings, Qt, QThread, QTimer, pyqtSignal, pyqtSlot
from PyQt6.QtGui import QCloseEvent, QFont, QFontDatabase, QMouseEvent
from PyQt6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QBoxLayout,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from converter import (
    ConversionError,
    ConversionResult,
    WordSession,
    convert_docx,
    ensure_output_dir,
    locate_word_app,
    merge_pdfs,
    next_merge_path,
    next_output_path,
    word_version,
)
from i18n import DEFAULT_LANGUAGE, normalize_language, tr


APP_NAME = "DocxPDF"
APP_VERSION = "1.0.8"
# Word can become unstable after a long sequence of COM exports even when each
# document is closed correctly. Keep batches bounded so one large folder cannot
# poison the rest of the conversion job.
WORD_SESSION_DOCUMENT_LIMIT = 32
# A Word COM instance is expensive and can become unstable when too many
# instances are automated at once. Two isolated sessions give batch jobs real
# overlap while keeping memory and COM pressure predictable. Set
# DOCXPDF_WORD_WORKERS=1..4 when troubleshooting a particular machine.
DEFAULT_PARALLEL_WORD_WORKERS = 2
MAX_PARALLEL_WORD_WORKERS = 4
SUPPORTED_THEMES = ("light", "dark")


def normalize_theme(value: object) -> str:
    theme = str(value or "").strip().lower()
    return theme if theme in SUPPORTED_THEMES else "light"


class DropPanel(QFrame):
    files_dropped = pyqtSignal(object)
    choose_requested = pyqtSignal()

    def __init__(self) -> None:
        super().__init__()
        self.setObjectName("dropPanel")
        self.setAcceptDrops(True)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setMinimumHeight(128)

        self.title_label = QLabel()
        self.title_label.setObjectName("dropTitle")
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.hint_label = QLabel()
        self.hint_label.setObjectName("dropHint")
        self.hint_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 26, 24, 26)
        layout.setSpacing(7)
        layout.addStretch()
        layout.addWidget(self.title_label)
        layout.addWidget(self.hint_label)
        layout.addStretch()

    def set_language(self, language: str) -> None:
        self.title_label.setText(tr(language, "drop_title"))
        self.hint_label.setText(tr(language, "drop_hint"))

    @staticmethod
    def _docx_paths(event) -> list[Path]:
        if not event.mimeData().hasUrls():
            return []
        return [
            Path(url.toLocalFile())
            for url in event.mimeData().urls()
            if url.isLocalFile() and Path(url.toLocalFile()).suffix.lower() == ".docx"
        ]

    def dragEnterEvent(self, event) -> None:
        if self._docx_paths(event):
            self.setProperty("dragActive", True)
            self.style().unpolish(self)
            self.style().polish(self)
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, event) -> None:
        self.setProperty("dragActive", False)
        self.style().unpolish(self)
        self.style().polish(self)
        event.accept()

    def dropEvent(self, event) -> None:
        paths = self._docx_paths(event)
        self.setProperty("dragActive", False)
        self.style().unpolish(self)
        self.style().polish(self)
        if paths:
            self.files_dropped.emit(paths)
            event.acceptProposedAction()

    def mousePressEvent(self, event: QMouseEvent) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            self.choose_requested.emit()
        super().mousePressEvent(event)


class ConversionWorker(QObject):
    file_started = pyqtSignal(str, int, int)
    file_succeeded = pyqtSignal(object, int, int)
    file_failed = pyqtSignal(str, str, int, int)
    merge_started = pyqtSignal()
    finished = pyqtSignal(object, object, object, bool)

    def __init__(
        self,
        sources: list[Path],
        output_dir: Path,
        overwrite: bool,
        merge: bool,
        language: str = DEFAULT_LANGUAGE,
        *,
        parallel_workers: int | None = None,
    ) -> None:
        super().__init__()
        self.sources = sources
        self.output_dir = output_dir
        self.overwrite = overwrite
        self.merge = merge
        self.language = normalize_language(language)
        self.cancel_requested = False
        self.parallel_workers = parallel_workers
        self._cancel_event = Event()
        self._stop_event = Event()

    def request_cancel(self) -> None:
        """Stop after the current Word export returns."""
        self.cancel_requested = True
        self._cancel_event.set()

    def _worker_count(self, total: int) -> int:
        if total < 2:
            return 1
        configured = self.parallel_workers
        if configured is None:
            raw = os.environ.get("DOCXPDF_WORD_WORKERS", "").strip()
            if raw:
                try:
                    configured = int(raw)
                except ValueError:
                    configured = DEFAULT_PARALLEL_WORD_WORKERS
            else:
                configured = DEFAULT_PARALLEL_WORD_WORKERS
        return max(1, min(MAX_PARALLEL_WORD_WORKERS, int(configured), total))

    @staticmethod
    def _close_session(session_context: Any | None, entered: bool) -> None:
        if session_context is None or not entered:
            return
        try:
            session_context.__exit__(None, None, None)
        except Exception:
            # The conversion result/error is more useful than a cleanup error;
            # WordSession itself already makes cleanup best-effort.
            pass

    def _record_task_error(
        self,
        errors: list[tuple[Path, str]],
        lock: Lock,
        path: Path,
        message: str,
    ) -> None:
        with lock:
            # A startup failure can be observed by more than one lane. Keep a
            # single batch-level error instead of showing duplicate messages.
            if not errors:
                errors.append((path, message))

    def _run_conversion_lane(
        self,
        work_queue: Queue[int],
        output_paths: dict[int, Path],
        results: dict[int, ConversionResult],
        source_errors: dict[int, tuple[Path, str]],
        task_errors: list[tuple[Path, str]],
        state_lock: Lock,
        total: int,
    ) -> None:
        """Convert a dynamic slice of files on one COM-initialized thread.

        Each lane owns its WordSession from start to finish. No COM object is
        shared between Python threads; only plain paths and result records cross
        the lane boundary.
        """
        session_context: Any | None = None
        session: Any | None = None
        session_entered = False
        documents_in_session = 0

        def close_session() -> None:
            nonlocal session_context, session, session_entered, documents_in_session
            self._close_session(session_context, session_entered)
            session_context = None
            session = None
            session_entered = False
            documents_in_session = 0

        def start_session() -> None:
            nonlocal session_context, session, session_entered, documents_in_session
            session_context = WordSession(language=self.language)
            try:
                session = session_context.__enter__()
                session_entered = True
                documents_in_session = 0
            except Exception:
                session_context = None
                session = None
                session_entered = False
                raise

        try:
            while not self._cancel_event.is_set() and not self._stop_event.is_set():
                try:
                    index = work_queue.get_nowait()
                except Empty:
                    break

                try:
                    if self._cancel_event.is_set() or self._stop_event.is_set():
                        continue
                    if session is None or documents_in_session >= WORD_SESSION_DOCUMENT_LIMIT:
                        close_session()
                        try:
                            start_session()
                        except ConversionError as exc:
                            self._record_task_error(
                                task_errors,
                                state_lock,
                                Path(tr(self.language, "word_label")),
                                str(exc),
                            )
                            self._stop_event.set()
                            continue
                        except Exception as exc:
                            self._record_task_error(
                                task_errors,
                                state_lock,
                                Path(tr(self.language, "task_label")),
                                tr(self.language, "unexpected_error", error=exc),
                            )
                            self._stop_event.set()
                            continue

                    source = self.sources[index - 1]
                    self.file_started.emit(str(source), index, total)
                    output = output_paths[index]
                    try:
                        result = session.convert_docx(
                            source,
                            output,
                            overwrite=True if self.merge else self.overwrite,
                            timeout=300,
                        )
                    except ConversionError as exc:
                        message = str(exc)
                        with state_lock:
                            source_errors[index] = (source, message)
                        self.file_failed.emit(str(source), message, index, total)
                        # A failed document can poison the current Word COM
                        # server. Drop this session before taking more work.
                        close_session()
                        if exc.abort_batch:
                            self._stop_event.set()
                        continue
                    except Exception as exc:
                        message = tr(self.language, "unexpected_error", error=exc)
                        with state_lock:
                            source_errors[index] = (source, message)
                        self.file_failed.emit(str(source), message, index, total)
                        close_session()
                        continue

                    with state_lock:
                        results[index] = result
                    documents_in_session += 1
                    self.file_succeeded.emit(result, index, total)
                finally:
                    work_queue.task_done()
        finally:
            close_session()

    @pyqtSlot()
    def run(self) -> None:
        results: list[ConversionResult] = []
        errors: list[tuple[Path, str]] = []
        total = len(self.sources)
        temporary_dir = None
        cancelled = self._cancel_event.is_set()

        if self.merge:
            import tempfile
            try:
                self.output_dir.mkdir(parents=True, exist_ok=True)
                temporary_dir = tempfile.TemporaryDirectory(
                    prefix="docxpdf-merge-",
                    dir=str(self.output_dir),
                )
            except OSError as exc:
                self.finished.emit(
                    [],
                    [
                        (
                            Path(tr(self.language, "merge_result_label")),
                            tr(self.language, "output_dir_worker_error", error=exc),
                        )
                    ],
                    None,
                    False,
                )
                return

        try:
            output_paths: dict[int, Path] = {}
            reserved_outputs: set[Path] = set()
            for index, source in enumerate(self.sources, start=1):
                if self.merge:
                    output = Path(temporary_dir.name) / f"part-{index:04d}.pdf"
                else:
                    output = next_output_path(
                        source,
                        self.output_dir,
                        self.overwrite,
                        language=self.language,
                        reserved_paths=reserved_outputs,
                    )
                    reserved_outputs.add(output.resolve())
                output_paths[index] = output

            work_queue: Queue[int] = Queue()
            for index in range(1, total + 1):
                work_queue.put(index)
            results_by_index: dict[int, ConversionResult] = {}
            source_errors: dict[int, tuple[Path, str]] = {}
            task_errors: list[tuple[Path, str]] = []
            state_lock = Lock()
            self._stop_event.clear()
            worker_count = self._worker_count(total)

            with ThreadPoolExecutor(
                max_workers=worker_count,
                thread_name_prefix="DocxPDF-Word",
            ) as executor:
                futures = [
                    executor.submit(
                        self._run_conversion_lane,
                        work_queue,
                        output_paths,
                        results_by_index,
                        source_errors,
                        task_errors,
                        state_lock,
                        total,
                    )
                    for _ in range(worker_count)
                ]
                for future in futures:
                    future.result()

            results = [
                results_by_index[index]
                for index in range(1, total + 1)
                if index in results_by_index
            ]
            errors = [
                source_errors[index]
                for index in range(1, total + 1)
                if index in source_errors
            ]
            errors.extend(task_errors)
            cancelled = self._cancel_event.is_set()

            merged_result = None
            if self.merge and not errors and not cancelled:
                self.merge_started.emit()
                try:
                    temporary_pdfs = [result.output for result in results]
                    merged_result = merge_pdfs(
                        temporary_pdfs,
                        next_merge_path(
                            self.output_dir,
                            self.overwrite,
                            language=self.language,
                        ),
                        overwrite=self.overwrite,
                        language=self.language,
                    )
                except ConversionError as exc:
                    errors.append(
                        (Path(tr(self.language, "merge_result_label")), str(exc))
                    )
                except Exception as exc:  # Defensive boundary for a GUI worker.
                    errors.append(
                        (
                            Path(tr(self.language, "merge_result_label")),
                            tr(self.language, "unexpected_error", error=exc),
                        )
                    )
            self.finished.emit(results, errors, merged_result, cancelled)
        except ConversionError as exc:
            errors.append((Path(tr(self.language, "word_label")), str(exc)))
            self.finished.emit(results, errors, None, cancelled)
        except Exception as exc:  # Keep the GUI from staying stuck on an unexpected worker error.
            errors.append(
                (
                    Path(tr(self.language, "task_label")),
                    tr(self.language, "unexpected_error", error=exc),
                )
            )
            self.finished.emit(results, errors, None, cancelled)
        finally:
            if temporary_dir is not None:
                temporary_dir.cleanup()


class MainWindow(QMainWindow):
    def __init__(self, *, language: str | None = None) -> None:
        super().__init__()
        self.settings = QSettings()
        saved_language = self.settings.value("ui/language", "")
        saved_theme = self.settings.value("ui/theme", "light")
        environment_language = os.environ.get("DOCXPDF_LANGUAGE", "")
        if language or environment_language or saved_language:
            requested_language = language or environment_language or str(saved_language)
        else:
            requested_language = (
                "zh"
                if QLocale.system().language() == QLocale.Language.Chinese
                else "en"
            )
        self.language = normalize_language(requested_language)
        self.theme = normalize_theme(saved_theme)
        self.sources: list[Path] = []
        self.items: dict[str, QListWidgetItem] = {}
        self.item_states: dict[str, tuple[str, dict[str, object]]] = {}
        self.thread: QThread | None = None
        self.worker: ConversionWorker | None = None
        self.running = False
        self.job_merge = False
        self.close_after_cancel = False
        self._completed_files = 0
        self._status_key = "select_files"
        self._status_values: dict[str, object] = {}

        self.setMinimumSize(520, 540)
        self.resize(840, 740)
        self._build_ui()
        self._retranslate_ui()
        self._apply_theme()
        self._apply_responsive_layout()
        self._refresh_word_status()
        self._refresh_actions()

    def _build_ui(self) -> None:
        root = QWidget()
        root.setObjectName("root")
        root.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        self.scroll_area = QScrollArea()
        self.scroll_area.setObjectName("contentScroll")
        self.scroll_area.setFrameShape(QFrame.Shape.NoFrame)
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.scroll_area.setWidget(root)
        self.setCentralWidget(self.scroll_area)

        layout = QVBoxLayout(root)
        self.root_layout = layout
        layout.setContentsMargins(38, 32, 38, 30)
        # Keep the file list and output section visually distinct at the
        # default 840×740 window size; 16px made their minimum heights collide.
        layout.setSpacing(12)

        header_row = QHBoxLayout()
        self.header_row = header_row
        self.eyebrow_label = QLabel()
        self.eyebrow_label.setObjectName("eyebrow")
        self.eyebrow_label.setWordWrap(True)
        header_row.addWidget(self.eyebrow_label, 1)
        self.language_combo = QComboBox()
        self.language_combo.setObjectName("languageCombo")
        self.language_combo.addItem("中文", "zh")
        self.language_combo.addItem("English", "en")
        language_popup = self.language_combo.view()
        language_popup.setObjectName("languagePopup")
        self.language_combo.setMinimumWidth(104)
        language_index = self.language_combo.findData(self.language)
        self.language_combo.setCurrentIndex(max(0, language_index))
        self.language_combo.currentIndexChanged.connect(self.change_language)
        header_row.addWidget(self.language_combo)

        self.theme_combo = QComboBox()
        self.theme_combo.setObjectName("themeCombo")
        self.theme_combo.addItem("", "light")
        self.theme_combo.addItem("", "dark")
        theme_popup = self.theme_combo.view()
        theme_popup.setObjectName("themePopup")
        self.theme_combo.setMinimumWidth(104)
        theme_index = self.theme_combo.findData(self.theme)
        self.theme_combo.setCurrentIndex(max(0, theme_index))
        self.theme_combo.currentIndexChanged.connect(self.change_theme)
        header_row.addWidget(self.theme_combo)
        layout.addLayout(header_row)

        self.title_label = QLabel()
        self.title_label.setObjectName("title")
        self.subtitle_label = QLabel()
        self.subtitle_label.setObjectName("subtitle")
        self.subtitle_label.setWordWrap(True)

        layout.addWidget(self.title_label)
        layout.addWidget(self.subtitle_label)

        self.workspace = QWidget()
        self.workspace.setObjectName("workspace")
        self.workspace.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )
        workspace_layout = QHBoxLayout(self.workspace)
        self.workspace_layout = workspace_layout
        workspace_layout.setContentsMargins(0, 0, 0, 0)
        workspace_layout.setSpacing(16)

        self.files_column = QWidget()
        self.files_column.setObjectName("filesColumn")
        self.files_column.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )
        files_column_layout = QVBoxLayout(self.files_column)
        self.files_column_layout = files_column_layout
        files_column_layout.setContentsMargins(0, 0, 0, 0)
        files_column_layout.setSpacing(10)

        self.controls_column = QWidget()
        self.controls_column.setObjectName("controlsColumn")
        self.controls_column.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred
        )
        controls_column_layout = QVBoxLayout(self.controls_column)
        self.controls_column_layout = controls_column_layout
        controls_column_layout.setContentsMargins(0, 0, 0, 0)
        controls_column_layout.setSpacing(10)

        self.drop_panel = DropPanel()
        self.drop_panel.files_dropped.connect(self.add_files)
        self.drop_panel.choose_requested.connect(self.choose_files)
        files_column_layout.addWidget(self.drop_panel)

        list_header = QHBoxLayout()
        self.list_header = list_header
        self.list_title_label = QLabel()
        self.list_title_label.setObjectName("sectionTitle")
        self.count_label = QLabel()
        self.count_label.setObjectName("muted")
        list_header.addWidget(self.list_title_label)
        list_header.addWidget(self.count_label)
        self.order_hint_label = QLabel()
        self.order_hint_label.setObjectName("muted")
        self.order_hint_label.setWordWrap(True)
        list_header.addWidget(self.order_hint_label)
        list_header.addStretch()

        self.remove_button = QPushButton()
        self.remove_button.setObjectName("quietButton")
        self.remove_button.clicked.connect(self.remove_selected)
        self.clear_button = QPushButton()
        self.clear_button.setObjectName("quietButton")
        self.clear_button.clicked.connect(self.clear_files)
        list_header.addWidget(self.remove_button)
        list_header.addWidget(self.clear_button)
        files_column_layout.addLayout(list_header)

        self.file_list = QListWidget()
        self.file_list.setObjectName("fileList")
        self.file_list.setMinimumHeight(128)
        self.file_list.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.file_list.setTextElideMode(Qt.TextElideMode.ElideMiddle)
        self.file_list.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        self.file_list.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.file_list.model().rowsMoved.connect(self._sync_source_order)
        files_column_layout.addWidget(self.file_list, 1)

        self.output_title_label = QLabel()
        self.output_title_label.setObjectName("sectionTitle")
        controls_column_layout.addWidget(self.output_title_label)

        output_row = QHBoxLayout()
        self.output_row = output_row
        output_row.setSpacing(10)
        self.output_edit = QLineEdit()
        self.output_edit.setClearButtonEnabled(True)
        self.output_edit.textChanged.connect(self._refresh_actions)
        self.output_edit.textChanged.connect(self._update_output_path_preview)
        output_row.addWidget(self.output_edit, 1)
        self.output_button = QPushButton()
        self.output_button.setObjectName("secondaryButton")
        self.output_button.clicked.connect(self.choose_output_dir)
        output_row.addWidget(self.output_button)
        controls_column_layout.addLayout(output_row)

        self.output_path_preview = QLabel()
        self.output_path_preview.setObjectName("pathPreview")
        self.output_path_preview.setWordWrap(True)
        self.output_path_preview.setTextFormat(Qt.TextFormat.PlainText)
        self.output_path_preview.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        controls_column_layout.addWidget(self.output_path_preview)

        option_row = QHBoxLayout()
        self.option_row = option_row
        self.overwrite_box = QCheckBox()
        self.reveal_box = QCheckBox()
        self.reveal_box.setChecked(True)
        option_row.addWidget(self.overwrite_box)
        option_row.addSpacing(18)
        option_row.addWidget(self.reveal_box)
        option_row.addStretch()
        controls_column_layout.addLayout(option_row)

        self.output_hint_label = QLabel()
        self.output_hint_label.setObjectName("muted")
        self.output_hint_label.setWordWrap(True)
        controls_column_layout.addWidget(self.output_hint_label)

        footer = QHBoxLayout()
        self.footer = footer
        footer.setSpacing(12)
        self.word_status = QLabel()
        self.word_status.setObjectName("wordStatus")
        self.word_status.setWordWrap(True)
        footer.addWidget(self.word_status)
        footer.addStretch()
        self.cancel_button = QPushButton()
        self.cancel_button.setObjectName("quietButton")
        self.cancel_button.clicked.connect(self.cancel_job)
        self.cancel_button.setVisible(False)
        footer.addWidget(self.cancel_button)
        self.convert_button = QPushButton()
        self.convert_button.setObjectName("primaryButton")
        self.convert_button.setMinimumWidth(144)
        self.convert_button.clicked.connect(self.start_conversion)
        footer.addWidget(self.convert_button)
        self.merge_button = QPushButton()
        self.merge_button.setObjectName("secondaryButton")
        self.merge_button.setMinimumWidth(132)
        self.merge_button.clicked.connect(self.start_merge)
        footer.addWidget(self.merge_button)
        controls_column_layout.addLayout(footer)

        self.progress = QProgressBar()
        self.progress.setRange(0, 1)
        self.progress.setValue(0)
        self.progress.setTextVisible(False)
        controls_column_layout.addWidget(self.progress)

        self.status_label = QLabel()
        self.status_label.setObjectName("statusLabel")
        self.status_label.setWordWrap(True)
        controls_column_layout.addWidget(self.status_label)

        workspace_layout.addWidget(self.files_column)
        workspace_layout.addWidget(self.controls_column)
        layout.addWidget(self.workspace, 1)

    def _t(self, key: str, **values) -> str:
        return tr(self.language, key, **values)

    def _set_status(self, key: str, **values) -> None:
        self._status_key = key
        self._status_values = dict(values)
        self.status_label.setText(self._translated_status())

    def _translated_status(self) -> str:
        values = dict(self._status_values)
        quality_count = int(values.pop("_quality_count", 0))
        quality_key = str(values.pop("_quality_key", "quality_note"))
        if self._status_key in {"file_completed", "merge_completed", "all_completed"}:
            values["quality_note"] = (
                self._t(quality_key, count=quality_count) if quality_count else ""
            )
        return self._t(self._status_key, **values)

    def _render_item(self, key: str) -> None:
        item = self.items.get(key)
        state = self.item_states.get(key)
        if item is None or state is None:
            return
        state_key, values = state
        item.setText(
            f"{Path(key).name}    ·    {self._t('item_' + state_key, **values)}"
        )

    def _retranslate_ui(self) -> None:
        self.setWindowTitle(self._t("app_name"))
        self.eyebrow_label.setText(self._t("eyebrow"))
        self.title_label.setText(self._t("title"))
        self.subtitle_label.setText(self._t("subtitle"))
        self.drop_panel.set_language(self.language)
        self.list_title_label.setText(self._t("files_section"))
        self.count_label.setText(self._t("file_count", count=len(self.sources)))
        self.order_hint_label.setText(self._t("order_hint"))
        self.remove_button.setText(self._t("remove_selected"))
        self.clear_button.setText(self._t("clear"))
        self.file_list.setAccessibleName(self._t("file_list_accessible"))
        self.output_title_label.setText(self._t("output_section"))
        self.output_edit.setPlaceholderText(self._t("output_placeholder"))
        self.output_edit.setAccessibleName(self._t("output_accessible"))
        self._update_output_path_preview()
        self.output_button.setText(self._t("browse"))
        self.output_button.setToolTip(self._t("browse_tooltip"))
        self.overwrite_box.setText(self._t("overwrite"))
        self.overwrite_box.setToolTip(self._t("overwrite_tooltip"))
        self.reveal_box.setText(self._t("reveal"))
        self.output_hint_label.setText(self._t("output_hint"))
        self.cancel_button.setText(self._t("cancel_task"))
        self.cancel_button.setToolTip(self._t("cancel_tooltip"))
        self.convert_button.setText(self._t("convert_individual"))
        self.convert_button.setToolTip(self._t("convert_individual_tooltip"))
        self.merge_button.setText(self._t("convert_merge"))
        self.merge_button.setToolTip(self._t("convert_merge_tooltip"))
        self.progress.setAccessibleName(self._t("progress_accessible"))
        self.language_combo.setAccessibleName(self._t("language_accessible"))
        self.language_combo.setToolTip(self._t("language_tooltip"))
        light_index = self.theme_combo.findData("light")
        dark_index = self.theme_combo.findData("dark")
        if light_index >= 0:
            self.theme_combo.setItemText(light_index, self._t("theme_light"))
        if dark_index >= 0:
            self.theme_combo.setItemText(dark_index, self._t("theme_dark"))
        self.theme_combo.setAccessibleName(self._t("theme_accessible"))
        self.theme_combo.setToolTip(self._t("theme_tooltip"))
        for key in self.items:
            self._render_item(key)
        self.status_label.setText(self._translated_status())
        self._refresh_word_status()

    @pyqtSlot(int)
    def change_language(self, index: int) -> None:
        language = normalize_language(self.language_combo.itemData(index))
        if language == self.language:
            return
        self.language = language
        self.settings.setValue("ui/language", language)
        self._retranslate_ui()

    @pyqtSlot(int)
    def change_theme(self, index: int) -> None:
        theme = normalize_theme(self.theme_combo.itemData(index))
        if theme == self.theme:
            return
        self.theme = theme
        self.settings.setValue("ui/theme", theme)
        self._apply_theme()

    def _apply_theme(self) -> None:
        application = QApplication.instance()
        if application is not None:
            application.setStyleSheet(stylesheet_for_theme(self.theme))

    def _apply_responsive_layout(self) -> None:
        if not hasattr(self, "root_layout"):
            return

        width = self.width()
        wide_workspace = width >= 960
        narrow = width <= 640
        compact = width < 820
        horizontal = QBoxLayout.Direction.LeftToRight
        vertical = QBoxLayout.Direction.TopToBottom

        if wide_workspace:
            margins = (34, 28, 34, 28)
        elif narrow:
            margins = (12, 18, 12, 18)
        elif compact:
            margins = (24, 24, 24, 24)
        else:
            margins = (24, 26, 24, 26)
        self.root_layout.setContentsMargins(*margins)
        self.root_layout.setSpacing(12 if wide_workspace else 8 if narrow else 10)

        # The desktop client uses a real two-column workspace. Only fall back
        # to stacked columns when the window is too narrow to keep both panels
        # usable; the controls inside each panel remain horizontal.
        self.workspace_layout.setDirection(horizontal if wide_workspace else vertical)
        self.workspace_layout.setSpacing(16 if wide_workspace else 12)
        self.workspace_layout.setStretch(0, 3 if wide_workspace else 0)
        self.workspace_layout.setStretch(1, 2 if wide_workspace else 0)
        self.workspace.setSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Expanding if wide_workspace else QSizePolicy.Policy.Preferred,
        )
        self.files_column.setSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Expanding if wide_workspace else QSizePolicy.Policy.Preferred,
        )
        self.controls_column.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred
        )

        self.header_row.setDirection(horizontal)
        self.header_row.setStretch(0, 1)
        self.output_row.setDirection(horizontal)
        self.output_row.setStretch(0, 1)
        self.option_row.setDirection(horizontal)
        self.list_header.setDirection(horizontal)
        self.list_header.setStretch(3, 1)
        self.footer.setDirection(horizontal)
        self.footer.setStretch(1, 1)

        combo_width = 88 if narrow else 96 if compact else 104
        output_button_width = 84 if narrow else 92 if compact else 0
        convert_button_width = 116 if narrow else 128 if compact else 144
        merge_button_width = 104 if narrow else 116 if compact else 132
        self.language_combo.setMinimumWidth(combo_width)
        self.theme_combo.setMinimumWidth(combo_width)
        self.output_button.setMinimumWidth(output_button_width)
        self.cancel_button.setMinimumWidth(84 if narrow else 96 if compact else 0)
        self.convert_button.setMinimumWidth(convert_button_width)
        self.merge_button.setMinimumWidth(merge_button_width)

        self.eyebrow_label.setSizePolicy(
            QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred
        )
        self.order_hint_label.setSizePolicy(
            QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred
        )
        self.word_status.setSizePolicy(
            QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred
        )
        if wide_workspace:
            drop_height = 148
            file_list_height = 220
        elif narrow:
            drop_height = 104
            file_list_height = 112
        else:
            drop_height = 116 if compact else 128
            file_list_height = 120 if compact else 140
        self.drop_panel.setMinimumHeight(drop_height)
        self.file_list.setMinimumHeight(file_list_height)
        self.file_list.setSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Expanding if wide_workspace else QSizePolicy.Policy.Preferred,
        )
        button_policy = QSizePolicy.Policy.Preferred
        for button in (
            self.output_button,
            self.cancel_button,
            self.convert_button,
            self.merge_button,
        ):
            button.setSizePolicy(button_policy, QSizePolicy.Policy.Fixed)

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self._apply_responsive_layout()

    def _refresh_word_status(self) -> None:
        app = locate_word_app()
        version = word_version(app)
        if app:
            suffix = f" {version}" if version else ""
            self.word_status.setText(self._t("word_ready", version=suffix))
            self.word_status.setToolTip(str(app))
            self.word_status.setProperty("available", True)
        else:
            self.word_status.setText(self._t("word_missing"))
            self.word_status.setToolTip(self._t("word_missing_tooltip"))
            self.word_status.setProperty("available", False)
        self.word_status.style().unpolish(self.word_status)
        self.word_status.style().polish(self.word_status)

    def _refresh_actions(self) -> None:
        has_files = bool(self.sources)
        has_output = bool(self.output_edit.text().strip()) if hasattr(self, "output_edit") else False
        word_ready = locate_word_app() is not None
        enabled = not self.running if hasattr(self, "running") else True
        if hasattr(self, "convert_button"):
            self.convert_button.setEnabled(has_files and has_output and word_ready and enabled)
            self.merge_button.setEnabled(
                len(self.sources) >= 2 and has_output and word_ready and enabled
            )
            self.remove_button.setEnabled(has_files and enabled)
            self.clear_button.setEnabled(has_files and enabled)
            self.output_button.setEnabled(enabled)
            self.output_edit.setEnabled(enabled)
            self.reveal_box.setEnabled(enabled)
            self.drop_panel.setEnabled(enabled)
            self.overwrite_box.setEnabled(enabled)
            self.language_combo.setEnabled(enabled)
            self.theme_combo.setEnabled(enabled)
            self.cancel_button.setVisible(not enabled)
            self.cancel_button.setEnabled(not enabled)

    def _update_output_path_preview(self) -> None:
        path = self.output_edit.text().strip()
        has_path = bool(path)
        if has_path:
            self.output_path_preview.setText(
                self._t("output_path_preview", path=path)
            )
            self.output_edit.setToolTip(path)
        else:
            self.output_path_preview.setText(self._t("output_path_empty"))
            self.output_edit.setToolTip(self._t("output_placeholder"))
        self.output_path_preview.setProperty("hasPath", has_path)
        self.output_path_preview.style().unpolish(self.output_path_preview)
        self.output_path_preview.style().polish(self.output_path_preview)

    def _set_output_path(self, path: str) -> None:
        """Set a chosen/default path and reveal its beginning immediately."""
        self.output_edit.setText(path)
        self.output_edit.setCursorPosition(0)

    @pyqtSlot()
    def choose_files(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            self._t("choose_docs_title"),
            str(Path.home() / "Documents"),
            self._t("word_documents_filter"),
        )
        if paths:
            self.add_files([Path(path) for path in paths])

    @pyqtSlot(object)
    def add_files(self, paths) -> None:
        added = 0
        skipped = 0
        for raw_path in paths:
            path = Path(raw_path).expanduser()
            if not path.is_file() or path.suffix.lower() != ".docx" or path.name.startswith("~$"):
                skipped += 1
                continue
            resolved = path.resolve()
            key = str(resolved)
            if key in self.items:
                skipped += 1
                continue
            item = QListWidgetItem()
            item.setData(Qt.ItemDataRole.UserRole, key)
            item.setToolTip(key)
            self.file_list.addItem(item)
            self.sources.append(resolved)
            self.items[key] = item
            self.item_states[key] = ("waiting", {})
            self._render_item(key)
            added += 1

        if self.sources and not self.output_edit.text().strip():
            self._set_output_path(str(self.sources[0].parent))
        self.count_label.setText(self._t("file_count", count=len(self.sources)))
        if added:
            if skipped:
                self._set_status("files_added_skipped", added=added, skipped=skipped)
            else:
                self._set_status("files_added", added=added)
        elif skipped:
            self._set_status("no_valid_files")
        self._refresh_actions()

    @pyqtSlot()
    def remove_selected(self) -> None:
        selected = list(self.file_list.selectedItems())
        if not selected:
            return
        keys = {str(item.data(Qt.ItemDataRole.UserRole)) for item in selected}
        for item in selected:
            self.file_list.takeItem(self.file_list.row(item))
        self.sources = [path for path in self.sources if str(path) not in keys]
        for key in keys:
            self.items.pop(key, None)
            self.item_states.pop(key, None)
        self.count_label.setText(self._t("file_count", count=len(self.sources)))
        self._refresh_actions()

    @pyqtSlot()
    def clear_files(self) -> None:
        self.sources.clear()
        self.items.clear()
        self.item_states.clear()
        self.file_list.clear()
        self.count_label.setText(self._t("file_count", count=0))
        self._set_status("select_files")
        self._refresh_actions()

    @pyqtSlot()
    def choose_output_dir(self) -> None:
        initial = self.output_edit.text().strip() or str(Path.home() / "Documents")
        selected = QFileDialog.getExistingDirectory(
            self,
            self._t("choose_output_title"),
            initial,
        )
        if selected:
            self._set_output_path(str(Path(selected).resolve()))

    def _set_item_status(self, source: str | Path, state: str, **values) -> None:
        key = str(Path(source))
        if key in self.items:
            self.item_states[key] = (state, dict(values))
            self._render_item(key)

    def _sync_source_order(self, *_args) -> None:
        self.sources = [
            Path(self.file_list.item(index).data(Qt.ItemDataRole.UserRole))
            for index in range(self.file_list.count())
        ]

    def _show_warning(self, title_key: str, message: str) -> None:
        dialog = QMessageBox(self)
        dialog.setIcon(QMessageBox.Icon.Warning)
        dialog.setWindowTitle(self._t(title_key))
        dialog.setText(message)
        dialog.addButton(self._t("ok"), QMessageBox.ButtonRole.AcceptRole)
        dialog.exec()

    @pyqtSlot()
    def start_conversion(self) -> None:
        self._start_job(merge=False)

    @pyqtSlot()
    def start_merge(self) -> None:
        self._start_job(merge=True)

    def _start_job(self, *, merge: bool) -> None:
        if self.running:
            return
        if not self.sources or not self.output_edit.text().strip():
            return
        if merge and len(self.sources) < 2:
            return
        try:
            output_dir = ensure_output_dir(
                self.output_edit.text().strip(),
                language=self.language,
            )
        except ConversionError as exc:
            self._show_warning("output_unavailable_title", str(exc))
            self._set_status(
                "file_error",
                name=self._t("output_section"),
                message=str(exc),
            )
            return
        self.output_edit.setText(str(output_dir))

        self.running = True
        self.job_merge = merge
        self.progress.setRange(0, len(self.sources) + (1 if merge else 0))
        self.progress.setValue(0)
        self._completed_files = 0
        self._set_status("starting_word" if not merge else "starting_merge")
        for source in self.sources:
            self._set_item_status(source, "waiting")
        self._refresh_actions()

        self.thread = QThread(self)
        self.worker = ConversionWorker(
            list(self.sources),
            output_dir,
            self.overwrite_box.isChecked(),
            merge,
            self.language,
        )
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.file_started.connect(self.on_file_started)
        self.worker.file_succeeded.connect(self.on_file_succeeded)
        self.worker.file_failed.connect(self.on_file_failed)
        self.worker.merge_started.connect(self.on_merge_started)
        self.worker.finished.connect(self.on_conversion_finished)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.thread.finished.connect(self.on_thread_finished)
        self.thread.start()

    @pyqtSlot()
    def cancel_job(self) -> None:
        if not self.running or self.worker is None:
            return
        self.worker.request_cancel()
        self.cancel_button.setEnabled(False)
        self._set_status("cancel_requested")

    @pyqtSlot(str, int, int)
    def on_file_started(self, source: str, index: int, total: int) -> None:
        self._set_item_status(source, "processing")
        self._set_status(
            "converting",
            index=index,
            total=total,
            name=Path(source).name,
        )

    @pyqtSlot(object, int, int)
    def on_file_succeeded(self, result: ConversionResult, index: int, total: int) -> None:
        if result.images_restored:
            self._set_item_status(
                result.source,
                "complete_quality",
                count=result.images_restored,
            )
        else:
            self._set_item_status(result.source, "complete")
        self._completed_files += 1
        self.progress.setValue(self._completed_files)
        size_mb = result.size_bytes / (1024 * 1024)
        self._set_status(
            "file_completed",
            index=index,
            total=total,
            name=result.output.name,
            size_mb=size_mb,
            _quality_count=result.images_restored,
        )

    @pyqtSlot(str, str, int, int)
    def on_file_failed(self, source: str, message: str, index: int, total: int) -> None:
        self._set_item_status(source, "failed")
        self._completed_files += 1
        self.progress.setValue(self._completed_files)
        self._set_status("file_error", name=Path(source).name, message=message)

    @pyqtSlot()
    def on_merge_started(self) -> None:
        self.progress.setValue(len(self.sources))
        self._set_status("merging")

    @pyqtSlot(object, object, object, bool)
    def on_conversion_finished(self, results, errors, merged_result, cancelled) -> None:
        self.running = False
        source_keys = {str(source) for source in self.sources}
        handled = {str(result.source) for result in results}
        handled.update(
            str(path) for path, _message in errors if str(path) in source_keys
        )
        completed_steps = len(handled) + (1 if merged_result else 0)
        self.progress.setValue(completed_steps)
        for source in self.sources:
            if str(source) not in handled:
                self._set_item_status(source, "cancelled" if cancelled else "stopped")
        self._refresh_actions()

        if cancelled:
            self._set_status("job_cancelled", count=len(results))
        elif merged_result and not errors:
            restored = sum(result.images_restored for result in results)
            self._set_status(
                "merge_completed",
                count=len(results),
                name=merged_result.output.name,
                pages=merged_result.page_count,
                _quality_count=restored,
            )
        elif results and not errors:
            restored = sum(result.images_restored for result in results)
            self._set_status(
                "all_completed",
                count=len(results),
                _quality_count=restored,
                _quality_key="quality_note_comma",
            )
        elif results:
            self._set_status(
                "partial_completed",
                success=len(results),
                failed=len(errors),
            )
        else:
            self._set_status("conversion_failed")

        if errors:
            details = "\n\n".join(f"{path.name}\n{message}" for path, message in errors[:5])
            if len(errors) > 5:
                details += "\n\n" + self._t("more_failures", count=len(errors) - 5)
            self._show_warning("partial_failure_title", details)

        reveal_path = merged_result.output if merged_result else None
        if reveal_path is None and results and not self.job_merge:
            reveal_path = results[0].output
        if reveal_path and self.reveal_box.isChecked():
            try:
                if merged_result or len(results) == 1:
                    subprocess.Popen(["explorer.exe", f"/select,{reveal_path}"])
                else:
                    os.startfile(str(reveal_path.parent))
            except OSError:
                pass

        self.worker = None
        self.thread = None

    @pyqtSlot()
    def on_thread_finished(self) -> None:
        if self.close_after_cancel:
            self.close_after_cancel = False
            QTimer.singleShot(0, self.close)

    def closeEvent(self, event: QCloseEvent) -> None:
        if self.running:
            dialog = QMessageBox(self)
            dialog.setIcon(QMessageBox.Icon.Question)
            dialog.setWindowTitle(self._t("conversion_in_progress_title"))
            dialog.setText(self._t("conversion_in_progress_message"))
            stop_button = dialog.addButton(
                self._t("stop_remaining"),
                QMessageBox.ButtonRole.AcceptRole,
            )
            keep_button = dialog.addButton(
                self._t("keep_running"),
                QMessageBox.ButtonRole.RejectRole,
            )
            dialog.setDefaultButton(keep_button)
            dialog.exec()
            if dialog.clickedButton() is stop_button:
                self.close_after_cancel = True
                self.cancel_job()
            event.ignore()
            return
        event.accept()


STYLE = """
QWidget#root {
    background: #f7f8fb;
    color: #111827;
}
QScrollArea#contentScroll {
    background: #f7f8fb;
    border: none;
}
QLabel#eyebrow {
    color: #64748b;
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 1px;
}
QLabel#title {
    color: #111827;
    font-size: 28px;
    font-weight: 700;
}
QLabel#subtitle, QLabel#muted, QLabel#statusLabel {
    color: #64748b;
    font-size: 13px;
}
QLabel#pathPreview {
    color: #334155;
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    border-radius: 6px;
    padding: 5px 8px;
    font-size: 12px;
}
QLabel#pathPreview[hasPath="false"] {
    color: #64748b;
}
QLabel#sectionTitle {
    color: #111827;
    font-size: 13px;
    font-weight: 700;
}
QFrame#dropPanel {
    background: #ffffff;
    border: 2px dashed #b9c4d2;
    border-radius: 12px;
}
QFrame#dropPanel[dragActive="true"] {
    background: #eff6ff;
    border-color: #2563eb;
}
QLabel#dropTitle {
    color: #111827;
    font-size: 18px;
    font-weight: 600;
    border: none;
}
QLabel#dropHint {
    color: #64748b;
    font-size: 12px;
    border: none;
}
QListWidget#fileList, QLineEdit {
    background: #ffffff;
    border: 1px solid #d9e0e9;
    border-radius: 8px;
    padding: 8px;
    selection-background-color: #dbeafe;
    selection-color: #111827;
}
QComboBox#languageCombo, QComboBox#themeCombo {
    color: #334155;
    background: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 7px;
    min-height: 26px;
    padding: 2px 8px;
}
QComboBox#languageCombo:hover, QComboBox#themeCombo:hover { border-color: #94a3b8; }
QComboBox#languageCombo:focus, QComboBox#themeCombo:focus { border: 2px solid #2563eb; }
QComboBox#languageCombo QAbstractItemView,
QComboBox#themeCombo QAbstractItemView,
QAbstractItemView#languagePopup,
QAbstractItemView#themePopup {
    color: #111827;
    background: #ffffff;
    border: 1px solid #cbd5e1;
    selection-background-color: #dbeafe;
    selection-color: #111827;
    outline: 0;
}
QComboBox#languageCombo QAbstractItemView::item,
QComboBox#themeCombo QAbstractItemView::item,
QAbstractItemView#languagePopup::item,
QAbstractItemView#themePopup::item {
    color: #111827;
    background: #ffffff;
    padding: 7px 10px;
    min-height: 24px;
}
QComboBox#languageCombo QAbstractItemView::item:hover,
QComboBox#themeCombo QAbstractItemView::item:hover,
QAbstractItemView#languagePopup::item:hover,
QAbstractItemView#themePopup::item:hover {
    color: #111827;
    background: #eff6ff;
}
QComboBox#languageCombo QAbstractItemView::item:selected,
QComboBox#themeCombo QAbstractItemView::item:selected,
QAbstractItemView#languagePopup::item:selected,
QAbstractItemView#themePopup::item:selected {
    color: #111827;
    background: #dbeafe;
}
QListWidget#fileList:focus, QLineEdit:focus {
    border: 2px solid #2563eb;
}
QListWidget#fileList::item {
    color: #111827;
    background: #ffffff;
    padding: 7px 6px;
    min-height: 16px;
}
QListWidget#fileList::item:hover { background: #f8fafc; }
QListWidget#fileList::item:selected {
    color: #111827;
    background: #dbeafe;
}
QListWidget#fileList::item:selected:active {
    color: #111827;
    background: #bfdbfe;
}
QPushButton {
    min-height: 28px;
    padding: 6px 14px;
    border-radius: 8px;
    font-weight: 600;
}
QPushButton#primaryButton {
    color: white;
    background: #2563eb;
    border: 1px solid #2563eb;
}
QPushButton#primaryButton:hover { background: #1d4ed8; }
QPushButton#primaryButton:pressed { background: #1e40af; }
QPushButton#primaryButton:disabled {
    color: #f8fafc;
    background: #9fb3cf;
    border-color: #9fb3cf;
}
QPushButton#secondaryButton {
    color: #1f2937;
    background: #ffffff;
    border: 1px solid #cbd5e1;
}
QPushButton#secondaryButton:hover { background: #f1f5f9; }
QPushButton#secondaryButton:pressed { background: #e2e8f0; }
QPushButton#quietButton {
    color: #2563eb;
    background: transparent;
    border: none;
    padding: 3px 7px;
}
QPushButton#quietButton:hover { background: #eff6ff; }
QPushButton#quietButton:disabled { color: #94a3b8; }
QPushButton:focus { border: 2px solid #1d4ed8; }
QLabel#wordStatus[available="true"] { color: #15803d; font-size: 12px; }
QLabel#wordStatus[available="false"] { color: #b91c1c; font-size: 12px; }
QProgressBar {
    background: #e2e8f0;
    border: none;
    border-radius: 4px;
    min-height: 8px;
    max-height: 8px;
}
QProgressBar::chunk {
    background: #2563eb;
    border-radius: 4px;
}
QCheckBox { color: #334155; spacing: 7px; }
QToolTip {
    color: #f8fafc;
    background: #1e293b;
    border: 1px solid #334155;
    padding: 5px;
}
"""


THEME_DARK_STYLE = """
QWidget#root, QScrollArea#contentScroll {
    background: #111827;
    color: #f8fafc;
}
QLabel#eyebrow, QLabel#subtitle, QLabel#muted, QLabel#statusLabel,
QLabel#dropHint {
    color: #94a3b8;
}
QLabel#title, QLabel#sectionTitle, QLabel#dropTitle {
    color: #f8fafc;
}
QLabel#pathPreview {
    color: #cbd5e1;
    background: #1f2937;
    border-color: #334155;
}
QLabel#pathPreview[hasPath="false"] { color: #94a3b8; }
QFrame#dropPanel {
    background: #1f2937;
    border-color: #475569;
}
QFrame#dropPanel[dragActive="true"] {
    background: #172554;
    border-color: #60a5fa;
}
QListWidget#fileList, QLineEdit {
    color: #f8fafc;
    background: #111827;
    border-color: #334155;
    selection-background-color: #1e3a5f;
    selection-color: #f8fafc;
}
QListWidget#fileList:focus, QLineEdit:focus {
    border-color: #60a5fa;
}
QLineEdit:disabled { color: #64748b; background: #1f2937; }
QComboBox#languageCombo, QComboBox#themeCombo {
    color: #e2e8f0;
    background: #1f2937;
    border-color: #475569;
}
QComboBox#languageCombo:hover, QComboBox#themeCombo:hover {
    border-color: #94a3b8;
}
QComboBox#languageCombo:focus, QComboBox#themeCombo:focus {
    border-color: #60a5fa;
}
QComboBox#languageCombo QAbstractItemView,
QComboBox#themeCombo QAbstractItemView,
QAbstractItemView#languagePopup,
QAbstractItemView#themePopup {
    color: #f8fafc;
    background: #1f2937;
    border-color: #475569;
    selection-background-color: #1e3a5f;
    selection-color: #f8fafc;
}
QComboBox#languageCombo QAbstractItemView::item,
QComboBox#themeCombo QAbstractItemView::item,
QAbstractItemView#languagePopup::item,
QAbstractItemView#themePopup::item {
    color: #f8fafc;
    background: #1f2937;
}
QComboBox#languageCombo QAbstractItemView::item:hover,
QComboBox#themeCombo QAbstractItemView::item:hover,
QAbstractItemView#languagePopup::item:hover,
QAbstractItemView#themePopup::item:hover {
    background: #334155;
}
QComboBox#languageCombo QAbstractItemView::item:selected,
QComboBox#themeCombo QAbstractItemView::item:selected,
QAbstractItemView#languagePopup::item:selected,
QAbstractItemView#themePopup::item:selected {
    background: #1e3a5f;
}
QListWidget#fileList::item {
    color: #f8fafc;
    background: #111827;
}
QListWidget#fileList::item:hover { background: #1f2937; }
QListWidget#fileList::item:selected { background: #1e3a5f; }
QListWidget#fileList::item:selected:active { background: #1d4ed8; }
QPushButton#primaryButton:hover { background: #3b82f6; }
QPushButton#primaryButton:pressed { background: #1d4ed8; }
QPushButton#primaryButton:disabled {
    color: #cbd5e1;
    background: #334155;
    border-color: #334155;
}
QPushButton#secondaryButton {
    color: #e2e8f0;
    background: #1f2937;
    border-color: #475569;
}
QPushButton#secondaryButton:hover { background: #334155; }
QPushButton#secondaryButton:pressed { background: #475569; }
QPushButton#quietButton { color: #60a5fa; }
QPushButton#quietButton:hover { background: #172554; }
QPushButton#quietButton:disabled { color: #64748b; }
QPushButton:focus { border-color: #60a5fa; }
QLabel#wordStatus[available="true"] { color: #4ade80; }
QLabel#wordStatus[available="false"] { color: #f87171; }
QProgressBar { background: #334155; }
QProgressBar::chunk { background: #3b82f6; }
QCheckBox { color: #cbd5e1; }
QToolTip {
    color: #f8fafc;
    background: #0f172a;
    border-color: #475569;
}
QScrollBar:vertical {
    background: #111827;
    width: 12px;
    margin: 2px;
}
QScrollBar::handle:vertical {
    background: #475569;
    min-height: 32px;
    border-radius: 5px;
}
QScrollBar::handle:vertical:hover { background: #64748b; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0;
    background: transparent;
}
QMessageBox, QDialog { background: #111827; color: #f8fafc; }
"""


def stylesheet_for_theme(theme: str) -> str:
    return STYLE + (THEME_DARK_STYLE if normalize_theme(theme) == "dark" else "")


def configure_application(app: QApplication) -> None:
    app.setApplicationName(APP_NAME)
    app.setApplicationDisplayName(APP_NAME)
    app.setApplicationVersion(APP_VERSION)
    app.setOrganizationName("Local")
    # Explicitly load the installed Windows CJK UI font. Qt's offscreen and
    # packaged runtimes do not always discover system fallback fonts reliably.
    windows_font = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts" / "msyh.ttc"
    if windows_font.is_file():
        QFontDatabase.addApplicationFont(str(windows_font))
    app.setFont(QFont("Microsoft YaHei UI", 10))
    app.setStyleSheet(stylesheet_for_theme("light"))


def main() -> int:
    app = QApplication(sys.argv)
    configure_application(app)
    window = MainWindow()
    window.show()
    return app.exec()


def diagnostic_convert(arguments: list[str]) -> int:
    """Run the packaged conversion path without creating a GUI window."""
    if len(arguments) != 3:
        return 2
    source, output, report = map(Path, arguments)
    try:
        result = convert_docx(
            source,
            output,
            overwrite=True,
            timeout=180,
            language=normalize_language(
                os.environ.get("DOCXPDF_LANGUAGE", DEFAULT_LANGUAGE)
            ),
        )
        message = f"OK\n{result.output}\n{result.size_bytes}\n"
        exit_code = 0
    except Exception as exc:
        message = f"ERROR\n{type(exc).__name__}: {exc}\n"
        exit_code = 1
    try:
        report.parent.mkdir(parents=True, exist_ok=True)
        report.write_text(message, encoding="utf-8")
    except OSError:
        return 3
    return exit_code


def diagnostic_ui(arguments: list[str]) -> int:
    """Launch the packaged UI briefly, then exit normally and write a report."""
    if len(arguments) != 1:
        return 2
    report = Path(arguments[0])
    try:
        qt_app = QApplication([sys.argv[0]])
        configure_application(qt_app)
        window = MainWindow()
        window.show()
        QTimer.singleShot(1000, qt_app.quit)
        exit_code = qt_app.exec()
        message = f"OK\nQt exit code: {exit_code}\n"
    except Exception as exc:
        message = f"ERROR\n{type(exc).__name__}: {exc}\n"
        exit_code = 1
    try:
        report.parent.mkdir(parents=True, exist_ok=True)
        report.write_text(message, encoding="utf-8")
    except OSError:
        return 3
    return exit_code


if __name__ == "__main__":
    if len(sys.argv) >= 2 and sys.argv[1] == "--diagnostic-convert":
        raise SystemExit(diagnostic_convert(sys.argv[2:]))
    if len(sys.argv) >= 2 and sys.argv[1] == "--diagnostic-ui":
        raise SystemExit(diagnostic_ui(sys.argv[2:]))
    raise SystemExit(main())
