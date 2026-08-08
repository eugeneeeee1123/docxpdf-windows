from __future__ import annotations

from typing import Any


DEFAULT_LANGUAGE = "zh"
SUPPORTED_LANGUAGES = ("zh", "en")


STRINGS: dict[str, dict[str, str]] = {
    # Application shell and controls.
    "app_name": {"zh": "DocxPDF 保真转换器", "en": "DocxPDF Fidelity Converter"},
    "eyebrow": {
        "zh": "WINDOWS · MICROSOFT WORD 原生引擎",
        "en": "WINDOWS · NATIVE MICROSOFT WORD ENGINE",
    },
    "title": {"zh": "DOCX → PDF，保持原始版式", "en": "DOCX → PDF, faithful to the original"},
    "subtitle": {
        "zh": "Word 以打印质量原生排版；导出后恢复可安全匹配的原始图片像素。",
        "en": "Word renders at print quality, then safely matched images are restored to their original pixels.",
    },
    "drop_title": {"zh": "拖放 DOCX 到这里", "en": "Drop DOCX files here"},
    "drop_hint": {
        "zh": "点击选择或拖放多个文件 · 支持中文及空格路径",
        "en": "Click to choose or drop multiple files · Unicode and spaced paths supported",
    },
    "files_section": {"zh": "待转换文件", "en": "Files to convert"},
    "file_count": {"zh": "{count} 个", "en": "{count} file(s)"},
    "order_hint": {"zh": "拖动可调整转换与合并顺序", "en": "Drag to change conversion and merge order"},
    "remove_selected": {"zh": "移除所选", "en": "Remove selected"},
    "clear": {"zh": "清空", "en": "Clear"},
    "file_list_accessible": {"zh": "待转换 DOCX 文件列表", "en": "DOCX files to convert"},
    "output_section": {"zh": "输出位置", "en": "Output location"},
    "output_placeholder": {"zh": "请选择保存 PDF 的文件夹", "en": "Choose a folder for the PDFs"},
    "output_accessible": {"zh": "PDF 输出文件夹", "en": "PDF output folder"},
    "browse": {"zh": "浏览…", "en": "Browse…"},
    "browse_tooltip": {"zh": "选择 PDF 保存位置", "en": "Choose where to save the PDFs"},
    "overwrite": {"zh": "覆盖同名 PDF", "en": "Overwrite PDFs with the same name"},
    "overwrite_tooltip": {
        "zh": "未勾选时会自动添加 -1、-2，避免覆盖已有文件。",
        "en": "When off, -1, -2, and so on are added to protect existing files.",
    },
    "reveal": {"zh": "完成后在文件资源管理器中显示", "en": "Show in File Explorer when finished"},
    "output_hint": {
        "zh": "独立转换沿用原文件名；合并会按列表顺序生成“合并结果.pdf”。",
        "en": "Individual PDFs keep their source names; merge creates “Merged.pdf” in list order.",
    },
    "cancel_task": {"zh": "取消任务", "en": "Cancel task"},
    "cancel_tooltip": {
        "zh": "完成当前文件后停止剩余任务",
        "en": "Stop the remaining files after the current one finishes",
    },
    "convert_individual": {"zh": "转换为独立 PDF", "en": "Convert to separate PDFs"},
    "convert_individual_tooltip": {
        "zh": "每个 DOCX 分别生成一个 PDF",
        "en": "Create one PDF for each DOCX file",
    },
    "convert_merge": {"zh": "转换并合并", "en": "Convert and merge"},
    "convert_merge_tooltip": {
        "zh": "按列表顺序生成一个 PDF；每个 DOCX 仍由 Word 独立渲染。",
        "en": "Create one PDF in list order; Word still renders each DOCX separately.",
    },
    "progress_accessible": {"zh": "转换进度", "en": "Conversion progress"},
    "language_accessible": {"zh": "界面语言", "en": "Interface language"},
    "language_tooltip": {"zh": "切换界面语言", "en": "Change interface language"},
    # Word and file states.
    "word_ready": {"zh": "● Microsoft Word{version} 已就绪", "en": "● Microsoft Word{version} is ready"},
    "word_missing": {"zh": "● 未找到 Microsoft Word", "en": "● Microsoft Word not found"},
    "word_missing_tooltip": {
        "zh": "请安装并激活桌面版 Microsoft Word",
        "en": "Install and activate the desktop version of Microsoft Word",
    },
    "item_waiting": {"zh": "等待", "en": "Waiting"},
    "item_processing": {"zh": "处理中", "en": "Processing"},
    "item_complete": {"zh": "完成", "en": "Done"},
    "item_complete_quality": {
        "zh": "完成 · 原图保真 {count} 张",
        "en": "Done · {count} original image(s) restored",
    },
    "item_failed": {"zh": "失败", "en": "Failed"},
    "item_cancelled": {"zh": "已取消", "en": "Cancelled"},
    "item_stopped": {"zh": "已停止", "en": "Stopped"},
    "quality_note": {
        "zh": " · 原图保真 {count} 张",
        "en": " · {count} original image(s) restored",
    },
    "quality_note_comma": {
        "zh": "，原图保真 {count} 张",
        "en": ", with {count} original image(s) restored",
    },
    # Status messages and dialogs.
    "select_files": {"zh": "请选择 DOCX 文件。", "en": "Choose one or more DOCX files."},
    "files_added": {
        "zh": "已加入 {added} 个文件。转换时不会修改原 DOCX。",
        "en": "Added {added} file(s). The original DOCX files will not be modified.",
    },
    "files_added_skipped": {
        "zh": "已加入 {added} 个文件，忽略 {skipped} 个无效或重复文件。转换时不会修改原 DOCX。",
        "en": "Added {added} file(s); skipped {skipped} invalid or duplicate file(s). Original DOCX files will not be modified.",
    },
    "no_valid_files": {
        "zh": "未添加文件：请选择有效且未重复的 .docx 文档。",
        "en": "No files added. Choose valid, non-duplicate .docx documents.",
    },
    "starting_word": {
        "zh": "正在启动独立的 Microsoft Word 后台实例…",
        "en": "Starting an isolated Microsoft Word background instance…",
    },
    "starting_merge": {
        "zh": "正在启动 Microsoft Word；转换完成后会按列表顺序合并…",
        "en": "Starting Microsoft Word; the PDFs will be merged in list order after conversion…",
    },
    "cancel_requested": {
        "zh": "已请求取消；将在当前文件完成后停止剩余任务。",
        "en": "Cancellation requested. Remaining files will stop after the current file finishes.",
    },
    "converting": {
        "zh": "正在转换 {index}/{total}：{name}",
        "en": "Converting {index}/{total}: {name}",
    },
    "file_completed": {
        "zh": "已完成 {index}/{total}：{name} · {size_mb:.1f} MB{quality_note}",
        "en": "Completed {index}/{total}: {name} · {size_mb:.1f} MB{quality_note}",
    },
    "file_error": {"zh": "{name}：{message}", "en": "{name}: {message}"},
    "merging": {
        "zh": "正在合并 PDF 页面，不会重新渲染或压缩图片…",
        "en": "Merging PDF pages without rerendering or recompressing images…",
    },
    "job_cancelled": {
        "zh": "任务已取消；取消前完成 {count} 个文件。",
        "en": "Task cancelled; {count} file(s) finished before cancellation.",
    },
    "merge_completed": {
        "zh": "合并完成：{count} 个 DOCX → {name}（{pages} 页）{quality_note}",
        "en": "Merge complete: {count} DOCX file(s) → {name} ({pages} pages){quality_note}",
    },
    "all_completed": {
        "zh": "全部完成：已生成 {count} 个 PDF{quality_note}。",
        "en": "All done: created {count} PDF file(s){quality_note}.",
    },
    "partial_completed": {
        "zh": "转换结束：成功 {success} 个，失败 {failed} 个。",
        "en": "Conversion finished: {success} succeeded, {failed} failed.",
    },
    "conversion_failed": {
        "zh": "转换失败，原 DOCX 未被修改。",
        "en": "Conversion failed. The original DOCX files were not modified.",
    },
    "output_unavailable_title": {"zh": "无法使用输出位置", "en": "Output location unavailable"},
    "partial_failure_title": {"zh": "部分文件未转换", "en": "Some files were not converted"},
    "more_failures": {"zh": "另有 {count} 个失败文件。", "en": "There are {count} additional failed file(s)."},
    "conversion_in_progress_title": {"zh": "转换仍在进行", "en": "Conversion in progress"},
    "conversion_in_progress_message": {
        "zh": "要停止剩余任务吗？当前文件完成后即可退出。",
        "en": "Stop the remaining files? The app will exit after the current file finishes.",
    },
    "stop_remaining": {"zh": "停止剩余任务", "en": "Stop remaining files"},
    "keep_running": {"zh": "继续转换", "en": "Keep converting"},
    "ok": {"zh": "确定", "en": "OK"},
    "choose_docs_title": {"zh": "选择 DOCX 文件", "en": "Choose DOCX files"},
    "word_documents_filter": {"zh": "Word 文档 (*.docx)", "en": "Word documents (*.docx)"},
    "choose_output_title": {"zh": "选择输出文件夹", "en": "Choose output folder"},
    # Worker-only labels and errors.
    "merge_result_label": {"zh": "合并结果", "en": "Merged result"},
    "word_label": {"zh": "Microsoft Word", "en": "Microsoft Word"},
    "task_label": {"zh": "任务", "en": "Task"},
    "output_dir_worker_error": {
        "zh": "无法创建输出目录：{error}",
        "en": "Unable to create the output folder: {error}",
    },
    "unexpected_error": {"zh": "未预期错误：{error}", "en": "Unexpected error: {error}"},
    # Converter errors.
    "source_not_found": {"zh": "找不到文件：{path}", "en": "File not found: {path}"},
    "source_extension": {"zh": "只支持 .docx 文件：{name}", "en": "Only .docx files are supported: {name}"},
    "source_lock": {
        "zh": "这是 Word 临时锁定文件，不能转换：{name}",
        "en": "This is a temporary Word lock file and cannot be converted: {name}",
    },
    "output_dir_create": {"zh": "无法创建输出文件夹：{error}", "en": "Unable to create output folder: {error}"},
    "output_not_dir": {"zh": "输出路径不是文件夹：{path}", "en": "The output path is not a folder: {path}"},
    "too_many_names": {
        "zh": "输出文件夹中同名文件过多：{path}",
        "en": "Too many files with the same base name in the output folder: {path}",
    },
    "pdf_read": {"zh": "无法读取 PDF：{error}", "en": "Unable to read the PDF: {error}"},
    "pdf_invalid": {
        "zh": "Word 返回成功，但输出文件不是有效的 PDF。",
        "en": "Word reported success, but the output is not a valid PDF.",
    },
    "pywin32_missing": {
        "zh": "缺少 Windows Word 转换组件 pywin32。请执行 pip install -r requirements.txt。",
        "en": "The pywin32 Word integration is missing. Run pip install -r requirements.txt.",
    },
    "word_no_detail": {
        "zh": "Microsoft Word 未返回错误详情。",
        "en": "Microsoft Word did not provide error details.",
    },
    "word_start_failed": {
        "zh": "无法启动 Microsoft Word。请确认桌面版 Word 已安装并激活；若仍失败，请在 Office 中执行“快速修复”。",
        "en": "Microsoft Word could not start. Confirm that desktop Word is installed and activated; if it still fails, run Office Quick Repair.",
    },
    "word_busy": {
        "zh": "Microsoft Word 正忙或连接已中断。请关闭 Word 中的弹窗，再重试本批任务。",
        "en": "Microsoft Word is busy or disconnected. Close any Word dialog boxes, then retry the batch.",
    },
    "document_protected": {
        "zh": "文档受密码或保护限制，Microsoft Word 无法自动导出。",
        "en": "The document is password-protected or restricted, so Microsoft Word cannot export it automatically.",
    },
    "word_permission": {
        "zh": "Microsoft Word 无法访问源文件或输出文件夹。请检查权限后重试。",
        "en": "Microsoft Word cannot access the source file or output folder. Check permissions and try again.",
    },
    "word_export_failed": {"zh": "Microsoft Word 导出失败：{detail}", "en": "Microsoft Word export failed: {detail}"},
    "pillow_restore_install": {
        "zh": "原图保真处理需要 Pillow。请执行 pip install -r requirements.txt。",
        "en": "Original-image restoration requires Pillow. Run pip install -r requirements.txt.",
    },
    "pillow_restore_required": {"zh": "原图保真处理需要 Pillow。", "en": "Original-image restoration requires Pillow."},
    "pypdf_restore_required": {"zh": "原图保真处理需要 pypdf。", "en": "Original-image restoration requires pypdf."},
    "restore_failed": {"zh": "原图保真处理失败：{error}", "en": "Original-image restoration failed: {error}"},
    "windows_only": {"zh": "当前版本只支持 Windows。", "en": "This version supports Windows only."},
    "session_not_started": {
        "zh": "Microsoft Word 转换会话尚未启动。",
        "en": "The Microsoft Word conversion session has not started.",
    },
    "output_pdf_extension": {"zh": "输出文件必须使用 .pdf 扩展名。", "en": "The output file must use the .pdf extension."},
    "same_input_output": {"zh": "输入和输出文件不能相同。", "en": "The input and output files cannot be the same."},
    "output_exists": {"zh": "输出文件已存在：{name}", "en": "The output file already exists: {name}"},
    "word_quality_unsupported": {
        "zh": "当前 Microsoft Word 不支持原图质量 PDF 导出。请更新 Word 后重试。",
        "en": "This Microsoft Word version does not support original-quality PDF export. Update Word and try again.",
    },
    "word_pdf_not_finished": {
        "zh": "Microsoft Word 返回成功，但没有完成 PDF 写入。请检查输出文件夹权限后重试。",
        "en": "Microsoft Word reported success but did not finish writing the PDF. Check output-folder permissions and try again.",
    },
    "output_write_failed": {"zh": "无法写入输出 PDF：{error}", "en": "Unable to write the output PDF: {error}"},
    "final_save_failed": {"zh": "无法保存最终 PDF：{error}", "en": "Unable to save the final PDF: {error}"},
    "merge_two_required": {"zh": "合并至少需要两个 PDF。", "en": "At least two PDFs are required to merge."},
    "merge_pdf_extension": {
        "zh": "合并输出文件必须使用 .pdf 扩展名。",
        "en": "The merged output must use the .pdf extension.",
    },
    "merge_output_exists": {"zh": "合并输出文件已存在：{name}", "en": "The merged output already exists: {name}"},
    "merge_pypdf_required": {
        "zh": "合并功能需要 pypdf，请先执行 pip install -r requirements.txt。",
        "en": "PDF merging requires pypdf. Run pip install -r requirements.txt first.",
    },
    "merge_input_missing": {"zh": "找不到可合并的 PDF：{path}", "en": "PDF to merge not found: {path}"},
    "merge_encrypted": {"zh": "PDF 已加密，无法合并：{name}", "en": "Encrypted PDF cannot be merged: {name}"},
    "merge_failed": {"zh": "PDF 合并失败：{error}", "en": "PDF merge failed: {error}"},
}


def normalize_language(value: Any, default: str = DEFAULT_LANGUAGE) -> str:
    normalized = str(value or "").strip().lower().replace("-", "_")
    if normalized.startswith("zh"):
        return "zh"
    if normalized.startswith("en"):
        return "en"
    return default if default in SUPPORTED_LANGUAGES else DEFAULT_LANGUAGE


def tr(language: str, key: str, **values: Any) -> str:
    language = normalize_language(language)
    template = STRINGS[key][language]
    return template.format(**values)
