# DocxPDF 保真转换器（Windows）

[English](README.en.md) | 简体中文

DocxPDF 是一个本地桌面工具，通过已安装的 Microsoft Word 将一个或多个 DOCX 转成 PDF，也可以按列表顺序合并为一个 PDF。源文档与转换结果不会上传到网络。

## 下载已打包版本

请到 [GitHub Releases](https://github.com/eugeneeeee1123/docxpdf-windows/releases) 下载最新的 `DocxPDF-Windows-x64.zip`。解压后运行 `DocxPDF\DocxPDF.exe`，不需要安装 Python；电脑仍需安装并激活桌面版 Microsoft Word。

## 功能

- 拖放或批量选择多个 `.docx`
- 每个 DOCX 单独生成 PDF
- 按可拖动的列表顺序转换并合并
- 自定义输出文件夹
- 默认不覆盖同名文件，自动使用 `-1`、`-2` 等安全文件名
- 后台处理、逐文件状态、取消批次，以及完成后定位文件
- 支持中文、空格和较长文件名路径
- 中文/English 界面即时切换，并记住用户选择

## 图片与排版保真策略

1. 使用隔离、隐藏的 Microsoft Word 实例，以只读方式打开 DOCX，并调用 Word 原生 PDF 导出器。
2. 开启 `OptimizeForImageQuality` 高质量导出选项。
3. 对导出的 PDF 图片与 DOCX 内原始图片进行安全匹配；只有高置信度匹配才会以 PDF 无损 Flate 图像流恢复原始像素，页面尺寸、位置和排版几何不改变。
4. 合并阶段只复制 PDF 页面，不栅格化页面，也不重新编码已有图片流。

这里的“原始图片恢复”表示 PDF 中的栅格像素与 DOCX 中的图片一致，并不表示 PDF 压缩字节与原始 PNG/JPEG 文件完全相同。无法安全匹配的特殊图片会保留 Word 的导出结果。

## 要求

- 64 位 Windows 10 或 Windows 11
- 已安装并激活的桌面版 Microsoft Word
- 只有从源码运行或重新构建时才需要 Python 3.10+

## 从源码运行

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe app.py
```

## 测试

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

当前测试覆盖中英文界面、翻译完整性、Word COM 参数与清理、原子输出、错误与取消、原始像素恢复、合并顺序，以及合并时图片流不被重新编码。

## 构建 Windows 程序

```powershell
powershell -ExecutionPolicy Bypass -File .\build_windows.ps1
```

构建脚本会先运行测试，再生成：

- `dist\DocxPDF\DocxPDF.exe`
- `dist\DocxPDF-Windows-x64.zip`

当前产物未做代码签名，也没有安装器；在其他电脑首次运行时可能出现 Windows SmartScreen 提示。

## 已知边界

- 排版最终由本机 Word 版本、字体、链接资源、动态字段和文档本身决定。
- 密码保护或禁止自动化导出的文档无法转换。
- 合并主要保证页面顺序与视觉内容；跨文档书签、表单、附件等高级 PDF 结构不保证完整保留。
- 本工具专注于 DOCX 转换与合并，不包含 Acrobat 的 PDF 编辑、OCR、签名或表单制作功能。
