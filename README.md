# DocxPDF 保真转换器（Windows）

[English](README.en.md) | 简体中文

DocxPDF 是一个本地桌面工具，通过已安装的 Microsoft Word 把一个或多个 DOCX 转成 PDF，也可按列表顺序合并为一个 PDF。源文档与转换结果不会上传到网络。

## 功能

- 拖放或批量选择多个 `.docx`
- 每个 DOCX 单独生成 PDF
- 按可拖动的列表顺序转换并合并
- 自定义输出文件夹
- 默认不覆盖同名文件，自动使用 `-1`、`-2` 等安全文件名
- 后台处理、逐文件状态、完成后定位文件、当前文件结束后取消批次
- 中文、空格及长文件名路径支持
- 中文/English 界面即时切换；首次启动跟随 Windows 语言，之后记住用户选择

## 保真策略

1. 使用独立、隐藏的 Microsoft Word 实例，以只读方式打开 DOCX，并通过 Word 的打印质量 PDF 导出保持原始排版。
2. 启用 Word 的 `OptimizeForImageQuality` 高质量选项。
3. Windows Word 仍可能对部分图片降采样。因此导出后会比较 PDF 图片与 DOCX 内的原始媒体；只有视觉内容和比例达到高置信匹配、且不存在歧义时，才用无损 Flate 图像流恢复原始像素。页面尺寸、内容流及摆放几何不变。
4. 合并只复制 PDF 页面，不栅格化页面，也不重新编码嵌入图片。

“原图保真”表示恢复后的栅格像素与 DOCX 内原图一致，不表示 PDF 内压缩流字节与 DOCX 图片文件完全相同。无法安全匹配的图片不会被盲目替换；WMF/EMF、外链图片、特殊效果及不常见高位深格式仍由 Word 自己处理。

## 系统要求

- Windows 10 或 Windows 11（64 位）
- 已安装并激活的桌面版 Microsoft Word
- 从源码运行或构建时需要 Python 3.10+

## 运行源码

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe app.py
```

## 测试

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

15 项测试覆盖中英文切换与翻译完整性、Word COM 参数与清理、原子输出、失败与取消、原图像素恢复、合并顺序，以及合并时图片流不被重编码。

## 构建 Windows 程序

```powershell
powershell -ExecutionPolicy Bypass -File .\build_windows.ps1
```

构建脚本会先运行测试，再生成：

- `dist\DocxPDF\DocxPDF.exe`
- `dist\DocxPDF-Windows-x64.zip`

当前产物未做代码签名或安装器封装，首次在其他电脑运行时可能出现 Windows SmartScreen 提示。正式公开分发前建议增加 Authenticode 签名与安装器。

## 已知边界

- 排版最终由本机 Word、字体、链接资源、动态字段及文档本身决定；缺失字体可能造成差异。
- 密码保护或阻止自动导出的文档无法转换。
- 合并主要保证页面顺序和视觉内容；跨文件书签、表单、附件等高级 PDF 结构不保证完整保留。
- 这是 DOCX 转换与合并工具，不包含 Acrobat 的 PDF 编辑、OCR、签名或表单编辑功能。
