# DocxPDF Fidelity Converter for Windows

English | [简体中文](README.md)

DocxPDF is a local desktop app that uses an installed copy of Microsoft Word to convert one or more DOCX files to PDF. It can also merge the converted documents in list order. Source documents and converted files are never uploaded.

## Features

- Select or drag and drop multiple `.docx` files
- Create a separate PDF for each DOCX
- Reorder files by dragging, then convert and merge them into one PDF
- Choose a custom output folder
- Protect existing files by adding `-1`, `-2`, and so on unless overwrite is enabled
- Background processing, per-file states, cancellation after the current file, and File Explorer reveal
- Unicode, spaces, and long filenames supported
- Instant 中文/English switching; follows the Windows language on first launch and remembers the user's choice

## Fidelity approach

1. An isolated, hidden Microsoft Word instance opens each DOCX read-only and exports it using Word's print-quality PDF renderer.
2. Word's `OptimizeForImageQuality` option is enabled.
3. Word for Windows may still downsample some images. After export, DocxPDF compares PDF images with the original DOCX media. Only unambiguous, high-confidence matches are restored as lossless Flate image streams. Page size, content stream, placement, and geometry are left unchanged.
4. Merging copies PDF pages without rasterizing them or re-encoding embedded image streams.

“Original image restored” means that the resulting raster pixels match the image stored in the DOCX. It does not mean that the PDF compression bytes are identical to the original image file. Images that cannot be matched safely are left as Word exported them. WMF/EMF, linked images, unusual high-bit-depth formats, and some special effects remain under Word's control.

## Requirements

- 64-bit Windows 10 or Windows 11
- An installed and activated desktop version of Microsoft Word
- Python 3.10+ only when running or building from source

## Run from source

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe app.py
```

## Test

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

The 15 tests cover both interface languages and translation completeness, Word COM arguments and cleanup, atomic output, failures and cancellation, original-pixel restoration, merge order, and preservation of image streams during merge.

## Build the Windows app

```powershell
powershell -ExecutionPolicy Bypass -File .\build_windows.ps1
```

The script runs the tests first, then creates:

- `dist\DocxPDF\DocxPDF.exe`
- `dist\DocxPDF-Windows-x64.zip`

The current build is unsigned and does not include an installer. Windows SmartScreen may warn on first launch on another computer. Add Authenticode signing and an installer before public distribution.

## Known boundaries

- Layout ultimately depends on the local Word version, installed fonts, linked resources, dynamic fields, and the document itself.
- Password-protected or automation-restricted documents cannot be converted.
- Merge prioritizes page order and visual content; advanced cross-document structures such as bookmarks, forms, and attachments are not guaranteed.
- This app focuses on DOCX conversion and merge. It does not include Acrobat-style PDF editing, OCR, signing, or form authoring.
