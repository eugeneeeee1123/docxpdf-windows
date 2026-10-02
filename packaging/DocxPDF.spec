from pathlib import Path

project_dir = Path(SPECPATH).resolve().parent

a = Analysis(
    [str(project_dir / "app.py")],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=["pythoncom", "pywintypes", "win32com.client"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

# Qt uses the ICU API provided by Windows. Other tools on the host PATH can
# provide an icuuc.dll with version-suffixed exports that are incompatible with
# Qt. Use the supported Windows 10/11 ICU and C runtimes instead of copying
# these operating-system DLLs from the build host's unrelated applications.
system_runtime_names = {"icuuc.dll", "icuin.dll", "icudt.dll", "ucrtbase.dll"}
a.binaries = [
    entry for entry in a.binaries
    if Path(entry[0]).name.lower() not in system_runtime_names
]

pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="DocxPDF",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="DocxPDF",
)
