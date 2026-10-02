param(
    [string]$DistDirectoryName = "dist",
    [string]$ZipName = "DocxPDF-Windows-x64.zip"
)

$ErrorActionPreference = "Stop"

$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $ProjectDir

$Python = Join-Path $ProjectDir ".venv\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $Python)) {
    $Launcher = Get-Command py -ErrorAction SilentlyContinue
    if ($null -eq $Launcher) {
        throw "Python launcher not found. Install Python 3.10 or newer."
    }
    & $Launcher.Source -3 -m venv (Join-Path $ProjectDir ".venv")
}

& $Python -m pip install -r (Join-Path $ProjectDir "requirements.txt")
if ($LASTEXITCODE -ne 0) {
    throw "Dependency installation failed."
}
& $Python -m unittest discover -s (Join-Path $ProjectDir "tests") -v
if ($LASTEXITCODE -ne 0) {
    throw "Tests failed; packaging stopped."
}

$DistPath = Join-Path $ProjectDir $DistDirectoryName
& $Python -m PyInstaller `
    --noconfirm `
    --clean `
    --distpath $DistPath `
    --workpath (Join-Path $ProjectDir "build") `
    (Join-Path $ProjectDir "packaging\DocxPDF.spec")
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller failed; no ZIP was created."
}

$AppDir = Join-Path $DistPath "DocxPDF"
$ZipPath = Join-Path $DistPath $ZipName
if (-not (Test-Path -LiteralPath (Join-Path $AppDir "DocxPDF.exe"))) {
    throw "PyInstaller did not create DocxPDF.exe."
}

Copy-Item -LiteralPath (Join-Path $ProjectDir "README.md"), (Join-Path $ProjectDir "README.en.md"), (Join-Path $ProjectDir "LICENSE") -Destination $AppDir
Compress-Archive -Path $AppDir -DestinationPath $ZipPath -Force
Write-Host "Created: $AppDir"
Write-Host "Created: $ZipPath"
