# PPTX Compressor - Build Script
# Auto-creates/uses virtual environment

Write-Host "====================================" -ForegroundColor Cyan
Write-Host "PPTX Compressor - Build Script" -ForegroundColor Cyan
Write-Host "====================================" -ForegroundColor Cyan
Write-Host ""

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Check ffmpeg.exe (included in project)
if (-not (Test-Path "ffmpeg.exe")) {
    Write-Host "[WARNING] ffmpeg.exe not found, please download from https://www.gyan.dev/ffmpeg/builds/" -ForegroundColor Yellow
}

# Virtual environment path
$venv_path = "venv"
$python_exe = "$venv_path\Scripts\python.exe"

# Check/create virtual environment
if (-not (Test-Path $python_exe)) {
    Write-Host "[INFO] Creating virtual environment..." -ForegroundColor Yellow
    python -m venv $venv_path
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[ERROR] Failed to create virtual environment" -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
}

# Install dependencies
Write-Host "[INFO] Installing dependencies..." -ForegroundColor Yellow
& $python_exe -m pip install --upgrade pip -r requirements.txt
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] Failed to install dependencies" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "[BUILD] Building with PyInstaller..." -ForegroundColor White
Write-Host ""

# Clean old build
if (Test-Path "build") { Remove-Item -Recurse -Force build }
if (Test-Path "dist") { Remove-Item -Recurse -Force dist }

# Build
& $python_exe -m PyInstaller --onefile --windowed `
    --add-data "ffmpeg.exe;." `
    --name "PPTXCompressor" `
    --clean compress_pptx.py

Write-Host ""
if (Test-Path "dist\PPTXCompressor.exe") {
    $size = (Get-Item "dist\PPTXCompressor.exe").Length / 1MB
    Write-Host "[SUCCESS] Build complete!" -ForegroundColor Green
    Write-Host "Output: dist\PPTXCompressor.exe" -ForegroundColor Green
    Write-Host "Size: $([math]::Round($size, 2)) MB" -ForegroundColor Green
} else {
    Write-Host "[FAILED] Build failed" -ForegroundColor Red
}

Read-Host "Press Enter to exit"
