# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

PPTX Compressor - A Windows desktop application for compressing videos and images in PowerPoint files (PPTX). The application uses tkinter for GUI and processes PPTX files as ZIP archives.

## Commands

### Run the application
```bash
pip install pillow
python compress_pptx.py
```

### Build executable
```powershell
# Using PowerShell script (recommended)
.\build.ps1

# Or manually
pip install pyinstaller pillow
python -m PyInstaller --onefile --windowed --add-data "ffmpeg.exe;." --name "PPTXCompressor" --clean compress_pptx.py
```

## Architecture

### Single-file application (`compress_pptx.py`)

**PPTXCompressor class** handles:
- GUI setup (tkinter)
- File selection and parameter configuration
- Compression pipeline in background thread

### PPTX Processing Flow

1. Extract PPTX (ZIP archive) to temp directory
2. Find media files in `ppt/media/` subdirectory
3. Compress images (PNG/JPEG/WebP) using PIL
4. Compress videos (MP4/AVI/MOV/WMV/MKV/FLV/WebM) via ffmpeg
5. Repack to new PPTX file with `_compressed` suffix

### Image Compression Strategy

**PNG**: PIL lossless compression with optimize and compress_level=9

**JPEG**: Quality optimization (95/90/85/80), remove metadata, progressive encoding

**WebP**: Lossless optimization

**GIF**: Preserved unchanged (to protect animations)

### Video Compression

Uses ffmpeg with H.264 codec:
- CRF value (18-51): Higher = more compression, lower quality
- Preset: Slower = better compression ratio
- Audio: AAC encoding

### External Tools

Located in project root:
- `ffmpeg.exe` - Required for video compression

## Key Implementation Details

- `sys.frozen` check distinguishes packaged exe vs development mode
- `sys._MEIPASS` provides temp path for bundled resources in packaged exe
- PPTX files are ZIP archives - media located at `ppt/media/`
- All compression operations preserve original if result is larger
- GUI updates via `root.after(0, callback)` for thread safety
