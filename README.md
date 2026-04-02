# PPTX Compressor

A standalone Windows desktop application for compressing videos and images in PowerPoint files (PPTX) to reduce file size.

一个独立的 Windows 桌面应用程序，用于压缩 PowerPoint 文件（PPTX）中的视频和图片，有效减小文件大小。

## 功能特性 / Features

### 视频压缩 / Video Compression

- 使用 H.264 编码（ffmpeg）
- 支持 MP4、AVI、MOV、WMV、MKV、FLV、WebM 等格式
- 可调节 CRF 值（18-51）控制压缩质量
- 可调节编码速度（preset）
- 可调节音频码率

### 图片压缩 / Image Compression

- **PNG 压缩**：使用 PIL 进行无损压缩
- **JPEG 压缩**：自动质量优化，移除元数据
- **WebP 压缩**：无损优化
- **GIF**：保持原样（保护动画）
- 支持大图自动缩放
- 可设置压缩阈值

### 其他特性 / Other Features

- 简洁的图形界面
- 实时进度显示
- 详细的压缩统计
- 自动命名输出文件（避免覆盖）

## 截图 / Screenshot

```
┌─────────────────────────────────────────────────────┐
│ PPTX Compressor                                     │
├─────────────────────────────────────────────────────┤
│ 【File Selection】                                  │
│ [C:\demo\presentation.pptx          ] [Browse]     │
├─────────────────────────────────────────────────────┤
│ 【Advanced Options】                                │
│ CRF (18-51):        [28] Higher = more compression │
│ Preset:             [slow ▼] Slower = better ratio │
│ Audio Bitrate:      [128k]                          │
├─────────────────────────────────────────────────────┤
│ 【Image Compression】                               │
│ Size Threshold (KB): [500] Images > this compressed│
│ Max Width (px):     [1920] Oversized scaled down   │
├─────────────────────────────────────────────────────┤
│                  [Start Compression]                │
├─────────────────────────────────────────────────────┤
│ 【Progress】                                        │
│ ████████████████████████████░░░░░░░░░ 70%          │
│ Compressing video: video2.mp4 (2/3)                │
├─────────────────────────────────────────────────────┤
│ 【Result】                                          │
│ Compression completed!                              │
│                                                      │
│ 【Video】                                           │
│   Original size: 85.32 MB                          │
│   Compressed: 21.45 MB                             │
│   Ratio: 74.9%                                     │
│   Total: 3 video file(s)                           │
│                                                      │
│ 【Image】                                           │
│   Original size: 12.50 MB                          │
│   Compressed: 3.20 MB                              │
│   Ratio: 74.4%                                     │
│   Processed 15 image(s), 12 compressed             │
│                                                      │
│ 【Total】                                           │
│   Compression ratio: 74.8%                         │
│                                                      │
│ Output: presentation_compressed.pptx               │
└─────────────────────────────────────────────────────┘
```

## 下载与安装 / Download & Installation

### 方式一：直接使用打包好的程序 / Option 1: Use Pre-built Package

1. 从 [Releases](../../releases) 下载最新版本的 `PPTXCompressor.exe`
2. 双击运行即可

### 方式二：自行打包 / Option 2: Build from Source

#### 前置要求 / Prerequisites

- Python 3.10+（推荐 Python 3.11 或 3.12）
- Windows 系统

#### 步骤 / Steps

1. **克隆仓库 / Clone Repository**

```bash
git clone https://github.com/your-repo/ppt-compress.git
cd ppt-compress
```

2. **确保 ffmpeg.exe 存在于项目根目录 / Ensure ffmpeg.exe exists in project root**

   项目已包含 ffmpeg（已 UPX 压缩），如需更新可从 [gyan.dev](https://www.gyan.dev/ffmpeg/builds/) 下载。

3. **运行打包脚本 / Run Build Script**

```powershell
.\build.ps1
```

脚本会自动：
- 创建虚拟环境 `venv/`
- 安装依赖（Pillow、PyInstaller）
- 打包程序

4. **运行 / Run**

打包完成后，可执行文件位于 `dist\PPTXCompressor.exe`

## 使用说明 / Usage

1. 运行程序
2. 点击"Browse"按钮，选择要压缩的 PPTX 文件
3. 根据需要调整参数（可选）
4. 点击"Start Compression"
5. 等待压缩完成，压缩后的文件保存在原文件同目录

### 参数说明 / Parameter Description

#### 视频参数 / Video Parameters

| 参数 | 默认值 | 范围 | 说明 |
|------|--------|------|------|
| CRF | 28 | 18-51 | 恒定质量因子。值越大压缩率越高，质量越低。推荐 23-32 |
| Preset | slow | ultrafast ~ veryslow | 编码速度。越慢压缩率越高，推荐 slow 或 medium |
| Audio Bitrate | 128k | - | AAC 音频比特率 |

#### 图片参数 / Image Parameters

| 参数 | 默认值 | 说明 |
|------|--------|------|
| Size Threshold | 500 KB | 只有超过此大小的图片才会被压缩 |
| Max Width | 1920 px | 宽度超过此值的图片会按比例缩小 |

## 压缩效果 / Compression Results

根据测试，典型 PPTX 文件可达到以下压缩效果：

| 内容类型 | 原始大小 | 压缩后 | 压缩率 |
|----------|----------|--------|--------|
| 高清视频 | 100 MB | 20-30 MB | 70-80% |
| PNG 截图 | 2 MB | 300-500 KB | 75-85% |
| JPEG 照片 | 500 KB | 150-250 KB | 50-70% |

## 技术细节 / Technical Details

### 视频压缩 / Video Compression

使用 ffmpeg 进行 H.264 编码：

```bash
ffmpeg -i input.mp4 \
    -c:v libx264 \
    -crf 28 \
    -preset slow \
    -profile:v high \
    -level 4.1 \
    -pix_fmt yuv420p \
    -c:a aac \
    -b:a 128k \
    -movflags +faststart \
    output.mp4
```

### 图片压缩 / Image Compression

#### PNG 压缩

使用 PIL 进行无损压缩，保留透明度。

#### JPEG 压缩

```
1. 移除 EXIF 等元数据
2. 尝试不同质量级别（95/90/85/80）
3. 选择最优质量（文件最小且质量可接受）
4. 使用渐进式编码
```

## 常见问题 / FAQ

**Q: 压缩后视频质量下降明显怎么办？**

A: 降低 CRF 值（如改为 23 或更低），或使用更慢的 preset（如 slower）。

**Q: 压缩速度太慢怎么办？**

A: 将 preset 改为 medium 或 fast，牺牲少量压缩率换取更快的速度。

**Q: 某些视频压缩后反而变大了？**

A: 程序会自动检测，如果压缩后更大则保留原视频。

## 开发 / Development

### 项目结构 / Project Structure

```
ppt-compress/
├── compress_pptx.py              # Main program
├── build.ps1                     # Build script (auto-creates venv)
├── requirements.txt              # Python dependencies
├── ffmpeg.exe                    # ffmpeg (required)
├── README.md                     # This file
└── .gitignore                    # Git ignore configuration
```

### 运行源码 / Run from Source

```bash
# Create virtual environment (optional)
python -m venv venv
venv\Scripts\python -m pip install -r requirements.txt

# Run program
venv\Scripts\python compress_pptx.py
```

### 打包 / Build

```powershell
.\build.ps1
```

打包产物位于 `dist\PPTXCompressor.exe`（约 40 MB）。

## 许可证 / License

MIT License

## 致谢 / Acknowledgments

- [ffmpeg](https://ffmpeg.org/) - 视频处理
- [Pillow](https://python-pillow.org/) - Python 图像处理库
