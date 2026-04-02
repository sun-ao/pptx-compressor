# PPTX 压缩工具设计文档

## 概述

一个独立的 Windows exe 程序，用于压缩 PPTX 文件中的视频和图片，无需安装 Python 或任何依赖。

## 目标用户

- 需要压缩 PPT 文件大小的办公用户
- 低配电脑用户（需保证兼容性）

## 核心功能

1. 选择 PPTX 文件
2. 自动检测并压缩其中的视频和图片
3. 输出压缩后的 PPTX 文件
4. 支持高级参数调整

## 技术方案

### 依赖

- Python 3.13.5
- tkinter（内置 GUI）
- Pillow（图片处理）
- ffmpeg（视频压缩，嵌入 exe）
- pngquant（PNG 有损压缩，可选，嵌入 exe）
- oxipng（PNG 无损优化，可选，嵌入 exe）

### 视频压缩

使用 ffmpeg 进行 H.264 编码：

```bash
ffmpeg -i input.mp4 -c:v libx264 -crf 28 -preset slow -profile:v high -level 4.1 -pix_fmt yuv420p -c:a aac -b:a 128k -movflags +faststart output.mp4
```

#### 可调参数

| 参数 | 默认值 | 范围 | 说明 |
|------|--------|------|------|
| CRF | 28 | 18-51 | 越大压缩率越高，质量越低 |
| Preset | slow | ultrafast~veryslow | 编码速度与压缩率权衡 |
| 音频码率 | 128k | - | AAC 音频码率 |

#### 支持的视频格式

`.mp4`, `.avi`, `.mov`, `.wmv`, `.mkv`, `.flv`, `.webm`

### 图片压缩

#### 支持的图片格式

`.png`, `.jpg`, `.jpeg`, `.bmp`, `.tiff`, `.tif`, `.gif`, `.webp`

#### 可调参数

| 参数 | 默认值 | 说明 |
|------|--------|------|
| 大小阈值 | 500 KB | 超过此大小的图片将被压缩 |
| 最大宽度 | 1920 px | 超大图片将按比例缩小 |

#### PNG 压缩策略（三级压缩）

1. **oxipng 无损优化**：适合图标/色块类 PNG，使用 `-o 6` 最高优化级别
2. **pngquant 有损压缩**：适合截图/照片类 PNG，质量 85-95（视觉无损）
3. **PIL 回退**：当上述工具不可用时，使用 PIL 进行无损压缩

#### JPEG 压缩策略

- 移除元数据
- 自动寻找最优质量级别（95/90/85/80）
- 使用渐进式编码和优化选项

#### WebP 压缩策略

- 使用无损压缩模式
- 优化方法级别 6

#### GIF 处理

- 保持原样（动图复杂，不做处理）

### 处理流程

```
1. 用户选择 PPTX 文件
2. 解压到临时目录（zipfile）
3. 扫描 ppt/media/ 目录检测视频和图片文件
4. 先压缩图片（占进度 0-50%）
   - 检查文件大小是否超过阈值
   - 检查是否需要缩放
   - 根据格式选择压缩策略
5. 再压缩视频（占进度 50-100%）
   - 调用 ffmpeg 进行 H.264 编码
   - 仅当压缩后更小时替换原文件
6. 重新打包为 PPTX
7. 清理临时文件
```

### 界面设计

- 单窗口布局（550x600）
- 文件选择区域
- 高级选项面板（视频参数）
- 图片压缩选项面板
- 进度条显示
- 结果文本区域（显示详细统计）

## 兼容性考虑

1. **内存优化**：逐个处理媒体文件，避免同时解压多个文件
2. **CPU 优化**：默认使用 slow preset，低配电脑可调为 medium/fast
3. **磁盘优化**：及时清理临时文件
4. **错误处理**：完善的异常捕获和用户提示
5. **编码兼容**：使用 yuv420p 像素格式，确保最大兼容性
6. **工具回退**：pngquant/oxipng 不可用时自动使用 PIL 回退

## 打包方案

使用 PyInstaller：

```bash
pyinstaller --onefile --windowed \
    --add-data "ffmpeg.exe;." \
    --add-data "pngquant.exe;." \
    --add-data "oxipng.exe;." \
    --hidden-import=PIL \
    --hidden-import=PIL.Image \
    --name "PPTX压缩工具" \
    --clean compress_pptx.py
```

## 文件结构

```
ppt-compress/
├── compress_pptx.py      # 主程序
├── build.bat             # 打包脚本
├── ffmpeg.exe            # ffmpeg（必需）
├── pngquant.exe          # PNG 有损压缩（可选）
├── oxipng.exe            # PNG 无损优化（可选）
└── dist/
    └── PPTX压缩工具.exe
```

## 压缩效果示例

| 媒体类型 | 原始大小 | 压缩后 | 压缩率 |
|----------|----------|--------|--------|
| 视频 (MP4) | 100 MB | 25 MB | 75% |
| PNG 截图 | 2 MB | 400 KB | 80% |
| JPEG 照片 | 500 KB | 200 KB | 60% |
