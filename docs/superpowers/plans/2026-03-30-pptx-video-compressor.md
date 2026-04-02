# PPTX 压缩工具 实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 创建一个独立的 Windows exe 程序，用于压缩 PPTX 文件中的视频和图片。

**Architecture:** 使用 tkinter 构建 GUI，通过 zipfile 解压/打包 PPTX，调用内嵌的 ffmpeg 压缩视频，使用 Pillow/pngquant/oxipng 压缩图片，使用 PyInstaller 打包为单文件 exe。

**Tech Stack:** Python 3.13, tkinter, zipfile, subprocess, ffmpeg, Pillow, pngquant, oxipng, PyInstaller

---

## 文件结构

```
ppt-compress/
├── compress_pptx.py      # 主程序（GUI + 核心逻辑）
├── build.bat             # 打包脚本
├── ffmpeg.exe            # ffmpeg（必需）
├── pngquant.exe          # PNG 有损压缩（可选）
├── oxipng.exe            # PNG 无损优化（可选）
└── dist/
    └── PPTX压缩工具.exe
```

---

### Task 1: 创建主程序框架

**Files:**
- Create: `compress_pptx.py`

- [x] **Step 1: 创建基础 GUI 框架**

包含文件选择、高级选项、图片压缩选项、进度显示、结果区域。

- [x] **Step 2: 实现视频压缩功能**

使用 ffmpeg 进行 H.264 编码，支持 CRF、preset、音频码率参数调整。

- [x] **Step 3: 实现图片压缩功能**

- PNG: 三级压缩策略（oxipng 无损 → pngquant 有损 → PIL 回退）
- JPEG: 质量优化，移除元数据
- WebP: 无损优化
- GIF: 保持原样
- 支持大图缩放

---

### Task 2: 创建打包脚本

**Files:**
- Create: `build.bat`

- [x] **Step 1: 创建打包脚本**

检查依赖（ffmpeg.exe, pngquant.exe, oxipng.exe），安装 PyInstaller 和 Pillow，执行打包。

- [x] **Step 2: 测试打包**

```bash
# 运行打包脚本
build.bat
```

预期：生成 `dist\PPTX压缩工具.exe`

---

### Task 3: 完善与优化

- [x] **Step 1: 添加图片压缩进度显示**

图片压缩占进度 0-50%，视频压缩占进度 50-100%。

- [x] **Step 2: 添加详细的结果统计**

显示视频和图片的原始大小、压缩后大小、压缩率，以及图片压缩方式统计。

- [x] **Step 3: 测试完整流程**

1. 运行 `dist\PPTX压缩工具.exe`
2. 选择一个包含视频和图片的 PPTX 文件
3. 点击"开始压缩"
4. 验证输出文件是否正确

---

## 使用说明

### 准备工作

1. 下载 ffmpeg.exe（必需）：https://www.gyan.dev/ffmpeg/builds/
2. 下载 pngquant.exe（可选）：https://pngquant.org/
3. 下载 oxipng.exe（可选）：https://github.com/shssoichiro/oxipng/releases

### 运行程序

1. 运行 `PPTX压缩工具.exe`
2. 选择 PPTX 文件
3. 可调整视频参数（CRF、Preset、音频码率）
4. 可调整图片参数（大小阈值、最大宽度）
5. 点击"开始压缩"
6. 压缩后的文件保存在原文件同目录，文件名添加 `_compressed` 后缀

### 参数说明

#### 视频参数

| 参数 | 默认值 | 推荐范围 | 说明 |
|------|--------|----------|------|
| CRF | 28 | 18-51 | 越大压缩率越高，质量越低。推荐 23-32 |
| Preset | slow | - | 越慢压缩率越高，推荐 slow 或 medium |
| 音频码率 | 128k | 64k-256k | AAC 音频码率 |

#### 图片参数

| 参数 | 默认值 | 说明 |
|------|--------|------|
| 大小阈值 | 500 KB | 超过此大小的图片将被压缩 |
| 最大宽度 | 1920 px | 超大图片将按比例缩小 |
