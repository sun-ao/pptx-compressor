# PPTX 压缩工具 - 构建优化实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 改造构建流程，让新用户 clone 项目后能轻松打出较小的 exe 包

**Architecture:** 使用虚拟环境隔离依赖，自动创建/使用 venv，确保打包时只包含必要的模块

**Tech Stack:** Python 虚拟环境 + PyInstaller + Pillow

---

### 文件结构

```
项目根目录/
├── .gitignore           # 修改：排除 venv/, dist/, build/
├── requirements.txt     # 新建：明确依赖列表
├── build.ps1            # 修改：添加虚拟环境自动管理
└── compress_pptx.py     # 无需修改
```

---

### Task 1: 创建 requirements.txt

**Files:**
- Create: `E:\anno\ppt-compress\requirements.txt`

- [ ] **Step 1: 创建 requirements.txt**

```
Pillow>=10.0.0
pyinstaller>=6.0.0
```

---

### Task 2: 更新 .gitignore

**Files:**
- Modify: `E:\anno\ppt-compress\.gitignore`

- [ ] **Step 1: 添加虚拟环境和构建产物排除规则**

在 .gitignore 末尾添加：
```
# 虚拟环境
venv/
.venv/
env/

# 构建产物
dist/
build/
*.spec
```

---

### Task 3: 重写 build.ps1

**Files:**
- Modify: `E:\anno\ppt-compress\build.ps1`

- [ ] **Step 1: 重写 build.ps1 添加虚拟环境支持**

新的逻辑：
1. 检查是否存在 venv，不存在则创建
2. 在虚拟环境中安装依赖
3. 使用虚拟环境的 Python 进行打包
4. 保留原有的排除模块列表

```powershell
# PPTX Compressor - Build Script
# 自动创建/使用虚拟环境

Write-Host "====================================" -ForegroundColor Cyan
Write-Host "PPTX Compressor - Build Script" -ForegroundColor Cyan
Write-Host "====================================" -ForegroundColor Cyan
Write-Host ""

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# 检查 ffmpeg.exe
if (-not (Test-Path "ffmpeg.exe")) {
    Write-Host "[ERROR] ffmpeg.exe not found" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# 虚拟环境路径
$venv_path = "venv"
$python_exe = if ($IsWindows -or $env:OS -eq "Windows_NT") {
    "$venv_path\Scripts\python.exe"
} else {
    "$venv_path/bin/python"
}

# 检查/创建虚拟环境
if (-not (Test-Path $python_exe)) {
    Write-Host "[INFO] Creating virtual environment..." -ForegroundColor Yellow
    python -m venv $venv_path
}

# 安装依赖
Write-Host "[INFO] Installing dependencies..." -ForegroundColor Yellow
& $python_exe -m pip install --upgrade pip
& $python_exe -m pip install -r requirements.txt

# 检查 UPX
$upx_available = $false
try {
    $null = Get-Command upx -ErrorAction Stop
    $upx_available = $true
    Write-Host "[INFO] UPX available, will compress" -ForegroundColor Green
} catch {
    Write-Host "[WARNING] UPX not available" -ForegroundColor Yellow
}

Write-Host "[BUILD] Building..." -ForegroundColor White
Write-Host ""

# Clean old build
if (Test-Path "build") { Remove-Item -Recurse -Force build }
if (Test-Path "dist") { Remove-Item -Recurse -Force dist }

# Build command - 使用虚拟环境中的 Python
& $python_exe -m PyInstaller --onefile --windowed `
    --add-data "ffmpeg.exe;." `
    --hidden-import=PIL --hidden-import=PIL.Image `
    --exclude-module=tkinter.tix `
    --exclude-module=unittest `
    --exclude-module=test `
    --exclude-module=pytest `
    --exclude-module=distutils `
    --exclude-module=setuptools `
    --exclude-module=pip `
    --exclude-module=wheel `
    --exclude-module=pkg_resources `
    --exclude-module=urllib3 `
    --exclude-module=requests `
    --exclude-module=chardet `
    --exclude-module=idna `
    --exclude-module=certifi `
    --exclude-module=cryptography `
    --exclude-module=jinja2 `
    --exclude-module=markupsafe `
    --exclude-module=numpy `
    --exclude-module=pandas `
    --exclude-module=matplotlib `
    --exclude-module=scipy `
    --exclude-module=IPython `
    --exclude-module=jupyter `
    --exclude-module=notebook `
    --exclude-module=xmlrpc `
    --exclude-module=sqlite3 `
    --exclude-module=email `
    --exclude-module=html `
    --exclude-module=http `
    --exclude-module=xml `
    --exclude-module=csv `
    --exclude-module=json `
    --exclude-module=base64 `
    --exclude-module=hashlib `
    --exclude-module=ssl `
    --exclude-module=asyncio `
    --exclude-module=multiprocessing `
    --exclude-module=concurrent `
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
```

---

### Task 4: 验证构建

**Files:**
- Test: `E:\anno\ppt-compress\dist\PPTXCompressor.exe`

- [ ] **Step 1: 运行 build.ps1 进行测试**

```powershell
./build.ps1
```

- [ ] **Step 2: 检查输出文件大小**

预期：约 36MB（与之前优化后的结果一致）

- [ ] **Step 3: 验证程序能正常运行**

运行 exe，确认 GUI 能正常启动

---

### Task 5: 清理旧的 spec 文件

**Files:**
- Delete: `E:\anno\ppt-compress\PPTXCompressor.spec`

- [ ] **Step 1: 删除不再需要的 spec 文件**

由于改用命令行参数构建，spec 文件不再需要

---

## 执行方式

**Plan complete and saved to `docs/superpowers/plans/2026-04-02-pptx-build-optimization.md`. Two execution options:**

1. **Subagent-Driven (recommended)** - 调度子任务执行
2. **Inline Execution** - 在当前会话中执行

选择哪种方式？
