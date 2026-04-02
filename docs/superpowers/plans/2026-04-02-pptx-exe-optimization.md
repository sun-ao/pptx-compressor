# PPTX 压缩工具 - 打包优化实施计划

> **For agentic workers:** 使用 superpowers:executing-plans 执行此计划

**目标:** 将打包后的 exe 从 289MB 优化到 50-80MB，实现独立运行

**方案:** 使用精简版 ffmpeg + PyInstaller 优化配置

**技术栈:** PyInstaller, ffmpeg-lite, UPX

---

## 文件修改清单

| 文件 | 操作 |
|------|------|
| `ffmpeg.exe` | 替换为精简版 (~20MB) |
| `PPTX压缩工具.spec` | 优化打包配置 |
| `build.ps1` | 更新打包脚本 |

---

## 实施步骤

### 步骤 1: 下载精简版 ffmpeg

- [ ] **Step 1: 下载 ffmpeg-lite**

从 gyan.dev 下载精简版 ffmpeg (只包含 H.264 编码):
```
URL: https://www.gyan.dev/ffmpeg/builds/ffmpeg-release-essentials.zip
提取: ffmpeg.exe (位于 bin 目录)
大小: 约 20-30MB (完整版约 80MB)
```

- [ ] **Step 2: 替换现有 ffmpeg.exe**

删除旧 ffmpeg.exe，替换为精简版

---

### 步骤 2: 优化 PyInstaller 配置

- [ ] **Step 1: 更新 spec 文件**

修改 `PPTX压缩工具.spec`:

```python
# 排除不需要的 Python 模块
excludes=[
    'tkinter.tix',
    'tkinter.ttk',
    'unittest',
    'test',
    'pytest',
    'distutils',
    'setuptools',
    'pip',
    'wheel',
    'xml.dom',
    'xml.etree',
    'xml.sax',
    'html.parser',
    'http.cookies',
    'http.cookiejar',
    'urllib.request',
    'urllib.parse',
    'urllib.error',
    'urllib3',
    'requests',
    'chardet',
    'idna',
    'ssl',
    'cryptography',
    'pyasn1',
    'cffi',
    'pycparser',
    'jinja2',
    'markupsafe',
    'certifi',
]

# 启用 UPX 压缩
upx=True,

# 优化级别
optimize=2,
```

- [ ] **Step 2: 更新打包脚本 build.ps1**

添加排除模块参数:
```powershell
python -m PyInstaller --onefile --windowed `
    --add-data 'ffmpeg.exe;.' `
    --add-data 'pngquant.exe;.' `
    --add-data 'oxipng.exe;.' `
    --hidden-import=PIL --hidden-import=PIL.Image `
    --exclude-module=tkinter.tix `
    --exclude-module=unittest `
    --exclude-module=test `
    --exclude-module=pytest `
    --exclude-module=distutils `
    --exclude-module=setuptools `
    --exclude-module=pip `
    --exclude-module=wheel `
    --exclude-module=xml.dom `
    --exclude-module=xml.etree `
    --exclude-module=xml.sax `
    --exclude-module=html.parser `
    --exclude-module=http.cookies `
    --exclude-module=http.cookiejar `
    --exclude-module=urllib.request `
    --exclude-module=urllib.parse `
    --exclude-module=urllib.error `
    --exclude-module=urllib3 `
    --exclude-module=requests `
    --exclude-module=chardet `
    --exclude-module=idna `
    --exclude-module=ssl `
    --exclude-module=cryptography `
    --name 'PPTX压缩工具' `
    --clean compress_pptx.py
```

---

### 步骤 3: 重新打包并验证

- [ ] **Step 1: 清理旧构建**

```powershell
Remove-Item -Recurse -Force build\dist -ErrorAction SilentlyContinue
```

- [ ] **Step 2: 执行打包**

```powershell
.\build.ps1
```

- [ ] **Step 3: 验证结果**

检查 exe 大小:
```powershell
Get-Item dist\PPTX压缩工具.exe | Select-Object Name, @{N='Size(MB)';E={[math]::Round($_.Length/1MB,2)}}
```

预期结果: 50-80MB

- [ ] **Step 4: 功能测试**

运行 exe 确保功能正常

---

## 预期效果

| 优化项 | 优化前 | 优化后 | 节省 |
|--------|--------|--------|------|
| ffmpeg | ~80MB | ~25MB | ~55MB |
| Python 标准库 | ~30MB | ~15MB | ~15MB |
| UPX 压缩 | 无 | 启用 | ~15MB |
| **总计** | **289MB** | **~60MB** | **~229MB** |

---

## 注意事项

1. 精简版 ffmpeg 只支持 H.264 视频编码，不影响核心功能
2. 排除的模块都是不需要的 GUI/网络/测试相关
3. 如果打包后 exe 启动失败，逐步移除排除项
