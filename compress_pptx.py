#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
PPTX Compressor - PPTX 压缩工具
用于压缩 PowerPoint 文件（PPTX）中的视频和图片以减小文件大小
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import zipfile
import os
import subprocess
import shutil
import tempfile
import threading
import time
import queue
import sys
from pathlib import Path
from typing import Optional, Tuple, List

try:
    from PIL import Image
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False


class PPTXCompressor:
    """PPTX 压缩工具主类 - 处理 GUI 和压缩流程"""

    # 支持的视频格式
    VIDEO_EXTENSIONS: set = {'.mp4', '.avi', '.mov', '.wmv', '.mkv', '.flv', '.webm'}

    # 支持的图片格式
    IMAGE_EXTENSIONS: set = {'.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.tif', '.gif', '.webp'}

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("文件压缩")
        self.root.geometry("550x600")
        self.root.resizable(False, False)

        # 变量
        self.input_file: tk.StringVar = tk.StringVar()
        self.crf_value: tk.StringVar = tk.StringVar(value="28")
        self.preset_value: tk.StringVar = tk.StringVar(value="slow")
        self.audio_bitrate: tk.StringVar = tk.StringVar(value="128k")
        self.image_threshold: tk.StringVar = tk.StringVar(value="500")  # KB
        self.max_image_width: tk.StringVar = tk.StringVar(value="1920")
        self.compressing: bool = False
        self.cancelled: bool = False
        self.ffmpeg_process: Optional[subprocess.Popen] = None

        self._create_widgets()

    def _create_widgets(self) -> None:
        """创建界面组件"""
        # 文件选择区域
        file_frame = ttk.LabelFrame(self.root, text="选择文件", padding=10)
        file_frame.pack(fill="x", padx=10, pady=(10, 0))

        ttk.Entry(file_frame, textvariable=self.input_file, width=50).pack(side="left", padx=(0, 10))
        ttk.Button(file_frame, text="浏览", command=self._select_file).pack(side="left")

        ttk.Label(self.root, text="支持 PPT 文件、视频文件（MP4/AVI/MOV等）、图片文件（PNG/JPG等）", foreground="gray").pack(anchor="w", padx=15, pady=(2, 5))

        # 高级选项区域
        self.options_frame = ttk.LabelFrame(self.root, text="压缩设置", padding=10)
        self.options_frame.pack(fill="x", padx=10, pady=5)

        # CRF 值
        crf_frame = ttk.Frame(self.options_frame)
        crf_frame.pack(fill="x", pady=2)
        ttk.Label(crf_frame, text="视频质量:", width=15).pack(side="left")
        ttk.Entry(crf_frame, textvariable=self.crf_value, width=10).pack(side="left")
        ttk.Label(crf_frame, text="数值越大 压缩越多 画质越低").pack(side="left", padx=10)

        # Preset
        preset_frame = ttk.Frame(self.options_frame)
        preset_frame.pack(fill="x", pady=2)
        ttk.Label(preset_frame, text="压缩速度:", width=15).pack(side="left")
        preset_combo = ttk.Combobox(preset_frame, textvariable=self.preset_value, width=15, state="readonly")
        preset_combo['values'] = ('ultrafast', 'superfast', 'veryfast', 'faster', 'fast', 'medium', 'slow', 'slower', 'veryslow')
        preset_combo.pack(side="left")
        ttk.Label(preset_frame, text="越慢 压缩效果越好").pack(side="left", padx=10)

        # 音频码率
        audio_frame = ttk.Frame(self.options_frame)
        audio_frame.pack(fill="x", pady=2)
        ttk.Label(audio_frame, text="音频质量:", width=15).pack(side="left")
        ttk.Entry(audio_frame, textvariable=self.audio_bitrate, width=10).pack(side="left")

        # 图片压缩选项
        image_frame = ttk.LabelFrame(self.root, text="图片压缩", padding=10)
        image_frame.pack(fill="x", padx=10, pady=5)

        # 图片大小阈值
        threshold_frame = ttk.Frame(image_frame)
        threshold_frame.pack(fill="x", pady=2)
        ttk.Label(threshold_frame, text="压缩起点 (KB):", width=18).pack(side="left")
        ttk.Entry(threshold_frame, textvariable=self.image_threshold, width=10).pack(side="left")
        ttk.Label(threshold_frame, text="只有大于这个值的图片才会被压缩").pack(side="left", padx=10)

        # 最大宽度
        maxwidth_frame = ttk.Frame(image_frame)
        maxwidth_frame.pack(fill="x", pady=2)
        ttk.Label(maxwidth_frame, text="最大宽度 (像素):", width=18).pack(side="left")
        ttk.Entry(maxwidth_frame, textvariable=self.max_image_width, width=10).pack(side="left")
        ttk.Label(maxwidth_frame, text="超过这个宽度的图片会自动缩小").pack(side="left", padx=10)

        # 按钮区域
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(pady=15)

        self.compress_btn = ttk.Button(btn_frame, text="开始压缩", command=self._start_compress)
        self.compress_btn.pack(side="left", padx=5)

        self.cancel_btn = ttk.Button(btn_frame, text="取消", command=self._cancel_compress)
        self.cancel_btn.pack(side="left", padx=5)
        self.cancel_btn.pack_forget()

        # 进度区域
        progress_frame = ttk.LabelFrame(self.root, text="进度", padding=10)
        progress_frame.pack(fill="x", padx=10, pady=5)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill="x")

        self.status_label = ttk.Label(progress_frame, text="")
        self.status_label.pack(pady=5)

        # 结果区域
        result_frame = ttk.LabelFrame(self.root, text="压缩结果", padding=10)
        result_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.result_text = tk.Text(result_frame, height=8, width=60, state="disabled")
        self.result_text.pack(fill="both", expand=True)

    @staticmethod
    def _get_file_type(file_path: str) -> Optional[str]:
        """根据扩展名判断文件类型，返回 'video'/'image'/'pptx'/None"""
        ext = Path(file_path).suffix.lower()
        if ext == '.pptx':
            return 'pptx'
        if ext in {'.mp4', '.avi', '.mov', '.wmv', '.mkv', '.flv', '.webm'}:
            return 'video'
        if ext in {'.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.tif', '.gif', '.webp'}:
            return 'image'
        return None

    def _select_file(self) -> None:
        """选择文件"""
        file_path = filedialog.askopenfilename(
            title="选择要压缩的文件",
            filetypes=[
                ("支持的格式", "*.pptx *.mp4 *.avi *.mov *.wmv *.mkv *.flv *.webm *.png *.jpg *.jpeg *.bmp *.tiff *.tif *.gif *.webp"),
                ("所有文件", "*.*"),
            ]
        )
        if file_path:
            self.input_file.set(file_path)

    def _start_compress(self) -> None:
        """开始压缩"""
        if self.compressing:
            return

        input_path = self.input_file.get()
        if not input_path:
            messagebox.showwarning("提示", "请先选择一个文件")
            return

        if not os.path.exists(input_path):
            messagebox.showerror("错误", "文件不存在")
            return

        self.compressing = True
        self.cancelled = False
        self.compress_btn.pack_forget()
        self.cancel_btn.pack(side="left", padx=5)
        self._clear_result()
        self.progress_var.set(0)

        # 在新线程中执行压缩
        thread = threading.Thread(target=self._compress, args=(input_path,))
        thread.daemon = True
        thread.start()

    def _cancel_compress(self) -> None:
        """取消压缩"""
        self.cancelled = True
        if self.ffmpeg_process:
            try:
                self.ffmpeg_process.terminate()
            except OSError:
                pass

    def _get_ffmpeg_path(self) -> str:
        """获取 ffmpeg 路径"""
        if getattr(sys, 'frozen', False):
            # 打包后的 exe
            base_path = sys._MEIPASS
            return os.path.join(base_path, 'ffmpeg.exe')
        else:
            # 开发环境
            return os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ffmpeg.exe')

    def _get_tool_path(self, tool_name: str) -> str:
        """获取工具路径 (pngquant, oxipng 等)"""
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_path, tool_name)

    def _compress_png_advanced(self, image_path: str, original_size: int) -> Tuple[int, bool, Optional[str]]:
        """
        高级 PNG 压缩：使用 PIL 进行无损压缩
        返回: (新文件大小, 是否被压缩, 压缩方式)
        """
        if not PIL_AVAILABLE:
            return original_size, False, None

        try:
            with Image.open(image_path) as img:
                mode = img.mode

                # 处理不同模式
                if mode == 'P':
                    if 'transparency' in img.info:
                        img.save(image_path, 'PNG', optimize=True, compress_level=9)
                    else:
                        img.convert('RGB').save(image_path, 'PNG', optimize=True, compress_level=9)
                elif mode in ('RGBA', 'LA'):
                    img.save(image_path, 'PNG', optimize=True, compress_level=9)
                else:
                    img.convert('RGB').save(image_path, 'PNG', optimize=True, compress_level=9)

                new_size = os.path.getsize(image_path)
                if new_size < original_size:
                    return new_size, True, "PIL Lossless"

            return original_size, False, None

        except (OSError, IOError, ValueError) as e:
            return original_size, False, None

    def _compress_image(self, image_path: str, threshold_kb: int, max_width: int) -> Tuple[int, bool, Optional[str]]:
        """
        压缩单个图片
        返回: (新文件大小, 是否被压缩, 压缩方式)
        """
        if not PIL_AVAILABLE:
            return os.path.getsize(image_path), False, None

        original_size = os.path.getsize(image_path)
        threshold_bytes = threshold_kb * 1024

        # 如果文件小于阈值，不压缩
        if original_size < threshold_bytes:
            return original_size, False, None

        try:
            with Image.open(image_path) as img:
                original_width, original_height = img.size
                mode = img.mode

                # 检查是否需要缩放
                need_resize = original_width > max_width
                new_width = max_width if need_resize else original_width
                new_height = int(original_height * (new_width / original_width)) if need_resize else original_height

                # 如果需要缩放
                if need_resize:
                    # 保持高质量缩放
                    img = img.copy()
                    img.thumbnail((new_width, new_height), Image.Resampling.LANCZOS)

                # 根据格式选择压缩方式
                ext = Path(image_path).suffix.lower()
                temp_path = image_path + ".compressed"

                if ext in {'.png', '.bmp', '.tiff', '.tif'}:
                    # PNG/BMP/TIFF -> 使用高级PNG压缩
                    if mode == 'P':
                        if 'transparency' in img.info:
                            img = img.convert('RGBA')
                            save_mode = 'RGBA'
                        else:
                            img = img.convert('RGB')
                            save_mode = 'RGB'
                    elif mode in ('RGBA', 'LA'):
                        save_mode = 'RGBA'
                    else:
                        save_mode = 'RGB'

                    # 如果需要缩放，先保存临时文件
                    if need_resize:
                        temp_png = image_path + ".resized.png"
                        if save_mode == 'RGBA':
                            img.save(temp_png, 'PNG')
                        else:
                            img.convert('RGB').save(temp_png, 'PNG')

                        # 替换原文件
                        os.remove(image_path)
                        os.rename(temp_png, image_path)

                    # 使用高级PNG压缩
                    new_size, was_compressed, compress_method = self._compress_png_advanced(image_path, original_size)
                    return new_size, was_compressed, compress_method

                elif ext in {'.jpg', '.jpeg'}:
                    # JPEG 优化 - 移除元数据，保持质量
                    if mode in ('RGBA', 'LA', 'P'):
                        img = img.convert('RGB')

                    # 尝试不同质量级别，找到最优
                    best_size = original_size
                    best_quality = None

                    for quality in [95, 90, 85, 80]:
                        img.save(temp_path, 'JPEG', quality=quality, optimize=True, progressive=True)
                        test_size = os.path.getsize(temp_path)
                        if test_size < best_size:
                            best_size = test_size
                            best_quality = quality

                    if best_quality is not None:
                        # 使用最优质量重新保存
                        img.save(temp_path, 'JPEG', quality=best_quality, optimize=True, progressive=True)
                        new_size = os.path.getsize(temp_path)
                        if new_size < original_size:
                            os.remove(image_path)
                            os.rename(temp_path, image_path)
                            return new_size, True, "JPEG Optimized"

                    if os.path.exists(temp_path):
                        os.remove(temp_path)
                    return original_size, False, None

                elif ext == '.gif':
                    # GIF 保持原样（动图复杂，不做处理）
                    return original_size, False, None

                elif ext == '.webp':
                    # WebP 优化
                    if mode in ('RGBA', 'LA', 'P'):
                        img.save(temp_path, 'WEBP', lossless=True, quality=100, method=6)
                    else:
                        img = img.convert('RGB')
                        img.save(temp_path, 'WEBP', lossless=True, quality=100, method=6)

                    new_size = os.path.getsize(temp_path)
                    if new_size < original_size:
                        os.remove(image_path)
                        os.rename(temp_path, image_path)
                        return new_size, True, "WebP Optimized"
                    else:
                        os.remove(temp_path)
                        return original_size, False, None

                return original_size, False, None

        except (OSError, IOError, ValueError) as e:
            return original_size, False, None

    def _validate_params(self) -> Optional[Tuple[int, int, int]]:
        """验证并返回参数 (crf, image_threshold, max_image_width)，失败返回 None"""
        try:
            crf = int(self.crf_value.get())
            if crf < 18 or crf > 51:
                raise ValueError
        except ValueError:
            self._show_error("视频质量必须是 18 到 51 之间的整数")
            return None

        try:
            image_threshold = int(self.image_threshold.get())
            if image_threshold < 0:
                raise ValueError
        except ValueError:
            self._show_error("压缩起点必须是正整数")
            return None

        try:
            max_image_width = int(self.max_image_width.get())
            if max_image_width < 100:
                raise ValueError
        except ValueError:
            self._show_error("最大宽度必须是大于 100 的整数")
            return None

        return crf, image_threshold, max_image_width

    def _compress(self, input_path: str) -> None:
        """执行压缩 - 自动识别文件类型"""
        try:
            ffmpeg_path = self._get_ffmpeg_path()
            if not os.path.exists(ffmpeg_path):
                self._show_error("找不到 ffmpeg.exe，请确保它和程序在同一目录下")
                return

            params = self._validate_params()
            if not params:
                return
            crf, image_threshold, max_image_width = params

            file_type = self._get_file_type(input_path)
            if not file_type:
                self._show_error("不支持这种文件格式")
                return

            if file_type in ('video', 'image'):
                self._compress_single_file(input_path, file_type, ffmpeg_path, crf, image_threshold, max_image_width)
            else:
                self._compress_pptx(input_path, ffmpeg_path, crf, image_threshold, max_image_width)

        except Exception as e:
            self._show_error(f"压缩失败: {str(e)}")
        finally:
            self.compressing = False
            self.ffmpeg_process = None
            self.root.after(0, self._restore_compress_btn)

    def _compress_single_file(self, input_path: str, file_type: str, ffmpeg_path: str, crf: int, image_threshold: int, max_image_width: int) -> None:
        """压缩单个文件（视频或图片）"""
        input_dir = os.path.dirname(input_path)
        input_name = os.path.basename(input_path)
        name_without_ext = os.path.splitext(input_name)[0]
        ext = Path(input_path).suffix.lower()

        if file_type == 'video':
            self._update_status(f"正在压缩: {input_name}")
            result = self._compress_single_video(
                input_path, ffmpeg_path, crf,
                lambda pct: self.root.after(0, lambda p=pct: self.progress_var.set(p * 100))
            )
            if result is None:
                return

            orig_size, comp_size, output_path = result
            if comp_size < orig_size:
                final_name = f"{name_without_ext}_compressed.mp4"
                final_path = os.path.join(input_dir, final_name)
                counter = 1
                while os.path.exists(final_path):
                    final_path = os.path.join(input_dir, f"{name_without_ext}_compressed_{counter}.mp4")
                    counter += 1
                os.rename(output_path, final_path)
                ratio = (1 - comp_size / orig_size) * 100
                result_text = (
                    f"压缩完成！\n\n"
                    f"原始大小: {orig_size / 1024 / 1024:.2f} MB\n"
                    f"压缩后: {comp_size / 1024 / 1024:.2f} MB\n"
                    f"压缩率: {ratio:.1f}%\n\n"
                    f"输出文件: {os.path.basename(final_path)}"
                )
            else:
                if os.path.exists(output_path):
                    os.remove(output_path)
                final_name = f"{name_without_ext}_compressed{ext}"
                final_path = os.path.join(input_dir, final_name)
                counter = 1
                while os.path.exists(final_path):
                    final_path = os.path.join(input_dir, f"{name_without_ext}_compressed_{counter}{ext}")
                    counter += 1
                shutil.copy2(input_path, final_path)
                result_text = (
                    f"压缩完成！\n\n"
                    f"原始大小: {orig_size / 1024 / 1024:.2f} MB\n"
                    f"压缩后: {orig_size / 1024 / 1024:.2f} MB\n"
                    f"压缩率: 0.0%（文件已经是最优状态）\n\n"
                    f"输出文件: {os.path.basename(final_path)}"
                )

            self.root.after(0, lambda: self._update_status("压缩完成！"))
            self.root.after(0, lambda: self._show_result(result_text))
            self.root.after(0, lambda: self.progress_var.set(100))

        elif file_type == 'image':
            temp_dir = tempfile.mkdtemp()
            try:
                temp_file = os.path.join(temp_dir, f"temp{ext}")
                shutil.copy2(input_path, temp_file)
                self._update_status(f"正在压缩: {input_name}")
                orig_size = os.path.getsize(input_file := temp_file)
                new_size, was_compressed, method = self._compress_image(
                    temp_file, 0, max_image_width
                )
                if self.cancelled:
                    return

                if was_compressed and new_size < orig_size:
                    final_name = f"{name_without_ext}_compressed{ext}"
                    final_path = os.path.join(input_dir, final_name)
                    counter = 1
                    while os.path.exists(final_path):
                        final_path = os.path.join(input_dir, f"{name_without_ext}_compressed_{counter}{ext}")
                        counter += 1
                    shutil.copy2(temp_file, final_path)
                    ratio = (1 - new_size / orig_size) * 100
                    result_text = (
                        f"压缩完成！\n\n"
                        f"原始大小: {orig_size / 1024 / 1024:.2f} MB\n"
                        f"压缩后: {new_size / 1024 / 1024:.2f} MB\n"
                        f"压缩率: {ratio:.1f}%\n"
                        f"压缩方式: {method or '未知'}\n\n"
                        f"输出文件: {os.path.basename(final_path)}"
                    )
                else:
                    final_name = f"{name_without_ext}_compressed{ext}"
                    final_path = os.path.join(input_dir, final_name)
                    counter = 1
                    while os.path.exists(final_path):
                        final_path = os.path.join(input_dir, f"{name_without_ext}_compressed_{counter}{ext}")
                        counter += 1
                    shutil.copy2(input_path, final_path)
                    result_text = (
                        f"Compression completed!\n\n"
                        f"Original: {orig_size / 1024 / 1024:.2f} MB\n"
                        f"Compressed: {orig_size / 1024 / 1024:.2f} MB\n"
                        f"Ratio: 0.0% (file already optimal)\n\n"
                        f"Output: {os.path.basename(final_path)}"
                    )

                self.root.after(0, lambda: self._update_status("压缩完成！"))
                self.root.after(0, lambda: self._show_result(result_text))
                self.root.after(0, lambda: self.progress_var.set(100))
            finally:
                try:
                    shutil.rmtree(temp_dir)
                except OSError:
                    pass

    def _compress_single_video(self, video_path: str, ffmpeg_path: str, crf: int, update_progress) -> Optional[Tuple[int, int, str]]:
        """压缩单个视频，返回 (原始大小, 压缩大小, 输出路径)，取消返回 None"""
        output_path = video_path + ".compressed.mp4"
        cmd = [
            ffmpeg_path, '-i', video_path,
            '-c:v', 'libx264', '-crf', str(crf),
            '-preset', self.preset_value.get(),
            '-profile:v', 'high', '-level', '4.1',
            '-pix_fmt', 'yuv420p',
            '-c:a', 'aac', '-b:a', self.audio_bitrate.get(),
            '-movflags', '+faststart', '-y', output_path,
        ]

        progress_queue: queue.Queue = queue.Queue()
        process = subprocess.Popen(
            cmd, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        self.ffmpeg_process = process

        def _reader(proc: subprocess.Popen, q: queue.Queue) -> None:
            try:
                buf = b''
                while True:
                    chunk = proc.stderr.read(1)
                    if not chunk:
                        break
                    if chunk in (b'\r', b'\n'):
                        line = buf.decode('utf-8', errors='replace').strip()
                        buf = b''
                        if 'Duration:' in line:
                            q.put(('duration', line))
                        elif 'time=' in line:
                            q.put(('time', line))
                    else:
                        buf += chunk
            except (OSError, ValueError):
                pass

        reader = threading.Thread(target=_reader, args=(process, progress_queue), daemon=True)
        reader.start()

        duration_secs: Optional[float] = None
        video_name = os.path.basename(video_path)
        while process.poll() is None:
            if self.cancelled:
                try:
                    process.terminate()
                except OSError:
                    pass
                break
            try:
                while True:
                    msg_type, data = progress_queue.get_nowait()
                    if msg_type == 'duration':
                        duration_secs = self._parse_duration(data)
                    elif msg_type == 'time' and duration_secs and duration_secs > 0:
                        current = self._parse_time(data)
                        if current is not None:
                            pct = min(current / duration_secs, 1.0)
                            update_progress(pct)
                            self.root.after(0, lambda n=video_name, p=pct: self.status_label.config(
                                text=f"正在压缩: {n} - {p * 100:.0f}%"
                            ))
            except queue.Empty:
                pass
            time.sleep(0.2)

        self.ffmpeg_process = None
        reader.join(timeout=1)

        if self.cancelled:
            if os.path.exists(output_path):
                os.remove(output_path)
            return None

        if process.returncode != 0:
            self._show_error(f"压缩失败: {video_name}")
            if os.path.exists(output_path):
                os.remove(output_path)
            return None

        orig_size = os.path.getsize(video_path)
        comp_size = os.path.getsize(output_path)
        return orig_size, comp_size, output_path

    def _compress_pptx(self, input_path: str, ffmpeg_path: str, crf: int, image_threshold: int, max_image_width: int) -> None:
        """压缩 PPTX 文件中的媒体"""
        temp_dir = tempfile.mkdtemp()
        try:
            self._update_status("正在解压 PPTX...")
            extract_dir = os.path.join(temp_dir, "extracted")
            with zipfile.ZipFile(input_path, 'r') as zip_ref:
                zip_ref.extractall(extract_dir)

            media_dir = os.path.join(extract_dir, "ppt", "media")
            if not os.path.exists(media_dir):
                self._show_error("PPTX 文件中没有找到媒体文件")
                return

            video_files: List[str] = []
            image_files: List[str] = []
            for root_dir, dirs, files in os.walk(media_dir):
                for f in files:
                    ext = Path(f).suffix.lower()
                    full = os.path.join(root_dir, f)
                    if ext in self.VIDEO_EXTENSIONS:
                        video_files.append(full)
                    elif ext in self.IMAGE_EXTENSIONS:
                        image_files.append(full)

            if not video_files and not image_files:
                self._show_error("PPTX 文件中没有找到视频或图片")
                return

            original_video_size = compressed_video_size = 0
            original_image_size = compressed_image_size = 0
            images_compressed = 0
            png_methods: dict = {}

            # 压缩图片
            if image_files and not self.cancelled:
                total_images = len(image_files)
                self._update_status("正在处理图片...")
                for i, image_path in enumerate(image_files):
                    if self.cancelled:
                        break
                    self._update_status(f"正在处理图片: {os.path.basename(image_path)} ({i + 1}/{total_images})")
                    original_size = os.path.getsize(image_path)
                    original_image_size += original_size
                    new_size, was_compressed, method = self._compress_image(image_path, image_threshold, max_image_width)
                    compressed_image_size += new_size
                    if was_compressed:
                        images_compressed += 1
                        if method:
                            png_methods[method] = png_methods.get(method, 0) + 1
                    self.root.after(0, lambda p=(i + 1) / total_images * 50: self.progress_var.set(p))

            # 压缩视频
            if video_files and not self.cancelled:
                total_videos = len(video_files)
                for i, video_path in enumerate(video_files):
                    if self.cancelled:
                        break
                    self._update_status(f"正在压缩视频: {os.path.basename(video_path)} ({i + 1}/{total_videos})")
                    result = self._compress_single_video(
                        video_path, ffmpeg_path, crf,
                        lambda pct, idx=i, total=total_videos: self.root.after(
                            0, lambda p=50 + (idx + pct) / total * 50: self.progress_var.set(p)
                        ),
                    )
                    if result is None:
                        if self.cancelled:
                            break
                        return
                    orig_size, comp_size, output_path = result
                    original_video_size += orig_size
                    if comp_size < orig_size:
                        os.remove(video_path)
                        os.rename(output_path, video_path)
                        compressed_video_size += comp_size
                    else:
                        os.remove(output_path)
                        compressed_video_size += orig_size
                    self.root.after(0, lambda p=50 + (i + 1) / total_videos * 50: self.progress_var.set(p))

            if self.cancelled:
                self._update_status("已取消")
                return

            self._update_status("正在打包 PPTX...")

            input_dir = os.path.dirname(input_path)
            name_without_ext = os.path.splitext(os.path.basename(input_path))[0]
            output_path = os.path.join(input_dir, f"{name_without_ext}_compressed.pptx")
            counter = 1
            while os.path.exists(output_path):
                output_path = os.path.join(input_dir, f"{name_without_ext}_compressed_{counter}.pptx")
                counter += 1

            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root_dir, dirs, files in os.walk(extract_dir):
                    for file in files:
                        file_path = os.path.join(root_dir, file)
                        zipf.write(file_path, os.path.relpath(file_path, extract_dir))

            result_text = "压缩完成！\n\n"
            if video_files:
                ov = original_video_size / 1024 / 1024
                cv = compressed_video_size / 1024 / 1024
                vr = (1 - compressed_video_size / original_video_size) * 100 if original_video_size > 0 else 0
                result_text += (
                    f"【视频】\n"
                    f"  原始大小: {ov:.2f} MB\n"
                    f"  压缩后: {cv:.2f} MB\n"
                    f"  压缩率: {vr:.1f}%\n"
                    f"  共 {len(video_files)} 个视频\n\n"
                )
            if image_files:
                oi = original_image_size / 1024 / 1024
                ci = compressed_image_size / 1024 / 1024
                ir = (1 - compressed_image_size / original_image_size) * 100 if original_image_size > 0 else 0
                result_text += (
                    f"【图片】\n"
                    f"  原始大小: {oi:.2f} MB\n"
                    f"  压缩后: {ci:.2f} MB\n"
                    f"  压缩率: {ir:.1f}%\n"
                    f"  共 {len(image_files)} 张图片，{images_compressed} 张被压缩\n"
                )
                if png_methods:
                    result_text += f"  压缩方式: {', '.join(f'{k}: {v}' for k, v in png_methods.items())}\n"
                result_text += "\n"
            total_o = original_video_size + original_image_size
            total_c = compressed_video_size + compressed_image_size
            total_r = (1 - total_c / total_o) * 100 if total_o > 0 else 0
            result_text += f"【总计】\n  整体压缩率: {total_r:.1f}%\n\n输出文件: {os.path.basename(output_path)}"

            self.root.after(0, lambda: self._update_status("压缩完成！"))
            self.root.after(0, lambda: self._show_result(result_text))
            self.root.after(0, lambda: self.progress_var.set(100))
        finally:
            try:
                shutil.rmtree(temp_dir)
            except OSError:
                pass

    def _restore_compress_btn(self) -> None:
        """恢复压缩按钮状态"""
        self.cancel_btn.pack_forget()
        self.compress_btn.pack(side="left", padx=5)

    def _update_status(self, text: str) -> None:
        """更新状态文本"""
        self.root.after(0, lambda: self.status_label.config(text=text))

    def _show_result(self, text: str) -> None:
        """显示结果文本"""
        self.result_text.config(state="normal")
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, text)
        self.result_text.config(state="disabled")

    def _clear_result(self) -> None:
        """清空结果文本"""
        self.result_text.config(state="normal")
        self.result_text.delete(1.0, tk.END)
        self.result_text.config(state="disabled")

    def _show_error(self, message: str) -> None:
        """显示错误消息"""
        self.root.after(0, lambda: messagebox.showerror("错误", message))
        self.root.after(0, lambda: self.status_label.config(text=""))

    @staticmethod
    def _parse_duration(line: str) -> Optional[float]:
        """从 ffmpeg 的 Duration 行解析总时长（秒）"""
        try:
            marker = "Duration:"
            idx = line.index(marker)
            time_str = line[idx + len(marker):].split(",")[0].strip()
            parts = time_str.split(":")
            return float(parts[0]) * 3600 + float(parts[1]) * 60 + float(parts[2])
        except (ValueError, IndexError):
            return None

    @staticmethod
    def _parse_time(line: str) -> Optional[float]:
        """从 ffmpeg 进度行解析当前时间（秒）"""
        try:
            marker = "time="
            idx = line.index(marker)
            time_str = line[idx + len(marker):].split(" ")[0].strip()
            parts = time_str.split(":")
            return float(parts[0]) * 3600 + float(parts[1]) * 60 + float(parts[2])
        except (ValueError, IndexError):
            return None


def main() -> None:
    """Main entry point"""
    root = tk.Tk()
    app = PPTXCompressor(root)
    root.mainloop()


if __name__ == "__main__":
    main()
