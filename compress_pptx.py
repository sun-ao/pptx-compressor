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
        self.root.title("PPTX Compressor")
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

        self._create_widgets()

    def _create_widgets(self) -> None:
        """创建界面组件"""
        # 文件选择区域
        file_frame = ttk.LabelFrame(self.root, text="File Selection", padding=10)
        file_frame.pack(fill="x", padx=10, pady=10)

        ttk.Entry(file_frame, textvariable=self.input_file, width=50).pack(side="left", padx=(0, 10))
        ttk.Button(file_frame, text="Browse", command=self._select_file).pack(side="left")

        # 高级选项区域
        self.options_frame = ttk.LabelFrame(self.root, text="Advanced Options", padding=10)
        self.options_frame.pack(fill="x", padx=10, pady=5)

        # CRF 值
        crf_frame = ttk.Frame(self.options_frame)
        crf_frame.pack(fill="x", pady=2)
        ttk.Label(crf_frame, text="CRF (18-51):", width=15).pack(side="left")
        ttk.Entry(crf_frame, textvariable=self.crf_value, width=10).pack(side="left")
        ttk.Label(crf_frame, text="Higher = more compression, lower quality").pack(side="left", padx=10)

        # Preset
        preset_frame = ttk.Frame(self.options_frame)
        preset_frame.pack(fill="x", pady=2)
        ttk.Label(preset_frame, text="Preset:", width=15).pack(side="left")
        preset_combo = ttk.Combobox(preset_frame, textvariable=self.preset_value, width=15, state="readonly")
        preset_combo['values'] = ('ultrafast', 'superfast', 'veryfast', 'faster', 'fast', 'medium', 'slow', 'slower', 'veryslow')
        preset_combo.pack(side="left")
        ttk.Label(preset_frame, text="Slower = better compression").pack(side="left", padx=10)

        # 音频码率
        audio_frame = ttk.Frame(self.options_frame)
        audio_frame.pack(fill="x", pady=2)
        ttk.Label(audio_frame, text="Audio Bitrate:", width=15).pack(side="left")
        ttk.Entry(audio_frame, textvariable=self.audio_bitrate, width=10).pack(side="left")

        # 图片压缩选项
        image_frame = ttk.LabelFrame(self.root, text="Image Compression", padding=10)
        image_frame.pack(fill="x", padx=10, pady=5)

        # 图片大小阈值
        threshold_frame = ttk.Frame(image_frame)
        threshold_frame.pack(fill="x", pady=2)
        ttk.Label(threshold_frame, text="Size Threshold (KB):", width=18).pack(side="left")
        ttk.Entry(threshold_frame, textvariable=self.image_threshold, width=10).pack(side="left")
        ttk.Label(threshold_frame, text="Images larger than this will be compressed").pack(side="left", padx=10)

        # 最大宽度
        maxwidth_frame = ttk.Frame(image_frame)
        maxwidth_frame.pack(fill="x", pady=2)
        ttk.Label(maxwidth_frame, text="Max Width (px):", width=18).pack(side="left")
        ttk.Entry(maxwidth_frame, textvariable=self.max_image_width, width=10).pack(side="left")
        ttk.Label(maxwidth_frame, text="Oversized images will be scaled down").pack(side="left", padx=10)

        # 压缩按钮
        self.compress_btn = ttk.Button(self.root, text="Start Compression", command=self._start_compress)
        self.compress_btn.pack(pady=15)

        # 进度区域
        progress_frame = ttk.LabelFrame(self.root, text="Progress", padding=10)
        progress_frame.pack(fill="x", padx=10, pady=5)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill="x")

        self.status_label = ttk.Label(progress_frame, text="")
        self.status_label.pack(pady=5)

        # 结果区域
        result_frame = ttk.LabelFrame(self.root, text="Result", padding=10)
        result_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.result_text = tk.Text(result_frame, height=8, width=60, state="disabled")
        self.result_text.pack(fill="both", expand=True)

    def _select_file(self) -> None:
        """选择 PPTX 文件"""
        file_path = filedialog.askopenfilename(
            title="Select PPTX File",
            filetypes=[("PowerPoint Files", "*.pptx"), ("All Files", "*.*")]
        )
        if file_path:
            self.input_file.set(file_path)

    def _start_compress(self) -> None:
        """开始压缩"""
        if self.compressing:
            return

        input_path = self.input_file.get()
        if not input_path:
            messagebox.showwarning("Warning", "Please select a PPTX file first")
            return

        if not os.path.exists(input_path):
            messagebox.showerror("Error", "File does not exist")
            return

        self.compressing = True
        self.compress_btn.config(state="disabled")
        self._clear_result()
        self.progress_var.set(0)

        # 在新线程中执行压缩
        thread = threading.Thread(target=self._compress, args=(input_path,))
        thread.daemon = True
        thread.start()

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
                if mode in ('RGBA', 'LA', 'P'):
                    if mode == 'P' and 'transparency' not in img.info:
                        img = img.convert('RGB')
                    # 保持 RGBA 模式保存
                    img.save(image_path, 'PNG', optimize=True, compress_level=9)
                else:
                    img.convert('RGB').save(image_path, 'PNG', optimize=True, compress_level=9)

                new_size = os.path.getsize(image_path)
                if new_size < original_size:
                    return new_size, True, "PIL Lossless"

            return original_size, False, None

        except Exception:  # pylint: disable=broad-except
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
                    # 先转换为标准格式
                    if mode in ('RGBA', 'LA', 'P'):
                        save_mode = 'RGBA'
                    elif mode == 'P':
                        if 'transparency' in img.info:
                            img = img.convert('RGBA')
                            save_mode = 'RGBA'
                        else:
                            img = img.convert('RGB')
                            save_mode = 'RGB'
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

        except Exception:  # pylint: disable=broad-except
            # 压缩失败，保留原文件
            return original_size, False, None

    def _compress(self, input_path: str) -> None:
        """执行压缩"""
        try:
            ffmpeg_path = self._get_ffmpeg_path()

            if not os.path.exists(ffmpeg_path):
                self._show_error("ffmpeg.exe not found. Please ensure ffmpeg.exe is in the same directory as the program.")
                return

            # 验证参数
            try:
                crf = int(self.crf_value.get())
                if crf < 18 or crf > 51:
                    raise ValueError("CRF out of range")
            except ValueError:
                self._show_error("CRF value must be an integer between 18 and 51")
                return

            # 验证图片参数
            try:
                image_threshold = int(self.image_threshold.get())
                if image_threshold < 0:
                    raise ValueError("Image threshold invalid")
            except ValueError:
                self._show_error("Image threshold must be a positive integer")
                return

            try:
                max_image_width = int(self.max_image_width.get())
                if max_image_width < 100:
                    raise ValueError("Max width invalid")
            except ValueError:
                self._show_error("Max width must be an integer greater than 100")
                return

            # 创建临时目录
            temp_dir = tempfile.mkdtemp()

            try:
                # 解压 PPTX
                self._update_status("Extracting PPTX...")
                extract_dir = os.path.join(temp_dir, "extracted")
                with zipfile.ZipFile(input_path, 'r') as zip_ref:
                    zip_ref.extractall(extract_dir)

                # 查找视频文件（递归搜索所有子目录）
                media_dir = os.path.join(extract_dir, "ppt", "media")
                if not os.path.exists(media_dir):
                    self._show_error("No media files found in PPTX")
                    return

                video_files: List[str] = []
                for root_dir, dirs, files in os.walk(media_dir):
                    for f in files:
                        if Path(f).suffix.lower() in self.VIDEO_EXTENSIONS:
                            video_files.append(os.path.join(root_dir, f))

                # 查找图片文件
                image_files: List[str] = []
                for root_dir, dirs, files in os.walk(media_dir):
                    for f in files:
                        if Path(f).suffix.lower() in self.IMAGE_EXTENSIONS:
                            image_files.append(os.path.join(root_dir, f))

                if not video_files and not image_files:
                    self._show_error("No video or image files found in PPTX")
                    return

                # 统计数据
                original_video_size = 0
                compressed_video_size = 0
                original_image_size = 0
                compressed_image_size = 0
                images_compressed = 0
                png_methods: dict = {}  # 记录PNG压缩方式统计

                # 先压缩图片
                if image_files:
                    total_images = len(image_files)
                    self._update_status("Processing images...")

                    for i, image_path in enumerate(image_files):
                        image_name = os.path.basename(image_path)
                        self._update_status(f"Processing image: {image_name} ({i+1}/{total_images})")

                        original_size = os.path.getsize(image_path)
                        original_image_size += original_size

                        new_size, was_compressed, compress_method = self._compress_image(
                            image_path, image_threshold, max_image_width
                        )
                        compressed_image_size += new_size
                        if was_compressed:
                            images_compressed += 1
                            # 记录压缩方式
                            if compress_method:
                                png_methods[compress_method] = png_methods.get(compress_method, 0) + 1

                        # 更新进度
                        progress = (i + 1) / total_images * 50  # 图片占前50%
                        self.root.after(0, lambda p=progress: self.progress_var.set(p))

                # 压缩视频
                if video_files:
                    total_videos = len(video_files)

                    for i, video_path in enumerate(video_files):
                        video_name = os.path.basename(video_path)
                        self._update_status(f"Compressing video: {video_name} ({i+1}/{total_videos})")

                        # 创建临时输出文件
                        output_path = video_path + ".compressed.mp4"

                        # 构建 ffmpeg 命令
                        cmd = [
                            ffmpeg_path,
                            '-i', video_path,
                            '-c:v', 'libx264',
                            '-crf', str(crf),
                            '-preset', self.preset_value.get(),
                            '-profile:v', 'high',
                            '-level', '4.1',
                            '-pix_fmt', 'yuv420p',
                            '-c:a', 'aac',
                            '-b:a', self.audio_bitrate.get(),
                            '-movflags', '+faststart',
                            '-y',  # 覆盖输出文件
                            output_path
                        ]

                        # 执行压缩
                        process = subprocess.Popen(
                            cmd,
                            stdout=subprocess.DEVNULL,
                            stderr=subprocess.DEVNULL,
                            creationflags=subprocess.CREATE_NO_WINDOW
                        )
                        process.wait()

                        if process.returncode != 0:
                            self._show_error(f"Failed to compress {video_name}")
                            return

                        # 比较大小，只有压缩后才替换
                        orig_size = os.path.getsize(video_path)
                        comp_size = os.path.getsize(output_path)
                        original_video_size += orig_size

                        if comp_size < orig_size:
                            # 压缩后更小，替换原文件
                            os.remove(video_path)
                            # 保持原扩展名
                            final_path = video_path
                            os.rename(output_path, final_path)
                            compressed_video_size += comp_size
                        else:
                            # 压缩后更大或相等，保留原视频
                            os.remove(output_path)
                            compressed_video_size += orig_size

                        # 更新进度 (视频占后50%)
                        progress = 50 + (i + 1) / total_videos * 50
                        self.root.after(0, lambda p=progress: self.progress_var.set(p))

                # 重新打包 PPTX
                self._update_status("Packaging PPTX...")

                # 生成输出文件名
                input_dir = os.path.dirname(input_path)
                input_name = os.path.basename(input_path)
                name_without_ext = os.path.splitext(input_name)[0]
                output_path = os.path.join(input_dir, f"{name_without_ext}_compressed.pptx")

                # 如果输出文件已存在，添加数字后缀
                counter = 1
                while os.path.exists(output_path):
                    output_path = os.path.join(input_dir, f"{name_without_ext}_compressed_{counter}.pptx")
                    counter += 1

                # 创建新的 PPTX
                with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for root_dir, dirs, files in os.walk(extract_dir):
                        for file in files:
                            file_path = os.path.join(root_dir, file)
                            arcname = os.path.relpath(file_path, extract_dir)
                            zipf.write(file_path, arcname)

                # 显示结果
                result_text = "Compression completed!\n\n"

                # 视频统计
                if video_files:
                    original_video_mb = original_video_size / 1024 / 1024
                    compressed_video_mb = compressed_video_size / 1024 / 1024
                    video_ratio = (1 - compressed_video_size / original_video_size) * 100 if original_video_size > 0 else 0
                    result_text += f"【Video】\n"
                    result_text += f"  Original size: {original_video_mb:.2f} MB\n"
                    result_text += f"  Compressed: {compressed_video_mb:.2f} MB\n"
                    result_text += f"  Ratio: {video_ratio:.1f}%\n"
                    result_text += f"  Total: {len(video_files)} video file(s)\n\n"

                # 图片统计
                if image_files:
                    original_image_mb = original_image_size / 1024 / 1024
                    compressed_image_mb = compressed_image_size / 1024 / 1024
                    image_ratio = (1 - compressed_image_size / original_image_size) * 100 if original_image_size > 0 else 0
                    result_text += f"【Image】\n"
                    result_text += f"  Original size: {original_image_mb:.2f} MB\n"
                    result_text += f"  Compressed: {compressed_image_mb:.2f} MB\n"
                    result_text += f"  Ratio: {image_ratio:.1f}%\n"
                    result_text += f"  Processed {len(image_files)} image(s), {images_compressed} compressed\n"
                    # 显示压缩方式统计
                    if png_methods:
                        methods_str = ", ".join([f"{k}: {v}" for k, v in png_methods.items()])
                        result_text += f"  Methods: {methods_str}\n"
                    result_text += "\n"

                # 总计
                total_original = original_video_size + original_image_size
                total_compressed = compressed_video_size + compressed_image_size
                total_ratio = (1 - total_compressed / total_original) * 100 if total_original > 0 else 0
                result_text += f"【Total】\n"
                result_text += f"  Compression ratio: {total_ratio:.1f}%\n\n"
                result_text += f"Output: {os.path.basename(output_path)}"

                # 更新状态为压缩成功
                self.root.after(0, lambda: self._update_status("Compression completed!"))
                self.root.after(0, lambda: self._show_result(result_text))
                self.root.after(0, lambda: self.progress_var.set(100))

            finally:
                # 清理临时目录
                try:
                    shutil.rmtree(temp_dir)
                except OSError:
                    pass

        except Exception as e:  # pylint: disable=broad-except
            self._show_error(f"Compression failed: {str(e)}")
        finally:
            self.compressing = False
            self.root.after(0, lambda: self.compress_btn.config(state="normal"))

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
        self.root.after(0, lambda: messagebox.showerror("Error", message))
        self.root.after(0, lambda: self.status_label.config(text=""))


def main() -> None:
    """Main entry point"""
    root = tk.Tk()
    app = PPTXCompressor(root)
    root.mainloop()


if __name__ == "__main__":
    main()
