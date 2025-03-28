import os
import sys
import io
import tempfile
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Scale
from PIL import Image, ImageTk, ImageEnhance, ImageOps
import threading

# 添加TkinterDnD2支持
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    # 如果没有安装tkinterdnd2，给出友好提示
    class TkinterDnD:
        @staticmethod
        def Tk():
            root = tk.Tk()
            messagebox.showwarning("缺少依赖", "未检测到tkinterdnd2库，拖放功能将不可用。\n请使用pip install tkinterdnd2安装。")
            return root
    DND_FILES = "<<DROP>>"

class ImageTrimmerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("图片素材修剪工具")
        self.root.geometry("1080x720")
        
        # 设置变量
        self.source_path = tk.StringVar()
        self.target_path = tk.StringVar()
        self.new_filename = tk.StringVar()
        self.format_type = tk.StringVar(value="PNG")
        self.width = tk.IntVar()
        self.height = tk.IntVar()
        self.original_image = None
        self.display_image = None
        self.original_width = 0
        self.original_height = 0
        self.crop_rect = None        # 裁剪矩形框引用
        self.crop_start = None       # 裁剪起始点
        self.is_cropping = False     # 是否正在裁剪标志
        self.zoom_scale = 1.0        # 缩放比例
        self.zoom_sensitivity = 0.2  # 增加缩放灵敏度，从0.1改为0.2
        self.preview_width = 0       # 预览图宽度
        self.preview_height = 0      # 预览图高度
        self.operation_mode = tk.StringVar(value="scale")  # 默认为缩放模式
        self.crop_offset = (0, 0)  # 添加这一行，用于跟踪裁剪框拖动偏移量
        self.image_on_canvas = None   # 画布上的图像引用
        self.preview_image = None     # 预览图像引用
        
        # 添加图像变换相关变量
        self.rotation_angle = 0  # 旋转角度
        self.is_flipped_h = False  # 水平翻转标志
        self.is_flipped_v = False  # 垂直翻转标志
        
        # 添加色彩调整相关变量
        self.brightness_value = tk.DoubleVar(value=1.0)  # 亮度值
        self.contrast_value = tk.DoubleVar(value=1.0)    # 对比度值
        self.saturation_value = tk.DoubleVar(value=1.0)  # 饱和度值
        
        # 是否需要重新渲染
        self.need_rerender = False
        
        # 添加裁剪预设
        self.crop_preset = tk.StringVar(value="自定义")
        
        # 创建界面
        self.create_widgets()
        
        # 设置拖放标识和提示
        self.drag_prompt = None
        self.setup_drag_drop()
        
    def create_widgets(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 左侧控制面板
        control_frame = ttk.LabelFrame(main_frame, text="控制面板", padding="10")
        control_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        
        # 源文件选择
        ttk.Label(control_frame, text="源文件:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(control_frame, textvariable=self.source_path, width=30).grid(row=0, column=1, pady=5)
        ttk.Button(control_frame, text="浏览...", command=self.browse_source).grid(row=0, column=2, padx=5, pady=5)
        
        # 目标文件夹选择
        ttk.Label(control_frame, text="目标文件夹:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(control_frame, textvariable=self.target_path, width=30).grid(row=1, column=1, pady=5)
        ttk.Button(control_frame, text="浏览...", command=self.browse_target).grid(row=1, column=2, padx=5, pady=5)
        
        # 新文件名
        ttk.Label(control_frame, text="新文件名:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(control_frame, textvariable=self.new_filename, width=30).grid(row=2, column=1, columnspan=2, pady=5, sticky=tk.W)
        
        # 格式选择 - 添加图标格式
        ttk.Label(control_frame, text="输出格式:").grid(row=3, column=0, sticky=tk.W, pady=5)
        format_combobox = ttk.Combobox(control_frame, textvariable=self.format_type, 
                                    values=["PNG", "JPEG", "GIF", "BMP", "TIFF", "ICO", "ICNS", "PNG图标集"])
        format_combobox.grid(row=3, column=1, columnspan=2, pady=5, sticky=tk.W)
        format_combobox.current(0)
        
        # 操作模式选择
        mode_frame = ttk.LabelFrame(control_frame, text="操作模式", padding="5")
        mode_frame.grid(row=4, column=0, columnspan=3, pady=10, sticky=tk.W+tk.E)
        
        ttk.Radiobutton(mode_frame, text="缩放", variable=self.operation_mode, 
                        value="scale", command=self.mode_changed).grid(row=0, column=0, padx=5, pady=5)
        ttk.Radiobutton(mode_frame, text="裁剪", variable=self.operation_mode, 
                        value="crop", command=self.mode_changed).grid(row=0, column=1, padx=5, pady=5)
        ttk.Radiobutton(mode_frame, text="缩放+裁剪", variable=self.operation_mode, 
                        value="both", command=self.mode_changed).grid(row=0, column=2, padx=5, pady=5)
        
        # 在操作模式框架中添加裁剪预设
        preset_frame = ttk.Frame(mode_frame)
        preset_frame.grid(row=1, column=0, columnspan=3, padx=5, pady=5, sticky=tk.W+tk.E)
        
        ttk.Label(preset_frame, text="裁剪预设:").pack(side=tk.LEFT, padx=5)
        preset_combobox = ttk.Combobox(preset_frame, textvariable=self.crop_preset, 
                                      values=["自定义", "正方形 (1:1)", "Instagram (4:5)", 
                                              "Facebook (16:9)", "Twitter (2:1)", 
                                              "LinkedIn (1.91:1)", "微信 (4:3)"])
        preset_combobox.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        preset_combobox.bind("<<ComboboxSelected>>", self.apply_crop_preset)
        
        # 添加显示裁剪框按钮
        self.show_crop_button = ttk.Button(control_frame, text="显示裁剪框", command=self.show_crop_box)
        self.show_crop_button.grid(row=5, column=0, columnspan=3, pady=5)
        self.show_crop_button.grid_remove()  # 初始隐藏
        
        # 尺寸设置
        size_frame = ttk.LabelFrame(control_frame, text="输出尺寸", padding="5")
        size_frame.grid(row=6, column=0, columnspan=3, pady=10, sticky=tk.W+tk.E)
        
        ttk.Label(size_frame, text="宽度:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(size_frame, textvariable=self.width, width=10).grid(row=0, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(size_frame, text="高度:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(size_frame, textvariable=self.height, width=10).grid(row=1, column=1, sticky=tk.W, pady=5)
        
        # 基础编辑模块 - 移到按钮上方
        basic_edit_frame = ttk.LabelFrame(control_frame, text="基础编辑", padding="5")
        basic_edit_frame.grid(row=7, column=0, columnspan=3, pady=10, sticky=tk.W+tk.E)
        
        # 添加图像变换按钮
        transform_frame = ttk.Frame(basic_edit_frame)
        transform_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(transform_frame, text="图像变换:").pack(side=tk.LEFT, padx=5)
        ttk.Button(transform_frame, text="向左旋转", command=lambda: self.rotate_image(-90)).pack(side=tk.LEFT, padx=2)
        ttk.Button(transform_frame, text="向右旋转", command=lambda: self.rotate_image(90)).pack(side=tk.LEFT, padx=2)
        ttk.Button(transform_frame, text="水平翻转", command=self.flip_horizontal).pack(side=tk.LEFT, padx=2)
        ttk.Button(transform_frame, text="垂直翻转", command=self.flip_vertical).pack(side=tk.LEFT, padx=2)
        
        # 添加色彩调整滑块
        color_frame = ttk.Frame(basic_edit_frame)
        color_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 亮度滑块
        brightness_frame = ttk.Frame(color_frame)
        brightness_frame.pack(fill=tk.X, pady=2)
        ttk.Label(brightness_frame, text="亮度:").pack(side=tk.LEFT, padx=5)
        brightness_slider = ttk.Scale(brightness_frame, from_=0.5, to=1.5, 
                                     variable=self.brightness_value, orient=tk.HORIZONTAL, 
                                     length=150, command=self.update_color_adjustments)
        brightness_slider.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # 对比度滑块
        contrast_frame = ttk.Frame(color_frame)
        contrast_frame.pack(fill=tk.X, pady=2)
        ttk.Label(contrast_frame, text="对比度:").pack(side=tk.LEFT, padx=5)
        contrast_slider = ttk.Scale(contrast_frame, from_=0.5, to=1.5, 
                                   variable=self.contrast_value, orient=tk.HORIZONTAL, 
                                   length=150, command=self.update_color_adjustments)
        contrast_slider.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # 饱和度滑块
        saturation_frame = ttk.Frame(color_frame)
        saturation_frame.pack(fill=tk.X, pady=2)
        ttk.Label(saturation_frame, text="饱和度:").pack(side=tk.LEFT, padx=5)
        saturation_slider = ttk.Scale(saturation_frame, from_=0.0, to=2.0, 
                                     variable=self.saturation_value, orient=tk.HORIZONTAL, 
                                     length=150, command=self.update_color_adjustments)
        saturation_slider.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # 重置色彩按钮
        reset_color_button = ttk.Button(color_frame, text="重置色彩", command=self.reset_color_adjustments)
        reset_color_button.pack(pady=5)
        
        # 功能按钮 - 移到编辑框架下方
        buttons_frame = ttk.Frame(control_frame)
        buttons_frame.grid(row=8, column=0, columnspan=3, pady=10)
        
        ttk.Button(buttons_frame, text="预览修改", command=self.preview_changes).grid(row=0, column=0, padx=5)
        ttk.Button(buttons_frame, text="应用更改", command=self.apply_changes).grid(row=0, column=1, padx=5)
        ttk.Button(buttons_frame, text="保存图片", command=self.save_image).grid(row=0, column=2, padx=5)
        ttk.Button(buttons_frame, text="复原图片", command=self.reset_image).grid(row=0, column=3, padx=5)
        
        # 图片信息
        self.info_frame = ttk.LabelFrame(control_frame, text="图片信息", padding="5")
        self.info_frame.grid(row=9, column=0, columnspan=3, pady=10, sticky=tk.W+tk.E)
        
        self.file_info_label = ttk.Label(self.info_frame, text="文件: 未选择")
        self.file_info_label.grid(row=0, column=0, sticky=tk.W, pady=2)
        
        self.size_info_label = ttk.Label(self.info_frame, text="大小: -")
        self.size_info_label.grid(row=1, column=0, sticky=tk.W, pady=2)
        
        self.dim_info_label = ttk.Label(self.info_frame, text="原始尺寸: -")
        self.dim_info_label.grid(row=2, column=0, sticky=tk.W, pady=2)
        
        self.format_info_label = ttk.Label(self.info_frame, text="格式: -")
        self.format_info_label.grid(row=3, column=0, sticky=tk.W, pady=2)
        
        # 预览尺寸信息
        self.preview_dim_label = ttk.Label(self.info_frame, text="预览尺寸: -")
        self.preview_dim_label.grid(row=4, column=0, sticky=tk.W, pady=2)
        
        # 缩放比例信息
        self.zoom_info_label = ttk.Label(self.info_frame, text="缩放比例: 1.0x")
        self.zoom_info_label.grid(row=5, column=0, sticky=tk.W, pady=2)
        
        # 裁剪框坐标信息
        self.crop_coords_label = ttk.Label(self.info_frame, text="裁剪框坐标: -")
        self.crop_coords_label.grid(row=6, column=0, sticky=tk.W, pady=2)
        
        # 右侧图片预览
        preview_frame = ttk.LabelFrame(main_frame, text="图片预览", padding="10")
        preview_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # 创建一个能够滚动的画布来放置图片
        self.canvas = tk.Canvas(preview_frame, bg="white")
        h_scrollbar = ttk.Scrollbar(preview_frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        v_scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        
        self.canvas.config(xscrollcommand=h_scrollbar.set, yscrollcommand=v_scrollbar.set)
        
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 修改为直接在画布上创建图像，而不是使用标签
        self.image_on_canvas = None
        
        # 绑定鼠标滚轮事件用于缩放图片
        self.canvas.bind("<MouseWheel>", self.zoom_image)  # Windows
        self.canvas.bind("<Button-4>", self.zoom_image)    # Linux上滚
        self.canvas.bind("<Button-5>", self.zoom_image)    # Linux下滚
        
        # 在图片预览区域添加鼠标事件绑定
        self.canvas.bind("<ButtonPress-1>", self.start_crop)
        self.canvas.bind("<B1-Motion>", self.update_crop)
        self.canvas.bind("<ButtonRelease-1>", self.end_crop)
        
        # 状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 绑定窗口大小调整事件
        self.root.bind("<Configure>", self.on_window_resize)
    
    def browse_source(self):
        file_path = filedialog.askopenfilename(
            title="选择图片文件",
            filetypes=[
                ("图片文件", "*.png *.jpg *.jpeg *.gif *.bmp *.tiff"),
                ("所有文件", "*.*")
            ]
        )
        if file_path:
            self.source_path.set(file_path)
            self.load_image()
            
            # 自动填充新文件名（不带扩展名）
            base_name = os.path.basename(file_path)
            name_without_ext = os.path.splitext(base_name)[0]
            self.new_filename.set(name_without_ext)
            
            # 如果未设置目标路径，默认使用源文件的目录
            if not self.target_path.get():
                self.target_path.set(os.path.dirname(file_path))
    
    def browse_target(self):
        folder_path = filedialog.askdirectory(title="选择目标文件夹")
        if folder_path:
            self.target_path.set(folder_path)
    
    def load_image(self):
        # 当加载新图片时，移除拖放提示
        if self.drag_prompt:
            self.canvas.delete(self.drag_prompt)
            self.drag_prompt = None
        
        try:
            self.status_var.set("正在加载图片...")
            self.root.update()
            
            path = self.source_path.get()
            if not path or not os.path.exists(path):
                messagebox.showerror("错误", "请选择有效的图片文件")
                self.status_var.set("就绪")
                return
            
            # 使用PIL打开图片
            self.original_image = Image.open(path)
            self.original_width, self.original_height = self.original_image.size
            
            # 更新输入框中的图片尺寸
            self.width.set(self.original_width)
            self.height.set(self.original_height)
            
            # 显示图片信息
            self.update_image_info(path)
            
            # 调整图片以适应显示区域
            self.display_image = self.original_image.copy()
            self.update_preview()
            
            # 确保加载新图片时清除已有的裁剪框
            if self.crop_rect:
                self.canvas.delete(self.crop_rect)
                self.crop_rect = None
                self.update_crop_coords_display(None)
            
            # 重置图像变换状态
            self.rotation_angle = 0
            self.is_flipped_h = False
            self.is_flipped_v = False
            
            # 重置色彩调整
            self.brightness_value.set(1.0)
            self.contrast_value.set(1.0)
            self.saturation_value.set(1.0)
            
            # 重置裁剪预设
            self.crop_preset.set("自定义")
            
            self.status_var.set("图片已加载")
        except Exception as e:
            messagebox.showerror("错误", f"加载图片时出错: {str(e)}")
            self.status_var.set("加载失败")
    
    def update_image_info(self, path):
        # 获取文件大小
        file_size = os.path.getsize(path)
        size_str = self.format_file_size(file_size)
        
        # 获取图片格式
        img_format = self.original_image.format or "未知"
        
        # 更新标签
        self.file_info_label.config(text=f"文件: {os.path.basename(path)}")
        self.size_info_label.config(text=f"大小: {size_str}")
        self.dim_info_label.config(text=f"原始尺寸: {self.original_width} x {self.original_height} 像素")
        self.format_info_label.config(text=f"格式: {img_format}")
        
        # 初始化预览尺寸和缩放信息
        self.preview_width = self.original_width
        self.preview_height = self.original_height
        self.preview_dim_label.config(text=f"预览尺寸: {self.preview_width} x {self.preview_height} 像素")
        self.zoom_info_label.config(text=f"缩放比例: {self.zoom_scale:.1f}x")
    
    def format_file_size(self, size_in_bytes):
        # 转换字节为更易读的格式
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size_in_bytes < 1024.0 or unit == 'GB':
                return f"{size_in_bytes:.2f} {unit}"
            size_in_bytes /= 1024.0
    
    def update_preview(self):
        if self.display_image:
            # 将PIL图像转换为Tkinter可用的PhotoImage
            preview = ImageTk.PhotoImage(self.display_image)
            
            # 保存对图像的引用，防止垃圾回收
            self.preview_image = preview
            
            # 如果已有图像，删除它
            if self.image_on_canvas:
                self.canvas.delete(self.image_on_canvas)
            
            # 计算图像在画布中的中心位置
            canvas_width = self.canvas.winfo_width()
            canvas_height = self.canvas.winfo_height()
            image_width, image_height = self.display_image.size
            
            # 计算将图像放在画布中心的坐标
            x_position = max(0, (canvas_width - image_width) / 2)
            y_position = max(0, (canvas_height - image_height) / 2)
            
            # 直接在画布上创建图像，确保居中
            self.image_on_canvas = self.canvas.create_image(
                x_position, y_position, 
                anchor=tk.NW, 
                image=preview
            )
            
            # 更新画布滚动区域，确保图片完全可见
            total_width = max(canvas_width, image_width + x_position * 2)
            total_height = max(canvas_height, image_height + y_position * 2)
            self.canvas.config(scrollregion=(0, 0, total_width, total_height))
            
            # 更新预览尺寸信息
            self.preview_width = image_width
            self.preview_height = image_height
            self.preview_dim_label.config(text=f"预览尺寸: {image_width} x {image_height} 像素")
            
            # 删除之前的裁剪框并重新创建，保持其在相同的画布位置
            if self.crop_rect:
                current_coords = self.canvas.coords(self.crop_rect)
                self.canvas.delete(self.crop_rect)
                if len(current_coords) == 4:
                    self.crop_rect = self.canvas.create_rectangle(
                        current_coords[0], current_coords[1], 
                        current_coords[2], current_coords[3],
                        outline="red", width=2
                    )
                    # 更新坐标显示
                    self.update_crop_coords_display(current_coords)
    
    def zoom_image(self, event):
        if not self.original_image:
            return
        
        # 保存当前裁剪框的绝对位置（如果有）
        crop_coords = None
        if self.crop_rect:
            crop_coords = self.canvas.coords(self.crop_rect)
        
        # 计算缩放因子
        if hasattr(event, 'delta'):  # Windows平台
            if event.delta > 0:  # 向上滚动/放大
                scale_factor = 1 + self.zoom_sensitivity
            else:  # 向下滚动/缩小
                scale_factor = 1 - self.zoom_sensitivity
        else:  # Linux/Mac平台
            if event.num == 4:  # 向上滚动/放大
                scale_factor = 1 + self.zoom_sensitivity
            else:  # 向下滚动/缩小
                scale_factor = 1 - self.zoom_sensitivity
        
        # 更新缩放比例
        old_scale = self.zoom_scale
        self.zoom_scale *= scale_factor
        
        # 限制缩放范围
        if self.zoom_scale < 0.05:
            self.zoom_scale = 0.05
        elif self.zoom_scale > 10.0:
            self.zoom_scale = 10.0
        
        # 计算新尺寸并缩放图像
        new_width = int(self.original_width * self.zoom_scale)
        new_height = int(self.original_height * self.zoom_scale)
        self.display_image = self.original_image.resize((new_width, new_height), Image.LANCZOS)
        
        # 更新图像显示
        self.update_preview()
        
        # 如果之前有裁剪框，恢复其位置
        if crop_coords and len(crop_coords) == 4:
            # 直接使用之前的坐标重新创建裁剪框
            self.canvas.delete(self.crop_rect)
            self.crop_rect = self.canvas.create_rectangle(
                crop_coords[0], crop_coords[1], 
                crop_coords[2], crop_coords[3],
                outline="red", width=2
            )
            self.update_crop_coords_display(crop_coords)
        
        # 更新缩放信息
        self.zoom_info_label.config(text=f"缩放比例: {self.zoom_scale:.1f}x")
        self.status_var.set(f"缩放: {self.zoom_scale:.2f}x")
    
    def preview_changes(self):
        if not self.original_image:
            messagebox.showwarning("警告", "请先选择一张图片")
            return
        
        try:
            # 获取用户输入的新尺寸
            new_width = self.width.get()
            new_height = self.height.get()
            
            if new_width <= 0 or new_height <= 0:
                messagebox.showerror("错误", "宽度和高度必须为正数")
                return
            
            # 获取操作模式
            mode = self.operation_mode.get()
            processed_image = self.original_image.copy()
            
            # 在"both"模式下的处理顺序：先缩放原图，再进行裁剪
            if mode == 'both':
                # 先应用缩放
                scaled_w = int(self.original_width * self.zoom_scale)
                scaled_h = int(self.original_height * self.zoom_scale)
                processed_image = processed_image.resize((scaled_w, scaled_h), Image.LANCZOS)
                
                # 再应用裁剪
                if self.crop_rect:
                    bbox = self.canvas.coords(self.crop_rect)
                    if bbox and len(bbox) == 4:
                        # 获取图像在画布上的位置
                        img_coords = self.canvas.coords(self.image_on_canvas)
                        
                        if not img_coords or len(img_coords) < 2:
                            # 如果无法获取图像坐标，使用默认(0,0)
                            img_x, img_y = 0, 0
                        else:
                            img_x, img_y = img_coords[0], img_coords[1]
                        
                        # 裁剪框相对于图像的坐标
                        x1 = bbox[0] - img_x
                        y1 = bbox[1] - img_y
                        x2 = bbox[2] - img_x
                        y2 = bbox[3] - img_y
                        
                        # 确保坐标正确排序
                        x1, x2 = min(x1, x2), max(x1, x2)
                        y1, y2 = min(y1, y2), max(y1, y2)
                        
                        # 计算裁剪区域与图片的交集
                        img_width = processed_image.width
                        img_height = processed_image.height
                        
                        # 确保裁剪区域至少部分在图片内
                        if x1 < img_width and y1 < img_height and x2 > 0 and y2 > 0:
                            # 调整裁剪框到图片范围内的部分
                            x1 = max(0, x1)
                            y1 = max(0, y1)
                            x2 = min(img_width, x2)
                            y2 = min(img_height, y2)
                            
                            # 确保裁剪区域有效
                            if x2 > x1 and y2 > y1:
                                processed_image = processed_image.crop((x1, y1, x2, y2))
                                self.status_var.set(f"应用缩放({self.zoom_scale:.1f}x)和裁剪至{processed_image.width}x{processed_image.height}")
                            else:
                                # 如果裁剪区域太小，显示警告
                                messagebox.showwarning("警告", "裁剪框与图片的交集太小，将只应用缩放")
                                self.status_var.set(f"仅应用缩放({self.zoom_scale:.1f}x)")
                        else:
                            # 如果裁剪区域完全在图片外，显示警告
                            messagebox.showwarning("警告", "裁剪框在图片外或未与图片重叠，将只应用缩放")
                            self.status_var.set(f"仅应用缩放({self.zoom_scale:.1f}x)")
                else:
                    self.status_var.set(f"应用缩放({self.zoom_scale:.1f}x)")
                
            # 单独的缩放模式
            elif mode == 'scale':
                scaled_w = int(self.original_width * self.zoom_scale)
                scaled_h = int(self.original_height * self.zoom_scale)
                processed_image = processed_image.resize((scaled_w, scaled_h), Image.LANCZOS)
                self.status_var.set(f"应用缩放: {self.zoom_scale:.1f}x")
                
            # 单独的裁剪模式
            elif mode == 'crop' and self.crop_rect:
                bbox = self.canvas.coords(self.crop_rect)
                if bbox and len(bbox) == 4:
                    # 获取图像在画布上的位置
                    img_coords = self.canvas.coords(self.image_on_canvas)
                    
                    if not img_coords or len(img_coords) < 2:
                        # 如果无法获取图像坐标，使用默认(0,0)
                        img_x, img_y = 0, 0
                    else:
                        img_x, img_y = img_coords[0], img_coords[1]
                    
                    # 裁剪框相对于图像的坐标
                    x1 = bbox[0] - img_x
                    y1 = bbox[1] - img_y
                    x2 = bbox[2] - img_x
                    y2 = bbox[3] - img_y
                    
                    # 确保坐标正确排序
                    x1, x2 = min(x1, x2), max(x1, x2)
                    y1, y2 = min(y1, y2), max(y1, y2)
                    
                    # 计算裁剪区域与图片的交集
                    img_width = processed_image.width
                    img_height = processed_image.height
                    
                    # 确保裁剪区域至少部分在图片内
                    if x1 < img_width and y1 < img_height and x2 > 0 and y2 > 0:
                        # 调整裁剪框到图片范围内的部分
                        x1 = max(0, x1)
                        y1 = max(0, y1)
                        x2 = min(img_width, x2)
                        y2 = min(img_height, y2)
                        
                        # 确保裁剪区域有效
                        if x2 > x1 and y2 > y1:
                            processed_image = processed_image.crop((x1, y1, x2, y2))
                            self.status_var.set(f"应用裁剪: {processed_image.width}x{processed_image.height}")
                        else:
                            messagebox.showwarning("警告", "裁剪框与图片的交集太小，无法应用裁剪")
                            self.status_var.set("裁剪区域无效，未进行修改")
                            return
                    else:
                        messagebox.showwarning("警告", "裁剪框未与图片重叠，无法应用裁剪")
                        self.status_var.set("裁剪区域无效，未进行修改")
                        return
            
            # 调整到最终尺寸
            if processed_image.size != (new_width, new_height):
                processed_image = processed_image.resize((new_width, new_height), Image.LANCZOS)
            
            # 更新显示
            self.display_image = processed_image
            self.update_preview()
            
        except Exception as e:
            messagebox.showerror("错误", f"预览时出错: {str(e)}")
            self.status_var.set("预览失败")
    
    def save_image(self):
        # 检查是否有图像要保存
        if not self.display_image:
            messagebox.showwarning("警告", "没有可保存的图片")
            return
        
        try:
            # 获取目标路径和文件名
            target_dir = self.target_path.get()
            if not target_dir:
                messagebox.showerror("错误", "请指定目标文件夹")
                return
                
            if not os.path.exists(target_dir):
                os.makedirs(target_dir)  # 如果目录不存在，创建它
                
            filename = self.new_filename.get()
            if not filename:
                messagebox.showerror("错误", "请指定文件名")
                return
                
            # 获取输出格式
            format_str = self.format_type.get()
            
            # 检查是否是图标格式
            if "ICO" in format_str:
                # 保存为ICO图标
                save_path = os.path.join(target_dir, f"{filename}.ico")
                IconConverter.create_ico(self.display_image, save_path)
                self.status_var.set(f"图标已保存到: {save_path}")
                messagebox.showinfo("保存成功", f"图标已保存到:\n{save_path}")
                return
            elif "ICNS" in format_str:
                # 保存为ICNS图标
                save_path = os.path.join(target_dir, f"{filename}.icns")
                IconConverter.create_icns(self.display_image, save_path)
                self.status_var.set(f"图标已保存到: {save_path}")
                messagebox.showinfo("保存成功", f"图标已保存到:\n{save_path}")
                return
            elif "PNG图标集" in format_str:
                # 保存为PNG图标集
                save_path = os.path.join(target_dir, f"{filename}.png")
                icons_dir = self.export_png_icon_set(self.display_image, save_path)
                self.status_var.set(f"PNG图标集已保存到: {icons_dir}")
                messagebox.showinfo("保存成功", f"PNG图标集已保存到:\n{icons_dir}")
                return
            
            # 处理常规图像格式
            format_map = {
                "PNG": "PNG",
                "JPEG": "JPEG",
                "GIF": "GIF",
                "BMP": "BMP",
                "TIFF": "TIFF"
            }
            
            # 从格式选择中提取格式代码
            format_code = format_str.split()[0]  # 取第一个单词作为格式代码
            if format_code not in format_map:
                messagebox.showerror("错误", f"不支持的格式: {format_str}")
                return
                
            # 确定文件扩展名
            ext_map = {
                "PNG": ".png",
                "JPEG": ".jpg",
                "GIF": ".gif",
                "BMP": ".bmp",
                "TIFF": ".tiff"
            }
            
            # 选择正确的扩展名
            ext = ext_map.get(format_code, ".png")
            if not filename.lower().endswith(ext):
                filename += ext
                
            save_path = os.path.join(target_dir, filename)
            
            # 保存图像
            self.display_image.save(save_path, format=format_map[format_code])
            self.status_var.set(f"图像已保存到: {save_path}")
            messagebox.showinfo("保存成功", f"图像已保存到:\n{save_path}")
            
        except Exception as e:
            messagebox.showerror("保存失败", f"保存图像时出错: {str(e)}")
            self.status_var.set("保存失败")

    def start_crop(self, event):
        if self.operation_mode.get() in ['crop', 'both'] and self.display_image and self.crop_rect:
            # 获取相对于画布的坐标
            x = self.canvas.canvasx(event.x)
            y = self.canvas.canvasy(event.y)
            
            # 获取当前裁剪框的坐标
            x1, y1, x2, y2 = self.canvas.coords(self.crop_rect)
            
            # 检查点击是否在裁剪框内
            if x1 <= x <= x2 and y1 <= y <= y2:
                self.is_cropping = True
                # 记录点击位置和裁剪框左上角的偏移量
                self.crop_offset = (x - x1, y - y1)
                self.status_var.set("正在移动裁剪框...")

    def update_crop(self, event):
        if self.is_cropping and self.crop_rect:
            x = self.canvas.canvasx(event.x)
            y = self.canvas.canvasy(event.y)
            
            # 计算新的左上角坐标
            new_x1 = x - self.crop_offset[0]
            new_y1 = y - self.crop_offset[1]
            
            # 获取当前裁剪框尺寸
            old_x1, old_y1, old_x2, old_y2 = self.canvas.coords(self.crop_rect)
            width = old_x2 - old_x1
            height = old_y2 - old_y1
            
            # 计算新的右下角坐标
            new_x2 = new_x1 + width
            new_y2 = new_y1 + height
            
            # 获取画布可滚动区域，确保允许裁剪框在整个画布区域移动
            scroll_region = self.canvas.bbox("all")
            
            # 更新裁剪框位置，不限制必须在图片内
            self.canvas.coords(
                self.crop_rect,
                new_x1, new_y1, new_x2, new_y2
            )
            
            # 更新裁剪起始点
            self.crop_start = (new_x1, new_y1)
            
            # 更新坐标显示
            self.update_crop_coords_display([new_x1, new_y1, new_x2, new_y2])

    def end_crop(self, event):
        if self.is_cropping:
            self.is_cropping = False
            self.status_var.set("裁剪框位置已更新")

    def mode_changed(self):
        """当操作模式改变时调用"""
        mode = self.operation_mode.get()
        if mode in ["crop", "both"]:
            self.show_crop_button.grid()  # 显示裁剪框按钮
            self.status_var.set(f"已切换到{'裁剪' if mode == 'crop' else '缩放+裁剪'}模式")
            
            # 如果切换到缩放+裁剪模式，给出提示
            if mode == "both":
                messagebox.showinfo("提示", "在此模式下，您可以:\n1. 使用鼠标滚轮进行缩放\n2. 点击'显示裁剪框'创建裁剪框\n3. 拖动裁剪框选择区域")
        else:
            self.show_crop_button.grid_remove()  # 隐藏裁剪框按钮
            if self.crop_rect:
                self.canvas.delete(self.crop_rect)
                self.crop_rect = None
            # 清除坐标显示
            self.update_crop_coords_display(None)
            self.status_var.set("已切换到缩放模式")

    def show_crop_box(self):
        """显示裁剪框"""
        if not self.display_image:
            messagebox.showwarning("警告", "请先选择一张图片")
            return
        
        try:
            # 获取用户设置的输出尺寸
            target_width = self.width.get()
            target_height = self.height.get()
            
            if target_width <= 0 or target_height <= 0:
                messagebox.showerror("错误", "宽度和高度必须为正数")
                return
            
            # 获取画布尺寸
            canvas_width = self.canvas.winfo_width()
            canvas_height = self.canvas.winfo_height()
            
            # 计算裁剪框坐标，使其居中于画布
            x1 = (canvas_width - target_width) / 2
            y1 = (canvas_height - target_height) / 2
            x2 = x1 + target_width
            y2 = y1 + target_height
            
            # 删除已有的裁剪框
            if self.crop_rect:
                self.canvas.delete(self.crop_rect)
            
            # 创建新的裁剪框
            self.crop_rect = self.canvas.create_rectangle(
                x1, y1, x2, y2,
                outline="red", width=2
            )
            
            # 保存裁剪起始点，方便后续拖动
            self.crop_start = (x1, y1)
            
            # 更新坐标显示
            self.update_crop_coords_display([x1, y1, x2, y2])
            
            # 更新状态
            self.status_var.set(f"裁剪框已创建: {target_width}x{target_height}")
            
        except Exception as e:
            messagebox.showerror("错误", f"创建裁剪框时出错: {str(e)}")
            self.status_var.set("创建裁剪框失败")

    def update_crop_coords_display(self, coords):
        """更新裁剪框坐标显示"""
        if coords and len(coords) == 4:
            x1, y1, x2, y2 = map(int, coords)
            self.crop_coords_label.config(text=f"裁剪框坐标: ({x1},{y1}) - ({x2},{y2})")
        else:
            self.crop_coords_label.config(text="裁剪框坐标: -")

    def reset_image(self):
        """复原图片到原始状态"""
        if not self.original_image:
            messagebox.showwarning("警告", "没有可复原的图片")
            return
        
        try:
            # 重置图像变换状态
            self.rotation_angle = 0
            self.is_flipped_h = False
            self.is_flipped_v = False
            
            # 重置色彩调整
            self.brightness_value.set(1.0)
            self.contrast_value.set(1.0)
            self.saturation_value.set(1.0)
            
            # 重置裁剪预设
            self.crop_preset.set("自定义")
            
            # 重置显示图像为原始图像的一个副本
            self.display_image = self.original_image.copy()
            
            # 重置缩放比例
            self.zoom_scale = 1.0
            self.zoom_info_label.config(text=f"缩放比例: {self.zoom_scale:.1f}x")
            
            # 如果有裁剪框，删除它
            if self.crop_rect:
                self.canvas.delete(self.crop_rect)
                self.crop_rect = None
            
            # 清除坐标显示
            self.update_crop_coords_display(None)
            
            # 更新预览
            self.update_preview()
            
            # 更新输入框中的图片尺寸为原始尺寸
            self.width.set(self.original_width)
            self.height.set(self.original_height)
            
            # 更新状态
            self.status_var.set("图片已复原到原始状态")
        except Exception as e:
            messagebox.showerror("错误", f"复原图片时出错: {str(e)}")
            self.status_var.set("复原失败")

    def apply_changes(self):
        """应用当前更改，使其成为新的基础状态"""
        if not self.display_image:
            messagebox.showwarning("警告", "请先预览修改效果")
            return
        
        try:
            # 确认用户的操作
            result = messagebox.askyesno("确认应用", "应用当前更改？这将作为新的基础状态，无法撤销到初始状态。")
            if not result:
                return
            
            # 设置当前图像为新的原始图像
            self.original_image = self.display_image.copy()
            self.original_width, self.original_height = self.original_image.size
            
            # 重置缩放比例
            self.zoom_scale = 1.0
            self.zoom_info_label.config(text=f"缩放比例: {self.zoom_scale:.1f}x")
            
            # 更新尺寸输入框
            self.width.set(self.original_width)
            self.height.set(self.original_height)
            
            # 清除裁剪框
            if self.crop_rect:
                self.canvas.delete(self.crop_rect)
                self.crop_rect = None
                self.update_crop_coords_display(None)
            
            # 更新显示
            self.display_image = self.original_image.copy()
            self.update_preview()
            
            # 更新状态
            self.status_var.set("已应用更改，当前状态设为新的基础状态")
            
        except Exception as e:
            messagebox.showerror("错误", f"应用更改时出错: {str(e)}")
            self.status_var.set("应用更改失败")

    def center_image_in_canvas(self):
        if self.display_image and self.image_on_canvas:
            # 获取画布尺寸
            canvas_width = self.canvas.winfo_width()
            canvas_height = self.canvas.winfo_height()
            
            # 获取图像尺寸
            image_width, image_height = self.display_image.size
            
            # 计算画布中心点
            canvas_center_x = canvas_width / 2
            canvas_center_y = canvas_height / 2
            
            # 计算图像中心点
            image_center_x = image_width / 2
            image_center_y = image_height / 2
            
            # 计算滚动条位置，使图像中心与画布中心对齐
            x_scroll = max(0, (image_center_x - canvas_center_x)) / max(1, image_width)
            y_scroll = max(0, (image_center_y - canvas_center_y)) / max(1, image_height)
            
            # 设置滚动条位置
            self.canvas.xview_moveto(x_scroll)
            self.canvas.yview_moveto(y_scroll)

    def on_window_resize(self, event):
        # 当窗口大小调整时，重新计算裁剪框和图像的居中位置
        self.center_image_in_canvas()

    def setup_drag_drop(self):
        """设置拖放功能"""
        try:
            # 为画布设置拖放目标
            self.canvas.drop_target_register(DND_FILES)
            self.canvas.dnd_bind('<<Drop>>', self.on_drop)
            
            # 添加拖放提示文本
            self.drag_prompt = self.canvas.create_text(
                self.canvas.winfo_width() // 2,
                self.canvas.winfo_height() // 2,
                text="拖放图片文件到此处",
                font=('Arial', 16),
                fill='gray'
            )
            
            # 为根窗口也设置拖放
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind('<<Drop>>', self.on_drop)
            
            # 创建拖放视觉反馈
            self.canvas.bind("<DragEnter>", self.on_drag_enter)
            self.canvas.bind("<DragLeave>", self.on_drag_leave)
        except Exception as e:
            # 如果拖放设置失败，记录日志但不中断程序
            print(f"设置拖放功能失败: {e}")

    def on_drop(self, event):
        """处理拖放图片事件"""
        # 获取拖放的文件路径
        file_path = event.data
        
        # 清理文件路径（去除{}，处理不同操作系统格式差异）
        if file_path.startswith('{') and file_path.endswith('}'):
            file_path = file_path[1:-1]
        
        # 处理Windows路径中的引号
        if file_path.startswith('"') and file_path.endswith('"'):
            file_path = file_path[1:-1]
        
        # 处理多文件拖放情况（我们只取第一个）
        if " " in file_path and os.path.exists(file_path.split(" ")[0]):
            file_path = file_path.split(" ")[0]
        
        # 验证文件类型
        valid_extensions = ['.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff']
        if any(file_path.lower().endswith(ext) for ext in valid_extensions):
            # 设置源文件路径并加载图片
            self.source_path.set(file_path)
            self.load_image()
            
            # 处理文件名和目标路径
            base_name = os.path.basename(file_path)
            name_without_ext = os.path.splitext(base_name)[0]
            self.new_filename.set(name_without_ext)
            
            # 如果未设置目标路径，使用源文件目录
            if not self.target_path.get():
                self.target_path.set(os.path.dirname(file_path))
            
            # 如果有提示文本，删除它
            if self.drag_prompt:
                self.canvas.delete(self.drag_prompt)
                self.drag_prompt = None
        else:
            messagebox.showerror("格式错误", "请拖放有效的图片文件")

    def on_drag_enter(self, event):
        """鼠标拖着文件进入区域时的视觉反馈"""
        self.canvas.config(bg="#e0f0ff")  # 改变背景色提示可以放置

    def on_drag_leave(self, event):
        """鼠标拖着文件离开区域时恢复正常"""
        self.canvas.config(bg="white")

    def rotate_image(self, angle):
        """旋转图像"""
        if not self.original_image:
            messagebox.showwarning("警告", "请先选择一张图片")
            return
        
        try:
            # 更新旋转角度
            self.rotation_angle = (self.rotation_angle + angle) % 360
            
            # 应用旋转到原始图像
            rotated_image = self.original_image.rotate(self.rotation_angle, expand=True, resample=Image.BICUBIC)
            
            # 如果有应用翻转，也要应用
            if self.is_flipped_h:
                rotated_image = ImageOps.mirror(rotated_image)
            if self.is_flipped_v:
                rotated_image = ImageOps.flip(rotated_image)
            
            # 应用色彩调整
            processed_image = self.apply_color_adjustments(rotated_image)
            
            # 更新显示图像
            self.original_width, self.original_height = processed_image.size
            
            # 调整显示图像大小
            self.display_image = processed_image.resize(
                (int(self.original_width * self.zoom_scale), 
                 int(self.original_height * self.zoom_scale)), 
                Image.LANCZOS
            )
            
            # 更新尺寸输入框
            self.width.set(self.original_width)
            self.height.set(self.original_height)
            
            # 更新预览
            self.update_preview()
            
            # 更新状态
            self.status_var.set(f"图像已旋转 {angle}°，当前旋转角度: {self.rotation_angle}°")
        
        except Exception as e:
            messagebox.showerror("错误", f"旋转图像时出错: {str(e)}")
            self.status_var.set("旋转失败")

    def flip_horizontal(self):
        """水平翻转图像"""
        if not self.original_image:
            messagebox.showwarning("警告", "请先选择一张图片")
            return
        
        try:
            # 切换水平翻转标志
            self.is_flipped_h = not self.is_flipped_h
            
            # 应用变换到原始图像
            flipped_image = self.original_image.rotate(self.rotation_angle, expand=True, resample=Image.BICUBIC)
            
            # 应用翻转
            if self.is_flipped_h:
                flipped_image = ImageOps.mirror(flipped_image)
            if self.is_flipped_v:
                flipped_image = ImageOps.flip(flipped_image)
            
            # 应用色彩调整
            processed_image = self.apply_color_adjustments(flipped_image)
            
            # 更新显示图像
            self.original_width, self.original_height = processed_image.size
            
            # 调整显示图像大小
            self.display_image = processed_image.resize(
                (int(self.original_width * self.zoom_scale), 
                 int(self.original_height * self.zoom_scale)), 
                Image.LANCZOS
            )
            
            # 更新预览
            self.update_preview()
            
            # 更新状态
            self.status_var.set(f"图像已{'应用' if self.is_flipped_h else '取消'}水平翻转")
        
        except Exception as e:
            messagebox.showerror("错误", f"翻转图像时出错: {str(e)}")
            self.status_var.set("翻转失败")

    def flip_vertical(self):
        """垂直翻转图像"""
        if not self.original_image:
            messagebox.showwarning("警告", "请先选择一张图片")
            return
        
        try:
            # 切换垂直翻转标志
            self.is_flipped_v = not self.is_flipped_v
            
            # 应用变换到原始图像
            flipped_image = self.original_image.rotate(self.rotation_angle, expand=True, resample=Image.BICUBIC)
            
            # 应用翻转
            if self.is_flipped_h:
                flipped_image = ImageOps.mirror(flipped_image)
            if self.is_flipped_v:
                flipped_image = ImageOps.flip(flipped_image)
            
            # 应用色彩调整
            processed_image = self.apply_color_adjustments(flipped_image)
            
            # 更新显示图像
            self.original_width, self.original_height = processed_image.size
            
            # 调整显示图像大小
            self.display_image = processed_image.resize(
                (int(self.original_width * self.zoom_scale), 
                 int(self.original_height * self.zoom_scale)), 
                Image.LANCZOS
            )
            
            # 更新预览
            self.update_preview()
            
            # 更新状态
            self.status_var.set(f"图像已{'应用' if self.is_flipped_v else '取消'}垂直翻转")
        
        except Exception as e:
            messagebox.showerror("错误", f"翻转图像时出错: {str(e)}")
            self.status_var.set("翻转失败")

    def update_color_adjustments(self, *args):
        """当色彩调整滑块改变时更新图像"""
        if not self.original_image:
            return
        
        # 设置重新渲染标志
        self.need_rerender = True
        
        # 使用定时器延迟处理，避免频繁更新
        self.root.after(200, self.do_color_update)

    def do_color_update(self):
        """实际执行色彩更新的方法"""
        if not self.need_rerender:
            return
        
        self.need_rerender = False
        
        try:
            # 从原始图像开始进行所有变换
            transformed_image = self.original_image.rotate(self.rotation_angle, expand=True, resample=Image.BICUBIC)
            
            # 应用翻转
            if self.is_flipped_h:
                transformed_image = ImageOps.mirror(transformed_image)
            if self.is_flipped_v:
                transformed_image = ImageOps.flip(transformed_image)
            
            # 应用色彩调整
            processed_image = self.apply_color_adjustments(transformed_image)
            
            # 更新显示图像
            self.original_width, self.original_height = processed_image.size
            
            # 调整显示图像大小
            self.display_image = processed_image.resize(
                (int(self.original_width * self.zoom_scale), 
                 int(self.original_height * self.zoom_scale)), 
                Image.LANCZOS
            )
            
            # 更新预览
            self.update_preview()
            
            # 更新状态
            self.status_var.set(f"色彩调整已应用 (亮度:{self.brightness_value.get():.2f}, 对比度:{self.contrast_value.get():.2f}, 饱和度:{self.saturation_value.get():.2f})")
        
        except Exception as e:
            messagebox.showerror("错误", f"应用色彩调整时出错: {str(e)}")
            self.status_var.set("色彩调整失败")

    def apply_color_adjustments(self, image):
        """应用色彩调整到给定图像"""
        try:
            brightness = self.brightness_value.get()
            contrast = self.contrast_value.get()
            saturation = self.saturation_value.get()
            
            # 应用亮度
            if brightness != 1.0:
                enhancer = ImageEnhance.Brightness(image)
                image = enhancer.enhance(brightness)
            
            # 应用对比度
            if contrast != 1.0:
                enhancer = ImageEnhance.Contrast(image)
                image = enhancer.enhance(contrast)
            
            # 应用饱和度
            if saturation != 1.0:
                enhancer = ImageEnhance.Color(image)
                image = enhancer.enhance(saturation)
            
            return image
        
        except Exception as e:
            print(f"应用色彩调整时出错: {str(e)}")
            return image  # 返回原始图像

    def reset_color_adjustments(self):
        """重置所有色彩调整为默认值"""
        self.brightness_value.set(1.0)
        self.contrast_value.set(1.0)
        self.saturation_value.set(1.0)
        
        self.update_color_adjustments()
        
        self.status_var.set("色彩调整已重置")

    def apply_crop_preset(self, event=None):
        """应用选定的裁剪预设"""
        if not self.original_image:
            messagebox.showwarning("警告", "请先选择一张图片")
            self.crop_preset.set("自定义")
            return
        
        preset = self.crop_preset.get()
        
        # 定义各种预设的宽高比
        presets = {
            "自定义": None,  # 不做特殊处理
            "正方形 (1:1)": (1, 1),
            "Instagram (4:5)": (4, 5),
            "Facebook (16:9)": (16, 9),
            "Twitter (2:1)": (2, 1),
            "LinkedIn (1.91:1)": (1.91, 1),
            "微信 (4:3)": (4, 3)
        }
        
        if preset not in presets or presets[preset] is None:
            return
        
        try:
            # 获取预设的宽高比
            ratio_w, ratio_h = presets[preset]
            
            # 获取原始图像的宽高
            orig_w, orig_h = self.original_width, self.original_height
            
            # 计算最适合的裁剪尺寸
            # 先尝试用原始宽度来计算
            crop_w = orig_w
            crop_h = int(crop_w * ratio_h / ratio_w)
            
            # 如果计算出的高度超出原图，就用原始高度来计算
            if crop_h > orig_h:
                crop_h = orig_h
                crop_w = int(crop_h * ratio_w / ratio_h)
            
            # 更新宽高输入框
            self.width.set(crop_w)
            self.height.set(crop_h)
            
            # 如果在裁剪或缩放+裁剪模式下，显示预设尺寸的裁剪框
            if self.operation_mode.get() in ['crop', 'both']:
                self.show_crop_box()
            
            self.status_var.set(f"已应用裁剪预设: {preset} ({crop_w}x{crop_h})")
        
        except Exception as e:
            messagebox.showerror("错误", f"应用裁剪预设时出错: {str(e)}")
            self.status_var.set("应用预设失败")

    def export_as_icon(self, format_type):
        """将当前图像导出为指定格式的图标
        
        Args:
            format_type: 'ico', 'icns' 或 'png_set'
        """
        if not self.original_image:
            messagebox.showwarning("警告", "请先选择一张图片")
            return
        
        try:
            # 获取要保存的文件路径
            file_types = []
            default_ext = ''
            if format_type == 'ico':
                file_types = [("ICO图标", "*.ico")]
                default_ext = '.ico'
                title = "保存为ICO图标"
            elif format_type == 'icns':
                file_types = [("ICNS图标", "*.icns")]
                default_ext = '.icns'
                title = "保存为ICNS图标"
            elif format_type == 'png_set':
                file_types = [("PNG图标集", "*.png")]
                default_ext = '.png'
                title = "保存为PNG图标集"
            
            # 默认文件名
            default_name = self.new_filename.get() or "icon"
            if not default_name.endswith(default_ext):
                default_name += default_ext
            
            # 获取保存路径
            save_path = filedialog.asksaveasfilename(
                title=title,
                initialdir=self.target_path.get(),
                initialfile=default_name,
                defaultextension=default_ext,
                filetypes=file_types
            )
            
            if not save_path:
                return  # 用户取消操作
            
            # 更新状态
            self.status_var.set(f"正在创建{format_type.upper()}图标...")
            self.root.update()
            
            # 处理图像转换
            # 先应用所有编辑
            transformed_image = self.original_image.rotate(self.rotation_angle, expand=True, resample=Image.BICUBIC)
            if self.is_flipped_h:
                transformed_image = ImageOps.mirror(transformed_image)
            if self.is_flipped_v:
                transformed_image = ImageOps.flip(transformed_image)
            processed_image = self.apply_color_adjustments(transformed_image)
            
            # 导出对应格式
            if format_type == 'ico':
                # 创建ICO文件
                IconConverter.create_ico(processed_image, save_path)
                self.status_var.set(f"已成功导出ICO图标: {save_path}")
                
            elif format_type == 'icns':
                # 创建ICNS文件
                IconConverter.create_icns(processed_image, save_path)
                self.status_var.set(f"已成功导出ICNS图标: {save_path}")
                
            elif format_type == 'png_set':
                # 创建PNG图标集
                self.export_png_icon_set(processed_image, save_path)
            
            # 弹出成功消息
            messagebox.showinfo("导出成功", f"图标已成功保存到:\n{save_path}")
            
        except Exception as e:
            messagebox.showerror("错误", f"导出图标时出错: {str(e)}")
            self.status_var.set(f"导出{format_type.upper()}图标失败")

    def export_png_icon_set(self, image, base_path):
        """导出一组不同尺寸的PNG图标
        
        Args:
            image: PIL图像对象
            base_path: 基本文件路径，将自动添加尺寸后缀
        """
        # 移除扩展名
        base_path = os.path.splitext(base_path)[0]
        base_dir = os.path.dirname(base_path)
        base_name = os.path.basename(base_path)
        
        # 创建图标集目录
        icons_dir = os.path.join(base_dir, f"{base_name}_icons")
        os.makedirs(icons_dir, exist_ok=True)
        
        # 确保图像是正方形
        width, height = image.size
        if width != height:
            # 取最小的边作为裁剪尺寸
            size = min(width, height)
            # 计算裁剪区域，使其居中
            left = (width - size) // 2
            top = (height - size) // 2
            right = left + size
            bottom = top + size
            image = image.crop((left, top, right, bottom))
        
        # 创建不同尺寸的图标
        sizes = [16, 32, 48, 64, 128, 256, 512, 1024]
        for size in sizes:
            resized = image.resize((size, size), Image.LANCZOS)
            icon_path = os.path.join(icons_dir, f"{base_name}_{size}x{size}.png")
            resized.save(icon_path, 'PNG')
        
        self.status_var.set(f"已成功导出PNG图标集到: {icons_dir}")
        return icons_dir

# 添加图标转换类
class IconConverter:
    """用于转换图像为各种图标格式的工具类"""
    
    @staticmethod
    def create_ico(image, output_path, sizes=None):
        """将PIL图像转换为.ico格式
        
        Args:
            image: PIL图像对象
            output_path: 输出的.ico文件路径
            sizes: 要包含的尺寸列表，默认为[16, 32, 48, 64, 128, 256]
        """
        if sizes is None:
            sizes = [16, 32, 48, 64, 128, 256]
        
        # 确保图像是正方形，否则进行裁剪
        width, height = image.size
        if width != height:
            # 取最小的边作为裁剪尺寸
            size = min(width, height)
            # 计算裁剪区域，使其居中
            left = (width - size) // 2
            top = (height - size) // 2
            right = left + size
            bottom = top + size
            image = image.crop((left, top, right, bottom))
        
        # 创建不同尺寸的图像
        icons = []
        for size in sizes:
            # 调整图像大小，保持纵横比
            resized_img = image.resize((size, size), Image.LANCZOS)
            icons.append(resized_img)
        
        # 保存为.ico文件
        icons[0].save(
            output_path, 
            format='ICO', 
            sizes=[(img.width, img.height) for img in icons],
            append_images=icons[1:]
        )
        return output_path
    
    @staticmethod
    def create_icns(image, output_path):
        """将PIL图像转换为.icns格式
        
        Args:
            image: PIL图像对象
            output_path: 输出的.icns文件路径
        """
        # 确保输出路径以.icns结尾
        if not output_path.lower().endswith('.icns'):
            output_path += '.icns'
            
        # 创建临时目录存放图标集
        with tempfile.TemporaryDirectory() as iconset_dir:
            iconset_path = os.path.join(iconset_dir, 'icon.iconset')
            os.makedirs(iconset_path, exist_ok=True)
            
            # 确保图像是正方形
            width, height = image.size
            if width != height:
                # 取最小的边作为裁剪尺寸
                size = min(width, height)
                # 计算裁剪区域，使其居中
                left = (width - size) // 2
                top = (height - size) // 2
                right = left + size
                bottom = top + size
                image = image.crop((left, top, right, bottom))
                
            # 创建所需的各种尺寸图标
            icon_sizes = [16, 32, 128, 256, 512]
            retina_sizes = [32, 64, 256, 512, 1024]
            
            # 生成正常尺寸图标
            for size in icon_sizes:
                resized = image.resize((size, size), Image.LANCZOS)
                icon_path = os.path.join(iconset_path, f'icon_{size}x{size}.png')
                resized.save(icon_path, 'PNG')
            
            # 生成Retina尺寸图标（2x分辨率）
            for i, size in enumerate(icon_sizes):
                retina_size = retina_sizes[i]
                resized = image.resize((retina_size, retina_size), Image.LANCZOS)
                icon_path = os.path.join(iconset_path, f'icon_{size}x{size}@2x.png')
                resized.save(icon_path, 'PNG')
            
            # 尝试使用iconutil（macOS）转换为icns
            try:
                if sys.platform == 'darwin':  # macOS系统
                    subprocess.run(['iconutil', '-c', 'icns', iconset_path, '-o', output_path], 
                                   check=True)
                    return output_path
            except (subprocess.SubprocessError, FileNotFoundError):
                pass
            
            # 如果iconutil失败或不是macOS，尝试使用PIL自行生成icns
            try:
                # 生成最大尺寸的PNG
                max_size = 1024
                max_image = image.resize((max_size, max_size), Image.LANCZOS)
                
                # 在临时目录中创建一个PNG
                png_path = os.path.join(iconset_dir, 'temp_icon.png')
                max_image.save(png_path, 'PNG')
                
                # 读取PNG数据
                with open(png_path, 'rb') as f:
                    png_data = f.read()
                
                # 创建简单的ICNS文件结构
                # ICNS格式有点复杂，这里是简化实现
                icns_data = b'icns' + len(png_data).to_bytes(4, byteorder='big') + b'ic10' + png_data
                
                # 写入ICNS文件
                with open(output_path, 'wb') as f:
                    f.write(icns_data)
                
                return output_path
            except Exception as e:
                raise Exception(f"无法创建ICNS文件: {str(e)}")

def main():
    # 使用TkinterDnD替代标准的Tk
    root = TkinterDnD.Tk()
    
    # 设置应用程序图标（如果有）
    try:
        if sys.platform.startswith('win'):
            root.iconbitmap('app_icon.ico')  # Windows平台
        else:
            logo = tk.PhotoImage(file='app_icon.png')  # Linux/Mac平台
            root.tk.call('wm', 'iconphoto', root._w, logo)
    except:
        pass  # 如果没有图标文件，忽略异常
    
    # 初始化应用
    app = ImageTrimmerApp(root)
    
    # 设置最小窗口大小 - 增加最小高度确保所有控件可见
    root.minsize(960, 960)
    
    # 启动主循环
    root.mainloop()

if __name__ == "__main__":
    main()