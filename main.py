import tkinter as tk
from tkinter import ttk, colorchooser, filedialog, messagebox
import ctypes
from ctypes import wintypes
import json
import os
import sys
import base64
from PIL import Image, ImageTk
import pystray
from pystray import MenuItem as item
import threading
import winreg
import subprocess
import keyboard
# import mouse
import win32gui
import win32con
from pynput import mouse as pynput_mouse
import time
import zlib
import tkinter.simpledialog as simpledialog

import tkinter.scrolledtext as scrolledtext
import datetime

# 解决高DPI下的缩放问题，防止DWM合成时的帧率不匹配
try:
    # 0: DPI_AWARENESS_INVALID
    # 1: DPI_AWARENESS_SYSTEM_AWARE (整个系统统一缩放)
    # 2: DPI_AWARENESS_PER_MONITOR_AWARE (每个显示器独立缩放)
    # 使用 1 可能会导致界面在高 DPI 屏幕上看起来较小，但兼容性最好，能解决 DWM 合成问题
    # 使用 0 则是由系统接管缩放，界面会模糊但大小正常
    # 为了解决锁帧问题我们启用了 DPI Awareness，但这会导致界面不再被系统自动放大
    
    # 尝试使用 ctypes.windll.shcore.SetProcessDpiAwareness(1)
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

# 自动适配屏幕缩放比例
def get_dpi_scaling():
    try:
        root = tk.Tk()
        scaling = root.tk.call('tk', 'scaling')
        root.destroy()
        return scaling
    except:
        return 1.33 # Default fallback

# Windows API constants for click-through
GWL_EXSTYLE = -20
WS_EX_LAYERED = 0x00080000
WS_EX_TRANSPARENT = 0x00000020
WS_EX_TOOLWINDOW = 0x00000080 # 不在Alt-Tab和任务栏显示
WS_EX_NOACTIVATE = 0x08000000 # 不激活窗口

class CrosshairOverlay(tk.Toplevel):
    def __init__(self, master, config):
        super().__init__(master)
        self.config = config
        self.title("Overlay")
        self._image_job = None
        self.last_center_x = None
        self.last_center_y = None
        
        # Remove decorations
        self.overrideredirect(True)
        self.wm_attributes("-topmost", True)
        
        # Transparency
        self.bg_color = "#000001"
        self.config_bg(self.bg_color)
        self.wm_attributes("-transparentcolor", self.bg_color)
        
        # Dimensions (fixed small area around center to minimize impact, but large enough for big crosshairs)
        self.width = 200
        self.height = 200
        
        # Canvas
        self.canvas = tk.Canvas(self, width=self.width, height=self.height, 
                                bg=self.bg_color, highlightthickness=0)
        self.canvas.pack()
        
        # Initial Draw
        self.redraw()
        
        # Apply click-through
        self.after(100, self.apply_click_through)
        
        self.image_ref = None
        
        # Start keep-on-top loop
        self.keep_on_top()

    def ensure_canvas_size(self, w, h):
        w = max(10, int(w))
        h = max(10, int(h))
        if self.width != w or self.height != h:
            self.width = w
            self.height = h
            self.canvas.config(width=self.width, height=self.height)
            if self.last_center_x is not None and self.last_center_y is not None:
                self.set_position(self.last_center_x, self.last_center_y)

    def config_bg(self, color):
        self.configure(bg=color)

    def keep_on_top(self):
        """Periodically enforce topmost status and styles"""
        try:
            if not self.winfo_viewable():
                self.after(5000, self.keep_on_top)
                return

            hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
            if hwnd == 0:
                hwnd = self.winfo_id()
            
            if hwnd:
                # 1. Check TopMost
                # GetWindowLong only returns styles, check extended style for WS_EX_TOPMOST (0x8)
                current_ex_style = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
                is_topmost = (current_ex_style & 0x8) != 0
                
                if not is_topmost:
                    # HWND_TOPMOST = -1, SWP_NOMOVE|SWP_NOSIZE|SWP_NOACTIVATE
                    ctypes.windll.user32.SetWindowPos(hwnd, -1, 0, 0, 0, 0, 0x0013)
                
                # 2. Enforce Styles (Transparent, Layered, ToolWindow, NoActivate)
                required = WS_EX_TRANSPARENT | WS_EX_LAYERED | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE
                
                if (current_ex_style & required) != required:
                    ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, current_ex_style | required)
                    
        except Exception:
            pass
        
        # Check every 5 seconds to minimize overhead
        self.after(5000, self.keep_on_top)

    def apply_click_through(self):
        try:
            hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
            if hwnd == 0:
                hwnd = self.winfo_id()
            
            # Get current style
            current_style = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
            
            # Add styles:
            # WS_EX_TRANSPARENT: Click-through
            # WS_EX_LAYERED: Alpha blending
            # WS_EX_TOOLWINDOW: Hide from Alt-Tab/Taskbar (helps with overlay behavior)
            # WS_EX_NOACTIVATE: Prevent window activation (helps with focus stealing)
            new_style = current_style | WS_EX_TRANSPARENT | WS_EX_LAYERED | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE
            
            if new_style != current_style:
                ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, new_style)
                # Force a redraw to apply style changes correctly if needed
                # ctypes.windll.user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0, 0x0027) 
        except Exception as e:
            print(f"Error setting click-through: {e}")

    def redraw(self):
        self.canvas.delete("all")
        size = self.config['size'].get()
        color = self.config['color'].get()
        thickness = self.config['thickness'].get()
        dot = self.config['dot'].get()
        style = self.config['style'].get()
        req_w = self.width
        req_h = self.height
        # Defer heavy custom image loading to avoid startup contention
        if style == "Custom" or style in ["自定义图片", "自定义"]:
            if self._image_job:
                try:
                    self.after_cancel(self._image_job)
                except Exception:
                    pass
                self._image_job = None
            self._image_job = self.after(200, self._draw_custom_image)
            return
        # Vector styles sizing
        base = max(size + thickness * 2, dot + thickness * 2)
        base = max(base, thickness * 2 + 2)
        req_w = req_h = base + 20
        self.ensure_canvas_size(req_w, req_h)
        cx, cy = self.width // 2, self.height // 2
        
        if style == "Cross" or style == "Both" or style == "十字" or style == "混合":
            # Horizontal
            self.canvas.create_line(cx - size//2, cy, cx + size//2, cy, 
                                    fill=color, width=thickness, tags="crosshair")
            # Vertical
            self.canvas.create_line(cx, cy - size//2, cx, cy + size//2, 
                                    fill=color, width=thickness, tags="crosshair")
            
        if style == "Dot" or style == "Both" or style == "圆点" or style == "混合":
            r = dot // 2
            self.canvas.create_oval(cx - r, cy - r, cx + r, cy + r, 
                                    fill=color, outline=color, tags="crosshair")
        
        if style == "Circle" or style == "圆圈":
            r = size // 2
            self.canvas.create_oval(cx - r, cy - r, cx + r, cy + r,
                                    outline=color, width=thickness, tags="crosshair")
                                    
        # Note: Custom image is handled in the fast-path above

    def _draw_custom_image(self):
        try:
            image_path = self.config.get('image_path', {}).get()
            if not (image_path and os.path.exists(image_path)):
                return
            with Image.open(image_path) as im:
                im = im.convert("RGBA")
                try:
                    scale = int(self.config.get('img_scale', {}).get())
                except Exception:
                    scale = 100
                # Crop transparent areas to reduce window size and DWM load
                bbox = im.getbbox()
                if bbox:
                    im = im.crop(bbox)

                scale = max(5, min(200, scale))
                if scale != 100:
                    new_w = max(1, int(im.width * scale / 100))
                    new_h = max(1, int(im.height * scale / 100))
                    im = im.resize((new_w, new_h), Image.LANCZOS)
                
                self.image_ref = ImageTk.PhotoImage(im)
                req_w, req_h = self.image_ref.width(), self.image_ref.height()
                max_w = max(1, self.winfo_screenwidth() * 2)
                max_h = max(1, self.winfo_screenheight() * 2)
                if req_w <= 0 or req_h <= 0 or req_w > 20000 or req_h > 20000:
                    return
                if req_w > max_w or req_h > max_h:
                    return
                self.canvas.delete("all")
                self.ensure_canvas_size(req_w, req_h)
                cx, cy = self.width // 2, self.height // 2
                self.canvas.create_image(cx, cy, image=self.image_ref, anchor="center", tags="crosshair")
        except Exception as e:
            print(f"Error loading image: {e}")
        finally:
            self._image_job = None

    def set_position(self, x, y):
        # x, y are center coordinates
        # We need to convert to top-left for geometry
        self.last_center_x = x
        self.last_center_y = y
        tl_x = x - self.width // 2
        tl_y = y - self.height // 2
        self.geometry(f"{self.width}x{self.height}+{tl_x}+{tl_y}")

import webbrowser

class ControlPanel:
    def __init__(self):
        self.root = tk.Tk()
        self.version = "v1.0"

        
        # 针对高DPI的字体缩放补偿
        # Tkinter 默认在 DPI Aware 模式下不会自动放大字体，导致界面很小
        # 需要手动根据 DPI 设置 scaling
        try:
            # 获取系统 DPI (96 is default)
            dpi = ctypes.windll.user32.GetDpiForSystem()
            scale_factor = dpi / 96.0
            self.root.tk.call('tk', 'scaling', scale_factor * 1.33) 
        except:
            pass

        # Check Admin Status
        self.is_admin = ctypes.windll.shell32.IsUserAnAdmin()
        title = "自定义准星"
        if self.is_admin:
            title += " - 管理员模式"
        else:
            title += " - 非管理员模式"
            
        self.root.title(title)
        
        try:
            self.root.iconbitmap(self.resource_path("tx.ico"))
        except:
            pass
            
        # 根据DPI动态调整窗口大小
        base_w, base_h = 360, 560
        try:
             dpi = ctypes.windll.user32.GetDpiForSystem()
             scale = dpi / 96.0
             if scale > 1.0:
                 base_w = int(base_w * scale)
                 base_h = int(base_h * scale)
        except:
             pass
             
        self.root.geometry(f"{base_w}x{base_h}")
        self.root.resizable(False, False)
        
        self.overlay = None
        
        # Configuration Variables
        self.screen_w = self.root.winfo_screenwidth()
        self.screen_h = self.root.winfo_screenheight()
        
        self.pos_x = tk.StringVar(value=str(self.screen_w // 2))
        self.pos_y = tk.StringVar(value=str(self.screen_h // 2))
        
        # Auto update on change
        self.pos_x.trace_add("write", self.update_pos)
        self.pos_y.trace_add("write", self.update_pos)
        
        self.config = {
            'size': tk.IntVar(value=20),
            'thickness': tk.IntVar(value=2),
            'color': tk.StringVar(value="#00FF00"),
            'dot': tk.IntVar(value=4),
            'img_scale': tk.IntVar(value=100),
            'style': tk.StringVar(value="十字"),
            'image_path': tk.StringVar(value=""),
            'force_admin': tk.BooleanVar(value=False),
            'hide_hotkey': tk.StringVar(value=""),
            'trigger_type': tk.StringVar(value="keyboard"),
            'trigger_mode': tk.StringVar(value="点击切换")
        }
        
        self.presets = {}
        self.current_preset_name = tk.StringVar()
        self.PRESET_PLACEHOLDER = "<--下拉选择准星预设-->"
        self.crosshair_visible = True
        
        self.mouse_listener = None
        self.keyboard_hooks = []
        
        # State tracking for hold modes
        self.is_pressed = False
        
        # Lock to prevent race conditions during trigger application
        self.trigger_lock = threading.Lock()
        
        self.load_config()
        
        self.create_widgets()
        
        self.start_overlay()
        
        # Trigger style change logic to set button state and load image if needed
        # Call this AFTER starting overlay so update_overlay works
        self.on_style_change(event="Startup") 
        self.update_preset_list()
        # Ensure visible on startup regardless of trigger mode
        self.root.after(0, lambda: self.set_visible(True, force=True))
        
        # Add keyboard bindings to the Control Panel for fine tuning
        self.root.bind("<Up>", lambda e: self.adjust_pos(0, -1))
        self.root.bind("<Down>", lambda e: self.adjust_pos(0, 1))
        self.root.bind("<Left>", lambda e: self.adjust_pos(-1, 0))
        self.root.bind("<Right>", lambda e: self.adjust_pos(1, 0))
        
        self.root.protocol("WM_DELETE_WINDOW", self.quit_application)
        
        # Start tray icon in separate thread
        self.tray_icon = None
        
        self.root.mainloop()

    def on_img_scale_change(self, val):
        try:
            v = int(float(val))
        except Exception:
            v = self.config['img_scale'].get()
        v = max(5, min(200, v))
        self.img_scale_label.set(f"图像({v}%)")
        # Ensure variable stays clamped
        if self.config['img_scale'].get() != v:
            self.config['img_scale'].set(v)
        self.update_overlay()

    def check_startup(self):
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_READ)
            try:
                winreg.QueryValueEx(key, "MoliCrosshair")
                self.startup_btn.configure(text="开机自启：开")
            except FileNotFoundError:
                self.startup_btn.configure(text="开机自启：关")
            winreg.CloseKey(key)
        except Exception as e:
            print(f"Error checking startup: {e}")

    def toggle_startup(self):
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_ALL_ACCESS)
            try:
                # Try to get value to see if it exists
                winreg.QueryValueEx(key, "MoliCrosshair")
                # If exists, delete it (Turn Off)
                winreg.DeleteValue(key, "MoliCrosshair")
                self.startup_btn.configure(text="开机自启：关")
            except FileNotFoundError:
                # If not exists, create it (Turn On)
                exe_path = os.path.abspath(sys.argv[0])
                winreg.SetValueEx(key, "MoliCrosshair", 0, winreg.REG_SZ, exe_path)
                self.startup_btn.configure(text="开机自启：开")
            winreg.CloseKey(key)
        except Exception as e:
            messagebox.showerror("错误", f"无法修改开机启动设置：{e}")

    def minimize_to_tray(self):
        self.root.withdraw()
        self.create_tray_icon()

    def create_tray_icon(self):
        if self.tray_icon:
            return
            
        try:
            icon_image = Image.open(self.resource_path("tx.ico"))
        except:
            # Fallback if icon load fails
            icon_image = Image.new('RGB', (64, 64), color = (73, 109, 137))
            
        def show_window(icon, item):
            icon.stop()
            self.root.after(0, self.root.deiconify)
            self.tray_icon = None

        def quit_app(icon, item):
            icon.stop()
            self.root.after(0, self.quit_application)

        menu = (item('显示设置', show_window, default=True), item('退出程序', quit_app))
        self.tray_icon = pystray.Icon("name", icon_image, "自定义准心", menu)
        
        # Run tray icon in a separate thread to avoid blocking main loop
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def restart_as_admin(self):
        try:
            # Set force_admin to True and save config
            self.config['force_admin'].set(True)
            self.save_config()
            
            # Re-run the program with admin rights
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
            sys.exit() # Directly exit without calling quit_application again which saves config
        except Exception as e:
            messagebox.showerror("错误", f"无法以管理员身份重启：{e}")

    def restart_as_normal(self):
        try:
            # Set force_admin to False and save config
            self.config['force_admin'].set(False)
            self.save_config()
            
            # Use explorer to launch the app, which typically de-elevates to user level
            # Quote the path to handle spaces
            exe_path = f'"{sys.executable}"'
            subprocess.Popen(f'explorer {exe_path}', shell=True)
            sys.exit() # Directly exit
        except Exception as e:
            messagebox.showerror("错误", f"无法重启：{e}")

    def quit_application(self):
        self.save_config()
        self.root.quit()
        sys.exit()

    def create_widgets(self):
        self.root.columnconfigure(0, weight=1)
        
        # Presets (Moved to Position Frame)
        
        # Style
        style_frame = ttk.LabelFrame(self.root, text="样式")
        style_frame.pack(fill="x", padx=10, pady=5)
        style_frame.columnconfigure(1, weight=1)
        
        ttk.Label(style_frame, text="类型:").grid(row=0, column=0, padx=5, pady=5)
        type_cb = ttk.Combobox(style_frame, textvariable=self.config['style'], 
                               values=["十字", "圆点", "混合", "圆圈", "自定义图片"], state="readonly")
        type_cb.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        type_cb.bind("<<ComboboxSelected>>", self.on_style_change)
        
        self.img_btn = ttk.Button(style_frame, text="选择图片", command=self.choose_image)
        self.img_btn.grid(row=0, column=2, padx=5, pady=5)
        
        ttk.Label(style_frame, text="颜色:").grid(row=1, column=0, padx=5, pady=5)
        ttk.Button(style_frame, text="选择", command=self.choose_color).grid(row=1, column=1, columnspan=2, padx=5, pady=5, sticky="ew")
        
        # Size Controls
        size_frame = ttk.LabelFrame(self.root, text="尺寸")
        size_frame.pack(fill="x", padx=10, pady=5)
        size_frame.columnconfigure(1, weight=1)
        
        self.add_slider(size_frame, "大小", self.config['size'], 5, 100, 2)
        self.add_slider(size_frame, "粗细", self.config['thickness'], 1, 10, 3)
        self.add_slider(size_frame, "圆点大小", self.config['dot'], 1, 20, 4)
        # Image scale with dynamic label: 图像(n%)
        self.img_scale_label = tk.StringVar(value=f"图像({self.config['img_scale'].get()}%)")
        ttk.Label(size_frame, textvariable=self.img_scale_label).grid(row=5, column=0, padx=5, pady=2)
        img_scale_slider = ttk.Scale(size_frame, from_=5, to=200, variable=self.config['img_scale'],
                                     orient="horizontal", command=self.on_img_scale_change)
        img_scale_slider.grid(row=5, column=1, sticky="ew", padx=5, pady=2)

        # Position Controls
        pos_frame = ttk.LabelFrame(self.root, text="位置 (使用方向键微调)")
        pos_frame.pack(fill="x", padx=10, pady=5)
        pos_frame.columnconfigure(1, weight=1)
        pos_frame.columnconfigure(3, weight=1)
        
        ttk.Label(pos_frame, text="X:").grid(row=0, column=0, padx=5)
        x_entry = tk.Entry(pos_frame, textvariable=self.pos_x, width=10)
        x_entry.grid(row=0, column=1, padx=5, sticky="ew")
        
        ttk.Label(pos_frame, text="Y:").grid(row=0, column=2, padx=5)
        y_entry = tk.Entry(pos_frame, textvariable=self.pos_y, width=10)
        y_entry.grid(row=0, column=3, padx=5, sticky="ew")
        
        # Center and Drag buttons in same row
        ctrl_btn_frame = ttk.Frame(pos_frame)
        ctrl_btn_frame.grid(row=1, column=0, columnspan=4, pady=5, sticky="ew", padx=5)
        ctrl_btn_frame.columnconfigure(0, weight=1)
        ctrl_btn_frame.columnconfigure(1, weight=1)

        ttk.Button(ctrl_btn_frame, text="居中", command=self.center_pos).grid(row=0, column=0, sticky="ew", padx=2)
        
        drag_btn = ttk.Button(ctrl_btn_frame, text="按住拖动准心")
        drag_btn.grid(row=0, column=1, sticky="ew", padx=2)
        drag_btn.bind("<ButtonPress-1>", self.drag_start)
        drag_btn.bind("<B1-Motion>", self.drag_move)
        
        # Presets (Moved here)
        preset_frame = ttk.LabelFrame(pos_frame, text="预设配置")
        preset_frame.grid(row=3, column=0, columnspan=4, sticky="ew", padx=5, pady=5)
        preset_frame.columnconfigure(1, weight=1)
        
        ttk.Label(preset_frame, text="方案:").grid(row=0, column=0, padx=5, pady=5)
        self.preset_cb = ttk.Combobox(preset_frame, textvariable=self.current_preset_name)
        self.preset_cb.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.preset_cb.bind("<<ComboboxSelected>>", self.load_preset)
        self.update_preset_list()
        
        # Set placeholder
        self.preset_cb.set(self.PRESET_PLACEHOLDER)
        
        btn_frame = ttk.Frame(preset_frame)
        btn_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
        btn_frame.columnconfigure(0, weight=1)
        btn_frame.columnconfigure(1, weight=1)
        btn_frame.columnconfigure(2, weight=1)
        btn_frame.columnconfigure(3, weight=1)
        
        ttk.Button(btn_frame, text="保存", command=self.save_preset).grid(row=0, column=0, sticky="ew", padx=2)
        ttk.Button(btn_frame, text="删除", command=self.delete_preset).grid(row=0, column=1, sticky="ew", padx=2)
        ttk.Button(btn_frame, text="导入", command=self.import_preset_code).grid(row=0, column=2, sticky="ew", padx=2)
        ttk.Button(btn_frame, text="分享", command=self.export_preset_code).grid(row=0, column=3, sticky="ew", padx=2)

        # System
        sys_frame = ttk.Frame(self.root)
        sys_frame.pack(fill="x", padx=10, pady=10)
        sys_frame.columnconfigure(0, weight=1)
        sys_frame.columnconfigure(1, weight=1)
        sys_frame.columnconfigure(2, weight=1)
        
        ttk.Button(sys_frame, text="隐藏到托盘", command=self.minimize_to_tray).grid(row=0, column=0, sticky="ew", padx=2)
        
        self.startup_btn = ttk.Button(sys_frame, text="开机自启：关", command=self.toggle_startup)
        self.startup_btn.grid(row=0, column=1, sticky="ew", padx=2)
        
        if not self.is_admin:
             ttk.Button(sys_frame, text="管理员启动", command=self.restart_as_admin).grid(row=0, column=2, sticky="ew", padx=2)
        else:
             ttk.Button(sys_frame, text="取消管理员", command=self.restart_as_normal).grid(row=0, column=2, sticky="ew", padx=2)
        
        # Check initial startup status
        self.check_startup()
        
        # Hotkey Frame
        hk_frame = ttk.Frame(self.root)
        hk_frame.pack(fill="x", padx=10, pady=5)
        hk_frame.columnconfigure(0, weight=1)
        hk_frame.columnconfigure(1, weight=1)
        hk_frame.columnconfigure(2, weight=1)
        
        self.toggle_btn = ttk.Button(hk_frame, text="点击隐藏准星", command=self.toggle_crosshair_visible)
        self.toggle_btn.grid(row=0, column=0, sticky="ew", padx=2)
        
        self.hotkey_btn = ttk.Button(hk_frame, text="绑定隐藏准星键", command=self.bind_hotkey)
        self.hotkey_btn.grid(row=0, column=1, sticky="ew", padx=2)
        
        self.mode_cb = ttk.Combobox(hk_frame, textvariable=self.config['trigger_mode'], 
                                   values=["点击切换", "按住隐藏", "按住显示"], state="readonly", width=8)
        self.mode_cb.grid(row=0, column=2, sticky="ew", padx=2)
        self.mode_cb.bind("<<ComboboxSelected>>", self.apply_trigger)
        
        # Apply trigger if exists
        if self.config['hide_hotkey'].get():
            self.hotkey_btn.configure(text=f"快捷键: {self.config['hide_hotkey'].get()}")
            self.apply_trigger()

        # Status
        status_frame = ttk.Frame(self.root)
        status_frame.pack(side="bottom", fill="x", pady=(0, 5))

        self.status_label = ttk.Label(status_frame, text="作者moligod（B站抖音快手小红书同名）反馈群：856078436", foreground="green", anchor="center")
        self.status_label.pack(side="top", pady=2)
        
        ttk.Label(self.root, text="如若出现问题优先管理员启动，游戏内用快捷键必须管理员启动", foreground="red").pack(side="bottom", pady=(5, 0))

        # Version & Update
        ver_frame = ttk.Frame(status_frame)
        ver_frame.pack(side="top", pady=2)
        
        ttk.Label(ver_frame, text=f"当前版本: {self.version}", foreground="gray").pack(side="left", padx=5)
        update_link = ttk.Label(ver_frame, text="[检查更新]", foreground="blue", cursor="hand2")
        update_link.pack(side="left", padx=5)
        update_link.bind("<Button-1>", lambda e: self.open_update_url())

    def open_update_url(self):
        webbrowser.open("https://github.com/moligod/FPSDiyAim/releases")

    def log(self, msg):
        pass

    def toggle_crosshair_visible(self):
        if not self.overlay:
            return
        self.set_visible(not self.crosshair_visible)

    # 增强的触发器应用逻辑
    def apply_trigger(self, _=None):
        with self.trigger_lock:  # 增加线程锁
            # 彻底清理旧钩子
            try:
                # 尝试一次性清理所有热键，确保万无一失
                try: keyboard.unhook_all_hotkeys()
                except: pass
                try: keyboard.unhook_all() # 补充清理全局 hook
                except: pass
                
                for hook in self.keyboard_hooks:
                    try: keyboard.remove_hotkey(hook)
                    except: pass
                    try: keyboard.unhook(hook)
                    except: pass
            except: pass
            self.keyboard_hooks.clear()
            
            # 停止鼠标监听
            if self.mouse_listener:
                try: self.mouse_listener.stop()
                except: pass
                self.mouse_listener = None
            
            key = self.config['hide_hotkey'].get()
            if not key:
                self.hotkey_btn.configure(text="绑定隐藏准星键")
                return
                
            t_type = self.config.get('trigger_type', tk.StringVar(value="keyboard")).get()
            mode = self.config.get('trigger_mode', tk.StringVar(value="点击切换")).get()
            
            # Pre-convert key to string to avoid overhead in hook
        target_key = str(key)
        
        # Initial visibility state based on mode
        self.is_pressed = False
        self.log(f"Applying trigger mode={mode}, type={t_type}, key={target_key}")
        
        # Debug: Log all inputs to verify if hooks are active
        if t_type == "keyboard":
            self.log("Keyboard hook active. Press any key to test...")
        elif t_type == "mouse":
            self.log("Mouse hook active. Click to test...")

        if mode == "按住显示" or mode == "Hold_Show":
            self.log(f"Initializing {mode} -> Hide")
            # Force hide first
            self.crosshair_visible = True # Trick to force update
            self.set_visible(False, force=True)
        else:
            self.log(f"Initializing {mode} -> Show")
            self.crosshair_visible = False # Trick to force update
            self.set_visible(True, force=True)
        
        # Set new hooks
        try:
            if t_type == "keyboard":
                    # 获取按键的 scan codes，用于更精确的匹配
                    target_scan_codes = set()
                    try:
                        # 尝试获取 scan codes，如果失败（如组合键），则降级到仅使用名称匹配
                        codes = keyboard.key_to_scan_codes(target_key)
                        if codes:
                            target_scan_codes = set(codes)
                        self.log(f"Target key: {target_key}, Scancodes: {target_scan_codes}")
                    except:
                        pass

                    # 定义全局钩子函数
                    def on_key_event(event):
                        # Debug log for ALL key events
                        # self.log(f"Debug: Key {event.name} ({event.scan_code}) {event.event_type}")

                        # 检查是否匹配目标键
                        is_match = False
                        if target_scan_codes and event.scan_code in target_scan_codes:
                            is_match = True
                        elif event.name and event.name.lower() == target_key.lower():
                            is_match = True
                        
                        if is_match:
                            if mode in ["点击切换", "切换", "Toggle"]:
                                if event.event_type == keyboard.KEY_DOWN:
                                    self.log("Key Toggle")
                                    self.root.after(0, self.toggle_crosshair_visible)
                            elif mode in ["按住隐藏", "Hold_Hide"]:
                                if event.event_type == keyboard.KEY_DOWN:
                                    # Hide is cheap + repeat, use force=True (or force=False if too fast, but Hide is cheap)
                                    # Actually for Hide, repeats are fine.
                                    self.root.after(0, lambda: self.set_visible(False, force=True))
                                elif event.event_type == keyboard.KEY_UP:
                                    self.log("Key Hide UP")
                                    self.root.after(0, lambda: self.set_visible(True, force=True))
                            elif mode in ["按住显示", "Hold_Show"]:
                                if event.event_type == keyboard.KEY_DOWN:
                                    # Show is expensive + repeat, use force=False to rely on state check
                                    self.root.after(0, lambda: self.set_visible(True, force=False))
                                elif event.event_type == keyboard.KEY_UP:
                                    # Hide is cheap + once, force it
                                    self.log("Key Show UP")
                                    self.root.after(0, lambda: self.set_visible(False, force=True))
                    
                    # 注册全局钩子
                    hook = keyboard.hook(on_key_event)
                    self.keyboard_hooks.append(hook)
                        
            elif t_type == "mouse":
                    btn_map = {
                        "Button.left": pynput_mouse.Button.left,
                        "Button.right": pynput_mouse.Button.right,
                        "Button.middle": pynput_mouse.Button.middle,
                        "Button.x1": pynput_mouse.Button.x1,
                        "Button.x2": pynput_mouse.Button.x2
                    }
                    
                    target_btn = btn_map.get(target_key)
                    if not target_btn:
                        if "left" in target_key: target_btn = pynput_mouse.Button.left
                        elif "right" in target_key: target_btn = pynput_mouse.Button.right
                        elif "middle" in target_key: target_btn = pynput_mouse.Button.middle
                    
                    if not target_btn:
                        print(f"Unknown mouse button: {target_key}")
                        return

                    if mode == "点击切换" or mode == "切换" or mode == "Toggle":
                        def on_click(x, y, button, pressed):
                            if button == target_btn and pressed:
                                self.log("Mouse Toggle")
                                self.root.after(0, self.toggle_crosshair_visible)
                        
                        self.mouse_listener = pynput_mouse.Listener(on_click=on_click)
                        self.mouse_listener.start()
                        
                    elif mode == "按住隐藏" or mode == "Hold_Hide":
                        def on_click(x, y, button, pressed):
                            if button == target_btn:
                                if pressed:
                                    if not self.is_pressed:
                                        self.is_pressed = True
                                        self.log("Mouse Hide DOWN")
                                        self.root.after(0, lambda: self.set_visible(False, force=True))
                                else:
                                    self.is_pressed = False
                                    self.log("Mouse Hide UP")
                                    self.root.after(0, lambda: self.set_visible(True, force=True))
                        self.mouse_listener = pynput_mouse.Listener(on_click=on_click)
                        self.mouse_listener.start()
                        
                    elif mode == "按住显示" or mode == "Hold_Show":
                        def on_click(x, y, button, pressed):
                            # Debug log for ALL clicks
                            # self.log(f"Debug: Mouse {button} {'Pressed' if pressed else 'Released'}")
                            
                            if button == target_btn:
                                if pressed:
                                    # self.log("Mouse Show DOWN")
                                    self.root.after(0, lambda: self.set_visible(True, force=False))
                                else:
                                    self.log("Mouse Show UP")
                                    self.root.after(0, lambda: self.set_visible(False, force=True))
                        self.mouse_listener = pynput_mouse.Listener(on_click=on_click)
                        self.mouse_listener.start()
                        
        except Exception as e:
            print(f"Error applying trigger: {e}")
            messagebox.showerror("错误", f"快捷键绑定失败: {e}")

    def set_visible(self, visible, force=False):
        if not self.overlay: return
        
        # Optimization: Don't do anything if state hasn't changed
        if not force and visible == self.crosshair_visible:
            print(f"DEBUG: Skipping set_visible {visible} (current={self.crosshair_visible}, force={force})")
            return
            
        # Strategy: Use Canvas item state instead of window visibility
        # This is the fastest method and avoids window manager issues
        try:
            state = 'normal' if visible else 'hidden'
            self.overlay.canvas.itemconfigure("crosshair", state=state)
            
            if visible:
                self.overlay.deiconify()
                # Ensure topmost just in case, but don't force it every time to avoid overhead
                self.overlay.lift()
                self.overlay.attributes('-topmost', True)
                
                self.crosshair_visible = True
                self.toggle_btn.configure(text="点击隐藏准星")
            else:
                self.crosshair_visible = False
                self.toggle_btn.configure(text="点击显示准星")
        except Exception as e:
            print(f"Error setting visibility: {e}")

    def bind_hotkey(self):
        self.hotkey_btn.configure(text="按键/鼠标 (ESC取消)...")
        self.root.update()
        
        # Temp storage for hooks to remove them later
        temp_hooks = []
        
        def finish_bind():
            for h in temp_hooks:
                try:
                    keyboard.unhook(h)
                except: pass
                try:
                    # mouse.unhook(h)
                    if hasattr(h, 'stop'): h.stop()
                except: pass
            
            # Re-apply with new config
            self.apply_trigger()
            self.save_config()
            
            # Feedback
            key_name = self.config['hide_hotkey'].get()
            if key_name:
                self.root.after(100, lambda: messagebox.showinfo("绑定成功", f"快捷键已设置为: {key_name}\n模式: {self.config['trigger_mode'].get()}"))

        def on_key(event):
            if event.event_type == 'down':
                key_name = event.name
                
                if key_name.lower() == 'esc':
                    # Cancel/Clear
                    self.config['hide_hotkey'].set("")
                    self.hotkey_btn.configure(text="绑定隐藏准星键")
                    finish_bind()
                    return True
                
                self.config['hide_hotkey'].set(key_name)
                self.config['trigger_type'].set("keyboard")
                self.hotkey_btn.configure(text=f"快捷键: {key_name}")
                finish_bind()
                return True

        def on_click(x, y, button, pressed):
            if pressed:
                btn_name = str(button) # e.g. Button.left
                self.config['hide_hotkey'].set(btn_name)
                self.config['trigger_type'].set("mouse")
                self.hotkey_btn.configure(text=f"快捷键: {btn_name}")
                finish_bind()
                # Stop listener
                return False 

        # Hook both
        h_k = keyboard.hook(on_key)
        # h_m = mouse.hook(on_mouse)
        h_m = pynput_mouse.Listener(on_click=on_click)
        h_m.start()
        
        temp_hooks.extend([h_k, h_m])

    def add_slider(self, parent, label, var, min_val, max_val, row):
        ttk.Label(parent, text=label).grid(row=row, column=0, padx=5, pady=2)
        scale = ttk.Scale(parent, from_=min_val, to=max_val, variable=var, orient="horizontal", command=self.update_overlay)
        scale.grid(row=row, column=1, sticky="ew", padx=5, pady=2)

    def choose_color(self):
        color = colorchooser.askcolor(color=self.config['color'].get())[1]
        if color:
            self.config['color'].set(color)
            self.update_overlay()
            
    def on_style_change(self, event=None):
        style = self.config['style'].get()
        # Always enable image button so user can click it directly to switch mode
        # If user switches dropdown manually, we check if image path is needed
        # Only prompt for image if it's a user interaction (event is not None) or if path is truly empty during init
        if (style == "Custom" or style in ["自定义图片", "自定义"]):
            if not self.config['image_path'].get():
                # If switched to Custom but no image, prompt to choose
                # Use 'after' to avoid blocking the event loop immediately
                if event and event != "Startup": # Only prompt if user manually triggered, not on startup
                    self.root.after(100, self.choose_image)
            else:
                # If we have a path, ensure overlay updates
                self.update_overlay()
            
        self.update_overlay()

    def choose_image(self):
        file_path = filedialog.askopenfilename(
            title="选择图片",
            filetypes=[("Image Files", "*.png;*.gif;*.jpg;*.jpeg;*.bmp;*.ppm;*.pnm")]
        )
        if file_path:
            self.config['image_path'].set(file_path)
            # Auto switch to Custom style
            self.config['style'].set("自定义图片")
            self.update_overlay()

    def drag_start(self, event):
        self._drag_start_x = event.x_root
        self._drag_start_y = event.y_root
        try:
            self._start_pos_x = int(float(self.pos_x.get()))
            self._start_pos_y = int(float(self.pos_y.get()))
        except:
            self._start_pos_x = self.screen_w // 2
            self._start_pos_y = self.screen_h // 2

    def drag_move(self, event):
        dx = event.x_root - self._drag_start_x
        dy = event.y_root - self._drag_start_y
        self.pos_x.set(str(self._start_pos_x + dx))
        self.pos_y.set(str(self._start_pos_y + dy))
        self.update_pos()

    def start_overlay(self):
        if self.overlay:
            self.overlay.destroy()
        self.overlay = CrosshairOverlay(self.root, self.config)
        self.update_pos()

    def update_overlay(self, _=None):
        if self.overlay:
            try:
                self.overlay.redraw()
            finally:
                # Ensure the newly drawn items align with current visibility state
                # so that switching样式或启动后的首次绘制不会出现状态不同步导致的“看不见”
                self.set_visible(self.crosshair_visible, force=True)

    def update_pos(self, *args):
        if self.overlay:
            try:
                # Handle potential float strings or empty strings
                x_str = self.pos_x.get().strip()
                y_str = self.pos_y.get().strip()
                
                if not x_str: x = self.screen_w // 2
                else: x = int(float(x_str))
                
                if not y_str: y = self.screen_h // 2
                else: y = int(float(y_str))
                
                self.overlay.set_position(x, y)
            except Exception as e:
                # print(f"Invalid position input: {e}")
                pass

    def center_pos(self):
        self.pos_x.set(str(self.screen_w // 2))
        self.pos_y.set(str(self.screen_h // 2))
        self.update_pos()

    def adjust_pos(self, dx, dy):
        try:
            current_x = int(float(self.pos_x.get()))
            current_y = int(float(self.pos_y.get()))
        except:
            current_x = self.screen_w // 2
            current_y = self.screen_h // 2
            
        self.pos_x.set(str(current_x + dx))
        self.pos_y.set(str(current_y + dy))
        self.update_pos()

    def update_preset_list(self):
        # Clean any accidental placeholder saved as a preset
        if self.PRESET_PLACEHOLDER in self.presets:
            try: del self.presets[self.PRESET_PLACEHOLDER]
            except: pass
        names = [n for n in self.presets.keys() if n != self.PRESET_PLACEHOLDER]
        self.preset_cb['values'] = names

    def save_preset(self):
        name = self.current_preset_name.get().strip()
        if not name or name == self.PRESET_PLACEHOLDER:
            messagebox.showerror("无效名称", "请先在上方输入或选择一个有效的预设名称。")
            return
            
        # Capture current settings
        preset_data = {
            "size": self.config['size'].get(),
            "thickness": self.config['thickness'].get(),
            "color": self.config['color'].get(),
            "dot": self.config['dot'].get(),
            "img_scale": self.config['img_scale'].get(),
            "style": self.config['style'].get(),
            "image_path": self.config['image_path'].get()
        }
        
        self.presets[name] = preset_data
        self.update_preset_list()
        
    def load_preset(self, event=None):
        name = self.current_preset_name.get()
        if name == self.PRESET_PLACEHOLDER:
            return
            
        if name in self.presets:
            data = self.presets[name]
            self.config['size'].set(data.get("size", 20))
            self.config['thickness'].set(data.get("thickness", 2))
            self.config['color'].set(data.get("color", "#00FF00"))
            self.config['dot'].set(data.get("dot", 4))
            self.config['img_scale'].set(data.get("img_scale", 100))
            self.config['style'].set(data.get("style", "十字"))
            self.config['image_path'].set(data.get("image_path", ""))
            
            # Refresh overlay
            self.on_style_change(event="PresetLoad")
            
    def delete_preset(self):
        name = self.current_preset_name.get()
        if name in self.presets:
            del self.presets[name]
            self.current_preset_name.set("")
            self.update_preset_list()

    def export_preset_code(self):
        name = self.current_preset_name.get()
        if not name:
            # If no preset selected, maybe export current settings?
            # Let's ask for a name first
            name = simpledialog.askstring("分享配置", "请为当前配置命名：", initialvalue="我的配置")
            if not name: return
            
            # Create temp data from current settings
            data = {
                "size": self.config['size'].get(),
                "thickness": self.config['thickness'].get(),
                "color": self.config['color'].get(),
                "dot": self.config['dot'].get(),
                "img_scale": self.config['img_scale'].get(),
                "style": self.config['style'].get(),
                "image_path": "" # Don't export local path
            }
        elif name in self.presets:
            data = self.presets[name].copy()
            data['image_path'] = "" # Clear image path for sharing
        else:
            return

        # Check if custom image (warn user)
        if self.config['style'].get() in ["Custom", "自定义图片", "自定义"]:
            messagebox.showwarning("提示", "自定义图片无法通过口令分享。\n对方只能看到配置参数，图片需要手动设置。")

        data['name'] = name
        
        try:
            # JSON -> String -> Bytes -> Compress -> Base64
            json_str = json.dumps(data, ensure_ascii=False)
            compressed = zlib.compress(json_str.encode('utf-8'))
            b64_str = base64.b64encode(compressed).decode('utf-8')
            
            code = f"MOLI#{b64_str}"
            
            # Copy to clipboard
            self.root.clipboard_clear()
            self.root.clipboard_append(code)
            self.root.update()
            
            messagebox.showinfo("分享成功", f"分享口令已复制到剪贴板！\n\n发送给好友，点击“导入”即可使用。")
            
        except Exception as e:
            messagebox.showerror("错误", f"生成口令失败: {e}")

    def import_preset_code(self):
        code = simpledialog.askstring("导入口令", "请粘贴 MOLI# 开头的分享口令：")
        if not code: return
        
        code = code.strip()
        if not code.startswith("MOLI#"):
            messagebox.showerror("错误", "无效的口令格式 (必须以 MOLI# 开头)")
            return
            
        try:
            b64_str = code[5:]
            compressed = base64.b64decode(b64_str)
            json_str = zlib.decompress(compressed).decode('utf-8')
            data = json.loads(json_str)
            
            name = data.get('name', '导入配置')
            
            # Check for conflict
            if name in self.presets:
                if not messagebox.askyesno("覆盖确认", f"方案 [{name}] 已存在，是否覆盖？"):
                    return

            # Sanitize data
            clean_data = {
                "size": data.get("size", 20),
                "thickness": data.get("thickness", 2),
                "color": data.get("color", "#00FF00"),
                "dot": data.get("dot", 4),
                "img_scale": data.get("img_scale", 100),
                "style": data.get("style", "十字"),
                "image_path": ""
            }
            
            self.presets[name] = clean_data
            self.current_preset_name.set(name)
            self.update_preset_list()
            self.load_preset()
            
            messagebox.showinfo("成功", f"方案 [{name}] 导入成功！")
            
        except Exception as e:
            messagebox.showerror("错误", f"解析口令失败: {e}")

    def resource_path(self, relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)

    def load_config(self):
        data = {}
        # Try loading from Registry first
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\MoligodCrosshair", 0, winreg.KEY_READ)
            json_str, _ = winreg.QueryValueEx(key, "Config")
            data = json.loads(json_str)
            winreg.CloseKey(key)
        except FileNotFoundError:
            pass # No config found, use defaults
        except Exception as e:
            print(f"Error loading registry config: {e}")

        # Apply config
        self.pos_x.set(str(data.get("pos_x", self.screen_w // 2)))
        self.pos_y.set(str(data.get("pos_y", self.screen_h // 2)))
        
        self.config['size'].set(data.get("size", 20))
        self.config['thickness'].set(data.get("thickness", 2))
        self.config['color'].set(data.get("color", "#00FF00"))
        self.config['dot'].set(data.get("dot", 4))
        _style_val = data.get("style", "十字")
        if _style_val == "自定义":
            _style_val = "自定义图片"
        self.config['style'].set(_style_val)
        self.config['img_scale'].set(data.get("img_scale", 100))
        self.config['image_path'].set(data.get("image_path", ""))
        self.config['force_admin'].set(data.get("force_admin", False))
        self.config['hide_hotkey'].set(data.get("hide_hotkey", ""))
        self.config['trigger_type'].set(data.get("trigger_type", "keyboard"))
        val = data.get("trigger_mode", "点击切换")
        if val == "切换": val = "点击切换"
        self.config['trigger_mode'].set(val)
        
        self.presets = data.get("presets", {})

    def save_config(self):
        try:
            x = int(float(self.pos_x.get()))
            y = int(float(self.pos_y.get()))
        except:
            x = self.screen_w // 2
            y = self.screen_h // 2
            
        data = {
            "pos_x": x,
            "pos_y": y,
            "size": self.config['size'].get(),
            "thickness": self.config['thickness'].get(),
            "color": self.config['color'].get(),
            "dot": self.config['dot'].get(),
            "img_scale": self.config['img_scale'].get(),
            "style": self.config['style'].get(),
            "image_path": self.config['image_path'].get(),
            "force_admin": self.config['force_admin'].get(),
            "hide_hotkey": self.config['hide_hotkey'].get(),
            "trigger_type": self.config['trigger_type'].get(),
            "trigger_mode": self.config['trigger_mode'].get(),
            "presets": self.presets
        }
        
        # Save to Registry
        try:
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\MoligodCrosshair")
            json_str = json.dumps(data, ensure_ascii=False)
            winreg.SetValueEx(key, "Config", 0, winreg.REG_SZ, json_str)
            winreg.CloseKey(key)
        except Exception as e:
            print(f"Error saving config to registry: {e}")

def check_force_admin():
    # Helper to check config before initializing UI
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\MoligodCrosshair", 0, winreg.KEY_READ)
        json_str, _ = winreg.QueryValueEx(key, "Config")
        data = json.loads(json_str)
        winreg.CloseKey(key)
        return data.get("force_admin", False)
    except:
        pass
    return False

if __name__ == "__main__":
    # Check if we should force admin
    should_be_admin = check_force_admin()
    is_admin = ctypes.windll.shell32.IsUserAnAdmin()
    
    if should_be_admin and not is_admin:
        # Relaunch as admin
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
        sys.exit()
    else:
        ControlPanel()
