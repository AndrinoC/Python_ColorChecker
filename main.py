# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk
import cv2
import numpy as np
import os
import json
import keyboard
import pyautogui
import mss
import traceback
import time
import pygetwindow as gw
import pydirectinput
import threading
import queue
import collections

CONFIG_FILE = "color_checker_config.json"
DEFAULT_AREA_TOLERANCE = 15
DEFAULT_CAPTURE_BOX_SIZE = 300
DEFAULT_CLICK_METHOD = "LAB"
DEFAULT_CLICK_BUTTON = "left"
DEFAULT_CLICKING_ENABLED = True
DEFAULT_PAUSE_HOTKEY = 'ctrl+shift+x'
DEFAULT_TOGGLE_CLICK_HOTKEY = 'ctrl+shift+c'
WORKER_SLEEP_TIME = 0.001
CLICK_COOLDOWN = 0.05
UI_UPDATE_INTERVAL = 50
COLOR_ANALYSIS_ENABLED = True
COLOR_ANALYSIS_RESIZE_WIDTH = 64
COLOR_ANALYSIS_K = 10
COLOR_ANALYSIS_MIN_PERCENT = 2.0
MOUSE_PIXEL_POLL_INTERVAL = 75

class ColorBotApp:
    def __init__(self, root):
        self.root = root
        self.sct_main = mss.mss()
        self.worker_thread = None
        self.stop_event = threading.Event()
        self.update_queue = queue.Queue()
        self.config_lock = threading.Lock()
        self.active_hotkeys = {}
        self.key_listener_hook = None

        self.color1_rgb = (204, 204, 204); self.color1_lab = None
        self.color2_rgb = (38, 120, 122); self.color2_lab = None
        self.area_tolerance = DEFAULT_AREA_TOLERANCE
        self.capture_box_size = DEFAULT_CAPTURE_BOX_SIZE
        self.click_method = DEFAULT_CLICK_METHOD
        self.click_button = DEFAULT_CLICK_BUTTON
        self.is_paused = True
        self.last_click_time = 0
        self.is_picker_active = False
        self.clicking_enabled = DEFAULT_CLICKING_ENABLED
        self.pause_hotkey_str = DEFAULT_PAUSE_HOTKEY
        self.toggle_click_hotkey_str = DEFAULT_TOGGLE_CLICK_HOTKEY
        self.is_listening_for_hotkey = False
        self.hotkey_target_widget = None; self.hotkey_target_attr = None
        self.detected_colors_tab_active = False
        self.mouse_pixel_polling_active = False

        self._init_ui()
        self._load_config()
        self._start_worker()
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
        self.root.after(UI_UPDATE_INTERVAL, self._check_queue)
        self.root.after(MOUSE_PIXEL_POLL_INTERVAL, self._update_mouse_pixel_info)

    def _init_ui(self):
        self.root.title("Color Detection & Clicker v3.1")
        self.root.geometry("950x700")
        self.root.configure(bg='#2E2E2E')

        style = ttk.Style(self.root); style.theme_use('clam')
        style.configure("TFrame", background="#333333")
        style.configure("TLabel", background="#333333", foreground="white", font=("Arial", 10))
        style.configure("TButton", background="#555555", foreground="white", font=("Arial", 10), padding=5)
        style.map("TButton", background=[('active', '#777777')])
        style.configure("TScale", background="#333333", troughcolor="#555555", sliderthickness=15)
        style.configure("TRadiobutton", background="#333333", foreground="white", font=("Arial", 10))
        style.map("TRadiobutton", background=[('active', '#444444')])
        style.configure("TCheckbutton", background="#333333", foreground="white", font=("Arial", 10))
        style.map("TCheckbutton", background=[('active', '#444444')])
        style.configure("Header.TLabel", font=("Arial", 11, "bold"))
        style.configure("Status.TLabel", font=("Arial", 12, "bold"))
        style.configure("ColorSwatch.TLabel", borderwidth=1, relief="solid", padding=2)
        style.configure("Picker.TButton", font=("Arial", 9), padding=(2,1))
        style.configure("Hotkey.TButton", font=("Arial", 9), padding=(3,1))
        style.configure("Treeview.Heading", font=('Arial', 10, 'bold'))
        style.configure("TNotebook", background="#2E2E2E", borderwidth=0)
        style.configure("TNotebook.Tab", font=("Arial", 10, "bold"), padding=[10, 5], background="#444444", foreground="#DDDDDD")
        style.map("TNotebook.Tab", background=[("selected", "#555555"), ('active', '#666666')], foreground=[("selected", "white")])
        style.configure("Small.TLabel", font=("Arial", 9))
        style.configure("PixelFinder.TLabel", font=("Arial", 10), anchor='w')
        style.configure("PixelFinderValue.TLabel", font=("Arial", 10), foreground="#A0A0A0", anchor='w')

        top_frame = ttk.Frame(self.root, padding=10); top_frame.pack(side=tk.TOP, fill=tk.X, pady=5)
        ttk.Label(top_frame, text="Bot Status:", style="Header.TLabel").pack(side=tk.LEFT, padx=5)
        self.status_var = tk.StringVar(value="Paused")
        self.status_value_label = ttk.Label(top_frame, textvariable=self.status_var, foreground="red", style="Status.TLabel")
        self.status_value_label.pack(side=tk.LEFT, padx=5)
        self.toggle_button = ttk.Button(top_frame, text="Pause/Resume", command=self._toggle_script)
        self.toggle_button.pack(side=tk.RIGHT, padx=5)

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        settings_tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(settings_tab, text=' Settings ')

        settings_left_frame = ttk.Frame(settings_tab); settings_left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        settings_right_frame = ttk.Frame(settings_tab); settings_right_frame.pack(side=tk.RIGHT, fill=tk.Y)

        controls_frame = ttk.Frame(settings_left_frame, padding=10); controls_frame.pack(fill=tk.BOTH, expand=True)
        controls_frame.columnconfigure(1, weight=1)
        controls_frame.columnconfigure(3, weight=1)

        self.color1_label = ttk.Label(controls_frame, text="Color 1", style="ColorSwatch.TLabel", width=15, anchor='center')
        self.color1_label.grid(row=0, column=0, padx=5, pady=5, sticky='ew')
        self.color1_rgb_entry = ttk.Entry(controls_frame, width=15, font=("Arial", 10))
        self.color1_rgb_entry.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        self.color1_value = tk.StringVar(value="Area: 0.00% / 0 px")
        ttk.Label(controls_frame, textvariable=self.color1_value, anchor='w', style="Small.TLabel").grid(row=0, column=3, padx=5, pady=5, sticky='ew')

        self.color2_label = ttk.Label(controls_frame, text="Color 2", style="ColorSwatch.TLabel", width=15, anchor='center')
        self.color2_label.grid(row=1, column=0, padx=5, pady=5, sticky='ew')
        self.color2_rgb_entry = ttk.Entry(controls_frame, width=15, font=("Arial", 10))
        self.color2_rgb_entry.grid(row=1, column=1, padx=5, pady=5, sticky='ew')
        self.color2_value = tk.StringVar(value="Area: 0.00% / 0 px")
        ttk.Label(controls_frame, textvariable=self.color2_value, anchor='w', style="Small.TLabel").grid(row=1, column=3, padx=5, pady=5, sticky='ew')

        ttk.Separator(controls_frame, orient=tk.HORIZONTAL).grid(row=2, column=0, columnspan=4, sticky='ew', pady=10)

        ttk.Label(controls_frame, text="Area Tolerance (ΔE):", anchor='w').grid(row=3, column=0, padx=5, pady=3, sticky='w')
        self.area_tolerance_var = tk.IntVar(value=self.area_tolerance)
        self.area_tolerance_value_label = ttk.Label(controls_frame, text=str(self.area_tolerance), width=4)
        self.area_tolerance_value_label.grid(row=3, column=1, padx=(0,5), pady=3, sticky='w')
        self.area_tolerance_slider = ttk.Scale(controls_frame, from_=0, to=100, orient=tk.HORIZONTAL, variable=self.area_tolerance_var, command=self._update_area_tolerance, length=200)
        self.area_tolerance_slider.grid(row=4, column=0, columnspan=4, padx=5, pady=2, sticky='ew')

        ttk.Label(controls_frame, text="Detection Box Size (px):", anchor='w').grid(row=5, column=0, padx=5, pady=3, sticky='w')
        self.box_size_var = tk.IntVar(value=self.capture_box_size)
        self.box_size_value_label = ttk.Label(controls_frame, text=str(self.capture_box_size), width=4)
        self.box_size_value_label.grid(row=5, column=1, padx=(0,5), pady=3, sticky='w')
        self.box_size_slider = ttk.Scale(controls_frame, from_=10, to=600, orient=tk.HORIZONTAL, variable=self.box_size_var, command=self._update_box_size, length=200)
        self.box_size_slider.grid(row=6, column=0, columnspan=4, padx=5, pady=2, sticky='ew')

        ttk.Separator(controls_frame, orient=tk.HORIZONTAL).grid(row=7, column=0, columnspan=4, sticky='ew', pady=10)

        ttk.Label(controls_frame, text="Click Trigger Settings", style="Header.TLabel").grid(row=8, column=0, columnspan=4, pady=(5,5), sticky='w')
        self.clicking_enabled_var = tk.BooleanVar(value=self.clicking_enabled)
        self.click_enable_check = ttk.Checkbutton(controls_frame, text="Enable Clicking", variable=self.clicking_enabled_var, command=self._toggle_clicking_callback)
        self.click_enable_check.grid(row=9, column=0, columnspan=2, padx=5, pady=5, sticky='w')

        center_pixel_frame = ttk.Frame(controls_frame); center_pixel_frame.grid(row=10, column=0, columnspan=4, pady=(0, 5), sticky='ew')
        ttk.Label(center_pixel_frame, text="Center Pixel:", anchor='w').pack(side=tk.LEFT, padx=(5,0))
        self.center_pixel_rgb_value = tk.StringVar(value="N/A")
        ttk.Label(center_pixel_frame, textvariable=self.center_pixel_rgb_value, width=12).pack(side=tk.LEFT, padx=5)
        self.center_pixel_swatch = tk.Label(center_pixel_frame, text="", bg="black", width=2, height=1, relief="sunken")
        self.center_pixel_swatch.pack(side=tk.LEFT, padx=5)

        click_method_frame = ttk.Frame(controls_frame); click_method_frame.grid(row=11, column=0, columnspan=4, pady=2, sticky='w')
        ttk.Label(click_method_frame, text="Click Match:", anchor='w').pack(side=tk.LEFT, padx=5)
        self.click_method_var = tk.StringVar(value=self.click_method)
        ttk.Radiobutton(click_method_frame, text="RGB", variable=self.click_method_var, value="RGB", command=self._update_click_settings).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(click_method_frame, text="LAB (ΔE)", variable=self.click_method_var, value="LAB", command=self._update_click_settings).pack(side=tk.LEFT, padx=5)

        click_button_frame = ttk.Frame(controls_frame); click_button_frame.grid(row=12, column=0, columnspan=4, pady=5, sticky='w')
        ttk.Label(click_button_frame, text="Click Button:", anchor='w').pack(side=tk.LEFT, padx=5)
        self.click_button_var = tk.StringVar(value=self.click_button)
        ttk.Radiobutton(click_button_frame, text="Left", variable=self.click_button_var, value="left", command=self._update_click_settings).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(click_button_frame, text="Right", variable=self.click_button_var, value="right", command=self._update_click_settings).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(click_button_frame, text="Middle", variable=self.click_button_var, value="middle", command=self._update_click_settings).pack(side=tk.LEFT, padx=5)

        ttk.Separator(controls_frame, orient=tk.HORIZONTAL).grid(row=13, column=0, columnspan=4, sticky='ew', pady=10)

        ttk.Label(controls_frame, text="Hotkey Settings", style="Header.TLabel").grid(row=14, column=0, columnspan=4, pady=(5,5), sticky='w')
        ttk.Label(controls_frame, text="Pause/Resume:", anchor='w').grid(row=15, column=0, padx=5, pady=5, sticky='w')
        self.pause_hotkey_label = ttk.Label(controls_frame, text=self.pause_hotkey_str, width=15, relief="sunken", padding=2)
        self.pause_hotkey_label.grid(row=15, column=1, padx=5, pady=5, sticky='ew')
        self.pause_hotkey_button = ttk.Button(controls_frame, text="Set", style="Hotkey.TButton", command=lambda: self._set_hotkey_listener('pause_hotkey_str', self.pause_hotkey_label))
        self.pause_hotkey_button.grid(row=15, column=2, padx=5, pady=5)

        ttk.Label(controls_frame, text="Toggle Clicking:", anchor='w').grid(row=16, column=0, padx=5, pady=5, sticky='w')
        self.toggle_click_hotkey_label = ttk.Label(controls_frame, text=self.toggle_click_hotkey_str, width=15, relief="sunken", padding=2)
        self.toggle_click_hotkey_label.grid(row=16, column=1, padx=5, pady=5, sticky='ew')
        self.toggle_click_hotkey_button = ttk.Button(controls_frame, text="Set", style="Hotkey.TButton", command=lambda: self._set_hotkey_listener('toggle_click_hotkey_str', self.toggle_click_hotkey_label))
        self.toggle_click_hotkey_button.grid(row=16, column=2, padx=5, pady=5)

        preview_frame = ttk.Frame(settings_right_frame, padding=10); preview_frame.pack(fill=tk.BOTH, expand=True)
        self.preview_size = (200, 200)
        ttk.Label(preview_frame, text="Detection Area Overlay", anchor='center').pack(pady=(0, 2))
        self.settings_overlay_preview_label = tk.Label(preview_frame, bg="black", width=self.preview_size[0], height=self.preview_size[1], relief="sunken")
        self.settings_overlay_preview_label.pack(pady=5)
        ttk.Label(preview_frame, text="Detection Area Capture", anchor='center').pack(pady=(5, 2))
        self.settings_capture_preview_label = tk.Label(preview_frame, bg="black", width=self.preview_size[0], height=self.preview_size[1], relief="sunken")
        self.settings_capture_preview_label.pack(pady=5)

        colors_tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(colors_tab, text=' Detected Colors ')

        colors_top_frame = ttk.Frame(colors_tab); colors_top_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 10))
        colors_top_left = ttk.Frame(colors_top_frame); colors_top_left.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        colors_top_right = ttk.Frame(colors_top_frame); colors_top_right.pack(side=tk.RIGHT, fill=tk.NONE)

        finder_frame = ttk.Frame(colors_top_left, padding=5, relief="groove", borderwidth=1)
        finder_frame.pack(fill=tk.X, expand=True)
        ttk.Label(finder_frame, text="Live Mouse Pixel Info:", style="Header.TLabel").grid(row=0, column=0, columnspan=3, pady=(0, 5), sticky='w')

        ttk.Label(finder_frame, text="Coords:", style="PixelFinder.TLabel").grid(row=1, column=0, padx=5, pady=2, sticky='w')
        self.mouse_coords_var = tk.StringVar(value="X: ---, Y: ---")
        ttk.Label(finder_frame, textvariable=self.mouse_coords_var, style="PixelFinderValue.TLabel").grid(row=1, column=1, columnspan=2, padx=5, pady=2, sticky='w')

        ttk.Label(finder_frame, text="RGB:", style="PixelFinder.TLabel").grid(row=2, column=0, padx=5, pady=2, sticky='w')
        self.mouse_rgb_var = tk.StringVar(value="---, ---, ---")
        ttk.Label(finder_frame, textvariable=self.mouse_rgb_var, style="PixelFinderValue.TLabel").grid(row=2, column=1, padx=5, pady=2, sticky='w')
        self.mouse_swatch_label = tk.Label(finder_frame, text="", bg="black", width=2, height=1, relief="sunken")
        self.mouse_swatch_label.grid(row=2, column=2, padx=5, pady=2, sticky='w')

        ttk.Label(finder_frame, text="LAB:", style="PixelFinder.TLabel").grid(row=3, column=0, padx=5, pady=2, sticky='w')
        self.mouse_lab_var = tk.StringVar(value="---, ---, ---")
        ttk.Label(finder_frame, textvariable=self.mouse_lab_var, style="PixelFinderValue.TLabel").grid(row=3, column=1, columnspan=2, padx=5, pady=2, sticky='w')

        colors_preview_frame = ttk.Frame(colors_top_right)
        colors_preview_frame.pack(fill=tk.BOTH)
        self.colors_preview_size = (150, 150)
        ttk.Label(colors_preview_frame, text="Overlay", anchor='center', style="Small.TLabel").pack()
        self.colors_overlay_preview_label = tk.Label(colors_preview_frame, bg="black", width=self.colors_preview_size[0], height=self.colors_preview_size[1], relief="sunken")
        self.colors_overlay_preview_label.pack(pady=(2, 5))
        ttk.Label(colors_preview_frame, text="Capture", anchor='center', style="Small.TLabel").pack()
        self.colors_capture_preview_label = tk.Label(colors_preview_frame, bg="black", width=self.colors_preview_size[0], height=self.colors_preview_size[1], relief="sunken")
        self.colors_capture_preview_label.pack(pady=(2, 0))

        colors_bottom_frame = ttk.Frame(colors_tab); colors_bottom_frame.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)
        ttk.Label(colors_bottom_frame, text=f"Dominant Colors in Capture Area (Top {COLOR_ANALYSIS_K}, >{COLOR_ANALYSIS_MIN_PERCENT:.1f}%):", anchor='w').pack(pady=(0, 5), fill=tk.X)

        cols = ('swatch', 'rgb', 'lab', 'percent')
        self.color_tree = ttk.Treeview(colors_bottom_frame, columns=cols, show='headings', height=8)
        self.color_tree.heading('swatch', text='Color'); self.color_tree.heading('rgb', text='RGB')
        self.color_tree.heading('lab', text='LAB'); self.color_tree.heading('percent', text='%')
        self.color_tree.column('swatch', width=60, anchor=tk.CENTER); self.color_tree.column('rgb', width=120, anchor=tk.W)
        self.color_tree.column('lab', width=120, anchor=tk.W); self.color_tree.column('percent', width=70, anchor=tk.E)

        tree_scrollbar = ttk.Scrollbar(colors_bottom_frame, orient="vertical", command=self.color_tree.yview)
        self.color_tree.configure(yscrollcommand=tree_scrollbar.set)
        tree_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.color_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.color1_rgb_entry.bind("<Return>", lambda e: self._set_color_variable(self.color1_label, self.color1_rgb_entry, "color1"))
        self.color1_rgb_entry.bind("<FocusOut>", lambda e: self._set_color_variable(self.color1_label, self.color1_rgb_entry, "color1"))
        self.color2_rgb_entry.bind("<Return>", lambda e: self._set_color_variable(self.color2_label, self.color2_rgb_entry, "color2"))
        self.color2_rgb_entry.bind("<FocusOut>", lambda e: self._set_color_variable(self.color2_label, self.color2_rgb_entry, "color2"))
        self.color1_label.bind("<Button-1>", lambda e: self._pick_color_from_screen(self.color1_label, self.color1_rgb_entry, "color1"))
        self.color2_label.bind("<Button-1>", lambda e: self._pick_color_from_screen(self.color2_label, self.color2_rgb_entry, "color2"))


    def _rgb_to_lab(self, rgb_tuple):
        if not isinstance(rgb_tuple, (tuple, list)) or len(rgb_tuple) != 3: return (0, 0, 0)
        try:
            rgb_clamped = tuple(max(0, min(255, int(c))) for c in rgb_tuple)
            rgb_np = np.uint8([[rgb_clamped]])
            lab_np = cv2.cvtColor(rgb_np, cv2.COLOR_RGB2LAB)
            return tuple(map(int, lab_np[0][0]))
        except Exception: return (0, 0, 0)

    def _lab_to_rgb(self, lab_tuple):
        if not isinstance(lab_tuple, (tuple, list)) or len(lab_tuple) != 3: return (0, 0, 0)
        try:
            lab_np = np.uint8([[lab_tuple]])
            rgb_np = cv2.cvtColor(lab_np, cv2.COLOR_LAB2RGB)
            return tuple(map(int, rgb_np[0][0]))
        except Exception: return (0, 0, 0)

    def _update_area_tolerance(self, val_str):
        with self.config_lock:
            try: self.area_tolerance = max(0, min(100, int(float(val_str))))
            except (ValueError, tk.TclError): pass
        if hasattr(self, 'area_tolerance_value_label'): self.area_tolerance_value_label.config(text=str(self.area_tolerance))

    def _update_box_size(self, val_str):
        with self.config_lock:
            try: self.capture_box_size = max(10, min(600, int(float(val_str))))
            except (ValueError, tk.TclError): pass
        if hasattr(self, 'box_size_value_label'): self.box_size_value_label.config(text=str(self.capture_box_size))

    def _update_click_settings(self):
        with self.config_lock:
            self.click_method = self.click_method_var.get()
            self.click_button = self.click_button_var.get()

    def _update_color_label_bg(self, label, rgb_tuple):
        try:
            r, g, b = [max(0, min(255, int(c))) for c in rgb_tuple]
            hex_color = f'#{r:02x}{g:02x}{b:02x}'
            luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
            text_color = "black" if luminance > 0.5 else "white"
            label.config(background=hex_color, foreground=text_color)
        except Exception: label.config(background="black", foreground="white", text="Error")

    def _parse_rgb_string(self, rgb_string):
        try:
            parts = [int(c.strip()) for c in rgb_string.split(',')]
            return tuple(parts) if len(parts) == 3 and all(0 <= c <= 255 for c in parts) else None
        except ValueError: return None

    def _set_color_variable(self, label_widget, entry_widget, color_var_name):
        rgb_tuple = self._parse_rgb_string(entry_widget.get())
        with self.config_lock:
            current_color = self.color1_rgb if color_var_name == "color1" else self.color2_rgb
            if rgb_tuple is None:
                entry_widget.delete(0, tk.END)
                entry_widget.insert(0, f"{current_color[0]}, {current_color[1]}, {current_color[2]}")
                return
            lab_value = self._rgb_to_lab(rgb_tuple)
            if color_var_name == "color1": self.color1_rgb, self.color1_lab = rgb_tuple, lab_value
            else: self.color2_rgb, self.color2_lab = rgb_tuple, lab_value
        self._update_color_label_bg(label_widget, rgb_tuple)
        label_widget.config(text=f"{rgb_tuple[0]},{rgb_tuple[1]},{rgb_tuple[2]}")

    def _pick_color_from_screen(self, label_widget, entry_widget, color_var_name):
        PAN_SPEED_FACTOR = 0.4
        if self.is_picker_active: return
        self.is_picker_active = True
        print("Opening color picker...")
        screenshot_pil, screenshot_np_rgb, img_width, img_height = None, None, 0, 0

        try:
            monitor_info = self.sct_main.monitors[0]
            print(f"Picker capturing monitor area: {monitor_info}")
            if not monitor_info or monitor_info["width"] <= 0 or monitor_info["height"] <= 0:
                raise Exception(f"Invalid monitor info for full desktop: {monitor_info}")

            sct_img = self.sct_main.grab(monitor_info)
            screenshot_np_bgra = np.array(sct_img)
            screenshot_np_rgb = cv2.cvtColor(screenshot_np_bgra, cv2.COLOR_BGRA2RGB)
            screenshot_pil = Image.frombytes("RGB", sct_img.size, sct_img.rgb)
            img_width, img_height = screenshot_pil.size

        except Exception as e:
            print(f"Picker Error (Grab): {e}"); traceback.print_exc()
            messagebox.showerror("Picker Error", f"Could not grab screen:\n{e}")
            self.is_picker_active = False
            return

        picker_window = tk.Toplevel(self.root)
        max_win_w = self.root.winfo_screenwidth() - 100
        max_win_h = self.root.winfo_screenheight() - 100
        win_w = min(img_width + 40, max_win_w)
        win_h = min(img_height + 40, max_win_h)
        picker_window.geometry(f"{win_w}x{win_h}")
        picker_window.title("LMB: Pick | MMB/RMB Drag: Scroll | Esc: Cancel")
        picker_window.attributes("-topmost", True)

        def on_esc(event=None):
            if picker_window and picker_window.winfo_exists():
                print("Picker cancelled (ESC).")
                self.is_picker_active = False
                picker_window.destroy()

        def on_close():
            if picker_window and picker_window.winfo_exists():
                print("Picker closed (WM_DELETE_WINDOW).")
                self.is_picker_active = False
                picker_window.destroy()

        try:
            picker_window.bind("<Escape>", on_esc)
            picker_window.protocol("WM_DELETE_WINDOW", on_close)
        except tk.TclError as bind_err:
            print(f"CRITICAL PICKER ERROR: Failed to bind events directly to Toplevel: {bind_err}")
            messagebox.showerror("Picker Critical Error", f"Failed basic picker window setup:\n{bind_err}")
            self.is_picker_active = False
            if picker_window and picker_window.winfo_exists():
                try: picker_window.destroy()
                except tk.TclError: pass
            return

        frame = ttk.Frame(picker_window); frame.pack(fill=tk.BOTH, expand=True)
        canvas = tk.Canvas(frame, bg="black", cursor="crosshair")
        v_scroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=canvas.yview)
        h_scroll = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=canvas.xview)
        canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        canvas.grid(row=0, column=0, sticky='nsew'); v_scroll.grid(row=0, column=1, sticky='ns'); h_scroll.grid(row=1, column=0, sticky='ew')
        frame.grid_rowconfigure(0, weight=1); frame.grid_columnconfigure(0, weight=1)

        try:
            img_tk = ImageTk.PhotoImage(screenshot_pil)
            canvas.create_image(0, 0, anchor='nw', image=img_tk)
            canvas.image = img_tk
            canvas.configure(scrollregion=canvas.bbox('all'))
        except Exception as e:
            print(f"Picker Error (Display): {e}")
            messagebox.showerror("Picker Error", f"Display error:\n{e}")
            if picker_window.winfo_exists(): picker_window.destroy()
            return

        panning = {'active': False, 'x': 0, 'y': 0}

        def on_lclick(e):
            if not canvas.winfo_exists(): return
            cx, cy = canvas.canvasx(e.x), canvas.canvasy(e.y)
            ix, iy = int(min(max(cx, 0), img_width - 1)), int(min(max(cy, 0), img_height - 1))
            try:
                rgb_tuple = tuple(screenshot_np_rgb[iy, ix])
                print(f"Pixel selected: ({ix}, {iy}) -> RGB: {rgb_tuple}")
            except IndexError:
                print("Warning: Click processed slightly outside image bounds.")
                return
            except Exception as err:
                print(f"Pixel get error: {err}")
                return

            saved_rgb = rgb_tuple
            print("Picker finished, scheduling color update.")
            self.is_picker_active = False

            if picker_window.winfo_exists():
                picker_window.destroy()

            self.root.after(10, lambda lw=label_widget, ew=entry_widget, cvn=color_var_name, sr=saved_rgb:
                            self._finalize_color_pick(lw, ew, cvn, sr))

        def on_mpress(e):
            if not canvas.winfo_exists(): return
            panning['active'], panning['x'], panning['y'] = True, e.x, e.y; canvas.config(cursor="fleur")

        def on_mmotion(e):
            if not canvas.winfo_exists(): return
            if panning['active']:
                dx, dy = e.x - panning['x'], e.y - panning['y']
                canvas.xview_scroll(int(-dx * PAN_SPEED_FACTOR), "units")
                canvas.yview_scroll(int(-dy * PAN_SPEED_FACTOR), "units")
                panning['x'], panning['y'] = e.x, e.y

        def on_mrelease(e):
            if not canvas.winfo_exists(): return
            panning['active'] = False; canvas.config(cursor="crosshair")

        def bind_canvas_events():
            if canvas.winfo_exists():
                canvas.bind("<Button-1>", on_lclick)
                canvas.bind("<ButtonPress-2>", on_mpress); canvas.bind("<ButtonPress-3>", on_mpress)
                canvas.bind("<B2-Motion>", on_mmotion); canvas.bind("<B3-Motion>", on_mmotion)
                canvas.bind("<ButtonRelease-2>", on_mrelease); canvas.bind("<ButtonRelease-3>", on_mrelease)
                print("Picker canvas events bound.")
            else:
                print("Warning: Picker canvas destroyed before events could be bound.")

        if canvas.winfo_exists():
            canvas.after(50, bind_canvas_events)
        else:
            print("CRITICAL PICKER ERROR: Canvas destroyed before event binding could be scheduled.")
            return

        picker_window.focus_force()

    def _finalize_color_pick(self, label_widget, entry_widget, color_var_name, rgb_tuple):
        """Helper function called via root.after to set the color after picker closes."""
        print(f"Finalizing pick for {color_var_name} with RGB: {rgb_tuple}")
        try:
            if entry_widget.winfo_exists() and label_widget.winfo_exists():
                entry_widget.delete(0, tk.END)
                entry_widget.insert(0, f"{rgb_tuple[0]}, {rgb_tuple[1]}, {rgb_tuple[2]}")
                self._set_color_variable(label_widget, entry_widget, color_var_name)
                print(f"Color {color_var_name} successfully updated.")
            else:
                print("Warning: Target widgets for color pick destroyed before finalization.")
        except Exception as e:
            print(f"Error during final color pick finalization: {e}")
            traceback.print_exc()


    def _processing_loop(self):
        sct_worker = None
        try:
            sct_worker = mss.mss(); print("Worker thread MSS instance created.")
            monitors = sct_worker.monitors
            primary_monitor = monitors[1] if len(monitors) > 1 else monitors[0]
            screen_width, screen_height = primary_monitor["width"], primary_monitor["height"]
            if screen_width <= 0 or screen_height <= 0:
                print("Error: Invalid screen dimensions detected, using pyautogui fallback"); screen_width, screen_height = pyautogui.size()

            while not self.stop_event.is_set():
                start_time = time.perf_counter()
                with self.config_lock:
                    paused = self.is_paused
                    do_click = self.clicking_enabled
                    area_tol = self.area_tolerance
                    box_size = self.capture_box_size
                    clk_method = self.click_method
                    c1_lab, c2_lab = self.color1_lab, self.color2_lab
                    c1_rgb, c2_rgb = self.color1_rgb, self.color2_rgb
                    clk_button = self.click_button
                    colors_tab_open = self.detected_colors_tab_active

                if paused or self.is_picker_active:
                    time.sleep(0.1); continue

                try:
                    mouse_x, mouse_y = pyautogui.position()
                    center_rgb_detected, click_decision = None, False

                    if do_click and not colors_tab_open:
                        is_mouse_over_window, is_app_focused = False, False
                        try:
                            active_window = gw.getActiveWindow()
                            if active_window and active_window.title == self.root.title():
                                is_app_focused = True
                                win_x, win_y, win_w, win_h = active_window.left, active_window.top, active_window.width, active_window.height
                                if win_x <= mouse_x < win_x + win_w and win_y <= mouse_y < win_y + win_h:
                                    is_mouse_over_window = True
                        except Exception: pass

                        can_check_center_and_click = not is_mouse_over_window

                        if can_check_center_and_click:
                            try:
                                center_sct_img = sct_worker.grab({"left": mouse_x, "top": mouse_y, "width": 1, "height": 1})
                                center_np_rgb = cv2.cvtColor(np.array(center_sct_img), cv2.COLOR_BGRA2RGB)
                                if center_np_rgb.size > 0:
                                    center_rgb_detected = tuple(center_np_rgb[0, 0])
                                    match1, match2 = False, False
                                    if clk_method == "LAB" and c1_lab and c2_lab:
                                        center_lab_np = np.array(self._rgb_to_lab(center_rgb_detected), dtype=np.float32)
                                        delta_e1 = np.sqrt(np.sum((center_lab_np - np.array(c1_lab, dtype=np.float32))**2))
                                        delta_e2 = np.sqrt(np.sum((center_lab_np - np.array(c2_lab, dtype=np.float32))**2))
                                        match1, match2 = delta_e1 <= area_tol, delta_e2 <= area_tol
                                    elif clk_method == "RGB" and c1_rgb and c2_rgb:
                                        match1 = self._colors_within_rgb_tolerance(center_rgb_detected, c1_rgb, area_tol)
                                        match2 = self._colors_within_rgb_tolerance(center_rgb_detected, c2_rgb, area_tol)
                                    if match1 or match2: click_decision = True
                            except mss.ScreenShotError: pass
                            except Exception as e: print(f"Worker Error (Center Grab): {e}"); center_rgb_detected = None

                    self.update_queue.put({"type": "center_pixel", "rgb": center_rgb_detected})

                    current_time = time.monotonic()
                    if click_decision and (current_time - self.last_click_time > CLICK_COOLDOWN):
                        try: pydirectinput.mouseDown(button=clk_button); pydirectinput.mouseUp(button=clk_button); self.last_click_time = current_time
                        except Exception as click_err: print(f"ERROR pydirectinput click: {click_err}")

                    half_size = box_size // 2
                    left, top = max(0, mouse_x - half_size), max(0, mouse_y - half_size)
                    capture_width = min(box_size, screen_width - left); capture_height = min(box_size, screen_height - top)

                    area_results = {"c1_pct": 0.0, "c1_cnt": 0, "c2_pct": 0.0, "c2_cnt": 0}
                    capture_img_for_ui, overlay_img_for_ui, dominant_colors_for_ui = None, None, []

                    if capture_width > 0 and capture_height > 0:
                        try:
                            capture_dict = {"left": left, "top": top, "width": capture_width, "height": capture_height}
                            area_sct_img = sct_worker.grab(capture_dict)
                            area_np_rgb = cv2.cvtColor(np.array(area_sct_img), cv2.COLOR_BGRA2RGB)

                            if area_np_rgb.size > 0:
                                capture_img_for_ui = area_np_rgb.copy()

                                if c1_lab is not None and c2_lab is not None:
                                    c1_lab_np = np.array(c1_lab, dtype=np.float32)
                                    c2_lab_np = np.array(c2_lab, dtype=np.float32)
                                    c1_pct, c1_cnt, mask1 = self._calculate_color_stats_lab(area_np_rgb, c1_lab_np, area_tol)
                                    c2_pct, c2_cnt, mask2 = self._calculate_color_stats_lab(area_np_rgb, c2_lab_np, area_tol)
                                    area_results = {"c1_pct": c1_pct, "c1_cnt": c1_cnt, "c2_pct": c2_pct, "c2_cnt": c2_cnt}
                                    if mask1 is not None and mask2 is not None:
                                        overlay = area_np_rgb.copy(); overlay[mask1 == 255] = (255, 0, 0); overlay[mask2 == 255] = (0, 0, 255)
                                        overlay_img_for_ui = overlay
                                else: area_results = {"c1_pct": -1.0, "c1_cnt": -1, "c2_pct": -1.0, "c2_cnt": -1}

                                if COLOR_ANALYSIS_ENABLED and colors_tab_open:
                                    dominant_colors_for_ui = self._analyze_dominant_colors(area_np_rgb)

                        except mss.ScreenShotError: pass
                        except Exception as e:
                            print(f"Worker Error (Area Grab/Calc): {e}"); traceback.print_exc()
                            area_results = {"c1_pct": -2.0, "c1_cnt": -2, "c2_pct": -2.0, "c2_cnt": -2}

                    self.update_queue.put({
                        "type": "area_update", "results": area_results,
                        "capture_img": capture_img_for_ui, "overlay_img": overlay_img_for_ui,
                        "dominant_colors": dominant_colors_for_ui
                    })

                except Exception as loop_err:
                    print(f"Critical Worker Error in loop: {loop_err}"); traceback.print_exc()
                    self.update_queue.put({"type": "error", "message": str(loop_err)}); time.sleep(0.5)

                elapsed = time.perf_counter() - start_time
                sleep_duration = max(0, WORKER_SLEEP_TIME - elapsed)
                if sleep_duration > 0:
                    time.sleep(sleep_duration)

        except Exception as thread_init_error:
            print(f"Critical worker init error: {thread_init_error}"); traceback.print_exc()
            self.update_queue.put({"type": "error", "message": f"Worker failed init: {thread_init_error}"})
        finally:
            if sct_worker: sct_worker.close(); print("Worker MSS closed.")
            print("Processing loop stopped.")

    def _colors_within_rgb_tolerance(self, rgb1, rgb2, tolerance):
        if rgb1 is None or rgb2 is None: return False
        try: return all(abs(int(c1) - int(c2)) <= tolerance for c1, c2 in zip(rgb1, rgb2))
        except: return False

    def _calculate_color_stats_lab(self, np_image_rgb, target_lab_np, tolerance):
        if np_image_rgb is None or np_image_rgb.size == 0 or target_lab_np is None: return 0, 0, None
        try:
            if len(np_image_rgb.shape) != 3 or np_image_rgb.shape[2] != 3: return 0, 0, None
            image_lab = cv2.cvtColor(np_image_rgb, cv2.COLOR_RGB2LAB)
            delta_lab_sq = np.sum((image_lab.astype(np.float32) - target_lab_np.astype(np.float32))**2, axis=2)
            mask = (delta_lab_sq <= float(tolerance)**2).astype(np.uint8) * 255
            count = cv2.countNonZero(mask)
            total = np_image_rgb.shape[0] * np_image_rgb.shape[1]
            return (count / total) * 100 if total > 0 else 0, count, mask
        except cv2.error as cv_err: print(f"OpenCV Error (LAB stats): {cv_err}"); return 0, 0, None
        except Exception as e: print(f"Error (LAB stats): {e}"); traceback.print_exc(); return 0, 0, None

    def _analyze_dominant_colors(self, np_image_rgb):
        if np_image_rgb is None or np_image_rgb.size == 0: return []
        try:
            h, w = np_image_rgb.shape[:2]
            scale = min(1.0, COLOR_ANALYSIS_RESIZE_WIDTH / w) if w > 0 else 1.0
            resized_img = cv2.resize(np_image_rgb, None, fx=scale, fy=scale, interpolation=cv2.INTER_LINEAR) if scale < 1.0 else np_image_rgb
            lab_img = cv2.cvtColor(resized_img, cv2.COLOR_RGB2LAB)
            pixels = lab_img.reshape(-1, 3).astype(np.float32)
            if len(pixels) < COLOR_ANALYSIS_K: return []

            criteria = (cv2.TERM_CRITERIA_EPS + cv2.TERM_CRITERIA_MAX_ITER, 10, 1.0)
            compactness, labels, centers = cv2.kmeans(pixels, COLOR_ANALYSIS_K, None, criteria, 10, cv2.KMEANS_PP_CENTERS)

            label_counts = collections.Counter(labels.flatten()); total_pixels = len(pixels)
            dominant_colors = []
            for i, center_lab in enumerate(centers):
                percentage = (label_counts[i] / total_pixels) * 100
                if percentage >= COLOR_ANALYSIS_MIN_PERCENT:
                    lab_tuple = tuple(map(int, center_lab))
                    dominant_colors.append({'rgb': self._lab_to_rgb(lab_tuple), 'lab': lab_tuple, 'percentage': percentage})

            return sorted(dominant_colors, key=lambda x: x['percentage'], reverse=True)
        except cv2.error as cv_err: print(f"OpenCV Error (Dominant Colors): {cv_err}"); return []
        except Exception as e: print(f"Error (Dominant Colors): {e}"); traceback.print_exc(); return []


    def _check_queue(self):
        try:
            while not self.update_queue.empty():
                self._process_queue_message(self.update_queue.get_nowait())
        except queue.Empty: pass
        except Exception as e: print(f"Error processing queue: {e}"); traceback.print_exc()
        finally:
            if self.root.winfo_exists(): self.root.after(UI_UPDATE_INTERVAL, self._check_queue)

    def _process_queue_message(self, message):
        msg_type = message.get("type")
        try:
            if not self.root.winfo_exists(): return

            if msg_type == "center_pixel":
                rgb = message.get("rgb")
                if rgb and hasattr(self, 'center_pixel_swatch') and self.center_pixel_swatch.winfo_exists():
                    hex_color = f'#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}'
                    self.center_pixel_rgb_value.set(f"{rgb[0]},{rgb[1]},{rgb[2]}")
                    self.center_pixel_swatch.config(bg=hex_color)
                elif hasattr(self, 'center_pixel_swatch') and self.center_pixel_swatch.winfo_exists():
                    self.center_pixel_rgb_value.set("N/A"); self.center_pixel_swatch.config(bg="black")

            elif msg_type == "area_update":
                results = message.get("results", {})
                c1p, c1c = results.get('c1_pct', 0), results.get('c1_cnt', 0)
                c2p, c2c = results.get('c2_pct', 0), results.get('c2_cnt', 0)
                stat_str = lambda p, c: f"Area: {p:.2f}% / {c} px"
                if c1p == -1.0: self.color1_value.set("Area: Invalid Target"); self.color2_value.set("Area: Invalid Target")
                elif c1p == -2.0: self.color1_value.set("Area: Error"); self.color2_value.set("Area: Error")
                else: self.color1_value.set(stat_str(c1p, c1c)); self.color2_value.set(stat_str(c2p, c2c))

                self._update_image_preview(self.settings_capture_preview_label, message.get("capture_img"), self.preview_size)
                self._update_image_preview(self.settings_overlay_preview_label, message.get("overlay_img"), self.preview_size)
                self._update_image_preview(self.colors_capture_preview_label, message.get("capture_img"), self.colors_preview_size)
                self._update_image_preview(self.colors_overlay_preview_label, message.get("overlay_img"), self.colors_preview_size)

                dominant_colors = message.get("dominant_colors", [])
                if dominant_colors:
                     self._update_color_treeview(dominant_colors)

            elif msg_type == "error":
                print(f"Worker Thread Error Reported: {message.get('message', 'Unknown worker error')}")
                self.status_var.set("Worker Error"); self.status_value_label.config(foreground="orange")

        except tk.TclError as e: print(f"TclError updating UI (Window closed?): {e}")
        except Exception as e: print(f"Error updating UI from queue: {e}"); traceback.print_exc()

    def _update_image_preview(self, label_widget, np_image, target_size):
        if not hasattr(label_widget, 'winfo_exists') or not label_widget.winfo_exists(): return
        if np_image is None or np_image.size == 0:
            label_widget.config(image='', width=target_size[0], height=target_size[1]); label_widget.image = None; return
        try:
            img_pil = Image.fromarray(np_image).copy()
            img_pil.thumbnail(target_size, Image.Resampling.LANCZOS)
            img_tk = ImageTk.PhotoImage(img_pil)
            label_widget.config(image=img_tk, width=img_tk.width(), height=img_tk.height()); label_widget.image = img_tk
        except tk.TclError: pass
        except Exception as e: print(f"Error updating preview: {e}"); label_widget.config(image=''); label_widget.image = None

    def _update_color_treeview(self, dominant_colors):
        if not hasattr(self, 'color_tree') or not self.color_tree.winfo_exists(): return
        try:
            self.color_tree.delete(*self.color_tree.get_children())
            for i, color_data in enumerate(dominant_colors):
                rgb, lab, percent = color_data['rgb'], color_data['lab'], color_data['percentage']
                hex_color = f'#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}'
                tag_name = f"color_{i}"
                luminance = (0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2]) / 255
                text_color = "black" if luminance > 0.5 else "white"
                self.color_tree.tag_configure(tag_name, background=hex_color, foreground=text_color)

                self.color_tree.insert('', tk.END, values=(
                    '███',
                    f"{rgb[0]}, {rgb[1]}, {rgb[2]}",
                    f"{lab[0]}, {lab[1]}, {lab[2]}",
                    f"{percent:.1f}%"
                ), tags=(tag_name,))
        except tk.TclError: pass
        except Exception as e: print(f"Error updating color treeview: {e}"); traceback.print_exc()

    def _update_mouse_pixel_info(self):
        """Polls mouse position and updates the pixel finder UI elements IF the tab is active."""
        is_active = False
        if hasattr(self, 'notebook') and self.notebook.winfo_exists():
            try:
                selected_tab_index = self.notebook.index(self.notebook.select())
                is_active = (selected_tab_index == 1)
            except tk.TclError: pass

        if is_active and not self.is_picker_active:
            try:
                x, y = pyautogui.position()
                self.mouse_coords_var.set(f"X: {x}, Y: {y}")

                pixel_dict = {"left": x, "top": y, "width": 1, "height": 1}
                sct_img = self.sct_main.grab(pixel_dict)
                np_rgb = cv2.cvtColor(np.array(sct_img), cv2.COLOR_BGRA2RGB)

                if np_rgb.size > 0:
                    rgb = tuple(np_rgb[0, 0])
                    lab = self._rgb_to_lab(rgb)
                    hex_color = f'#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}'

                    self.mouse_rgb_var.set(f"{rgb[0]}, {rgb[1]}, {rgb[2]}")
                    self.mouse_lab_var.set(f"{lab[0]}, {lab[1]}, {lab[2]}")
                    if hasattr(self, 'mouse_swatch_label') and self.mouse_swatch_label.winfo_exists():
                         self.mouse_swatch_label.config(bg=hex_color)
                else: raise Exception("Grabbed empty image")

            except (mss.ScreenShotError, pyautogui.FailSafeException):
                self.mouse_rgb_var.set("---, ---, ---"); self.mouse_lab_var.set("---, ---, ---")
                if hasattr(self, 'mouse_swatch_label') and self.mouse_swatch_label.winfo_exists(): self.mouse_swatch_label.config(bg="black")
            except Exception as e:
                self.mouse_rgb_var.set("Error"); self.mouse_lab_var.set("Error")
                if hasattr(self, 'mouse_swatch_label') and self.mouse_swatch_label.winfo_exists(): self.mouse_swatch_label.config(bg="red")
        else:
             if self.mouse_pixel_polling_active:
                 self.mouse_coords_var.set("X: ---, Y: ---")
                 self.mouse_rgb_var.set("---, ---, ---"); self.mouse_lab_var.set("---, ---, ---")
                 if hasattr(self, 'mouse_swatch_label') and self.mouse_swatch_label.winfo_exists(): self.mouse_swatch_label.config(bg="black")
                 self.mouse_pixel_polling_active = False

        if not is_active: self.mouse_pixel_polling_active = True

        if self.root.winfo_exists() and not self.stop_event.is_set():
            self.root.after(MOUSE_PIXEL_POLL_INTERVAL, self._update_mouse_pixel_info)


    def _toggle_script(self):
        with self.config_lock: self.is_paused = not self.is_paused
        self._update_status_ui()
        print(f"Script {'Paused' if self.is_paused else 'Running'}")

    def _toggle_clicking_callback(self):
        """ Toggles clicking enabled state via checkbutton or hotkey. """
        new_state = False
        with self.config_lock:
            self.clicking_enabled = not self.clicking_enabled
            new_state = self.clicking_enabled

        if hasattr(self, 'clicking_enabled_var'):
             try:
                 self.clicking_enabled_var.set(new_state)
             except tk.TclError:
                 print("Warning: Could not update clicking_enabled_var (widget destroyed?)")

        print(f"Clicking {'Enabled' if new_state else 'Disabled'}")


    def _on_tab_changed(self, event):
        """Update the flag indicating if the 'Detected Colors' tab is active."""
        try:
            selected_tab_index = self.notebook.index(self.notebook.select())
            is_colors_tab = (selected_tab_index == 1)
            with self.config_lock: self.detected_colors_tab_active = is_colors_tab
        except tk.TclError: pass
        except Exception as e: print(f"Error checking active tab: {e}")

    def _update_status_ui(self):
         status_text = "Paused" if self.is_paused else "Running"
         status_color = "red" if self.is_paused else "lime green"
         try:
            if self.root.winfo_exists():
                self.status_var.set(status_text); self.status_value_label.config(foreground=status_color)
         except tk.TclError: pass


    def _set_hotkey_listener(self, target_attribute, target_label_widget):
        if self.is_listening_for_hotkey: return
        print(f"Setting hotkey for: {target_attribute}")
        self.is_listening_for_hotkey = True
        self.hotkey_target_attr = target_attribute; self.hotkey_target_widget = target_label_widget
        self.original_label_text = target_label_widget.cget("text")
        target_label_widget.config(text="Press keys...", foreground="yellow")
        self.pause_hotkey_button.config(state=tk.DISABLED); self.toggle_click_hotkey_button.config(state=tk.DISABLED)
        self.key_listener_hook = keyboard.hook(self._on_key_press_for_hotkey, suppress=True)

    def _on_key_press_for_hotkey(self, event):
        if not self.is_listening_for_hotkey or event.event_type != keyboard.KEY_DOWN: return True
        try:
            new_hotkey_str = keyboard.get_hotkey_name()
            print(f"Captured hotkey: {new_hotkey_str}")
            self._stop_hotkey_listener()

            if not new_hotkey_str or keyboard.is_modifier(event.name) or not any(k not in keyboard.all_modifiers for k in keyboard.parse_hotkey(new_hotkey_str)):
                 print("Invalid hotkey (modifier only, empty, or no regular key). Reverting."); self._revert_hotkey_ui()
                 return True

            old_hotkey_str = getattr(self, self.hotkey_target_attr, None)
            setattr(self, self.hotkey_target_attr, new_hotkey_str)
            if self.hotkey_target_widget: self.hotkey_target_widget.config(text=new_hotkey_str, foreground="white")
            print(f"Updating {self.hotkey_target_attr}: Old='{old_hotkey_str}', New='{new_hotkey_str}'")
            self._reregister_hotkey(old_hotkey_str, new_hotkey_str, self._get_callback_for_attr(self.hotkey_target_attr))
        except Exception as e: print(f"Error processing hotkey: {e}"); traceback.print_exc(); self._stop_hotkey_listener(); self._revert_hotkey_ui()
        finally: self._enable_hotkey_buttons()
        return True

    def _stop_hotkey_listener(self):
        if self.key_listener_hook: keyboard.unhook(self.key_listener_hook); self.key_listener_hook = None; print("Hotkey listener stopped.")
        self.is_listening_for_hotkey = False; self.hotkey_target_attr = None; self.hotkey_target_widget = None

    def _revert_hotkey_ui(self):
         if self.hotkey_target_widget: self.hotkey_target_widget.config(text=self.original_label_text, foreground="white")

    def _enable_hotkey_buttons(self):
        try:
            if hasattr(self, 'pause_hotkey_button'): self.pause_hotkey_button.config(state=tk.NORMAL)
            if hasattr(self, 'toggle_click_hotkey_button'): self.toggle_click_hotkey_button.config(state=tk.NORMAL)
        except tk.TclError: pass

    def _get_callback_for_attr(self, attr_name):
        return self._toggle_script if attr_name == 'pause_hotkey_str' else (self._toggle_clicking_callback if attr_name == 'toggle_click_hotkey_str' else None)

    def _register_hotkeys(self):
        self._unregister_hotkeys(); print("Registering hotkeys...")
        pause_cb, toggle_cb = self._get_callback_for_attr('pause_hotkey_str'), self._get_callback_for_attr('toggle_click_hotkey_str')
        if pause_cb and self.pause_hotkey_str: self._register_single_hotkey(self.pause_hotkey_str, pause_cb)
        if toggle_cb and self.toggle_click_hotkey_str: self._register_single_hotkey(self.toggle_click_hotkey_str, toggle_cb)

    def _unregister_hotkeys(self):
        if not self.active_hotkeys: return
        print(f"Unregistering {len(self.active_hotkeys)} hotkeys: {list(self.active_hotkeys.keys())}")
        for combo in list(self.active_hotkeys.keys()):
             try: keyboard.remove_hotkey(combo); del self.active_hotkeys[combo]
             except Exception as e: print(f"Warn: Unregister hotkey '{combo}' failed: {e}")
        self.active_hotkeys.clear()

    def _register_single_hotkey(self, combo, callback):
        if not combo or not callback: return False
        try:
            keyboard.add_hotkey(combo, callback, trigger_on_release=False); self.active_hotkeys[combo] = callback
            print(f"Registered hotkey: '{combo}'"); return True
        except (ValueError, Exception) as e:
             perm_msg = "\n(Might require admin/root privileges)" if "permission" in str(e).lower() else ""
             print(f"FAIL register hotkey '{combo}': {e}{perm_msg}")
             messagebox.showerror("Hotkey Error", f"Could not register '{combo}'.\nError: {e}{perm_msg}")
        return False

    def _reregister_hotkey(self, old_combo, new_combo, callback):
        if old_combo and old_combo in self.active_hotkeys:
             try: keyboard.remove_hotkey(old_combo); del self.active_hotkeys[old_combo]; print(f"Unregistered old: '{old_combo}'")
             except Exception as e: print(f"Warn: Unregister old hotkey '{old_combo}' failed: {e}")
        if new_combo and callback: self._register_single_hotkey(new_combo, callback)


    def _start_worker(self):
        if self.worker_thread is None or not self.worker_thread.is_alive():
            self.stop_event.clear()
            self.worker_thread = threading.Thread(target=self._processing_loop, daemon=True)
            self.worker_thread.start(); print("Worker thread started.")

    def _save_config(self):
        with self.config_lock:
            config = {k: getattr(self, k) for k in [
                "color1_rgb", "color2_rgb", "area_tolerance", "capture_box_size",
                "click_method", "click_button", "clicking_enabled",
                "pause_hotkey_str", "toggle_click_hotkey_str" ]}
            config["pause_hotkey"] = config.pop("pause_hotkey_str")
            config["toggle_click_hotkey"] = config.pop("toggle_click_hotkey_str")
        try:
            with open(CONFIG_FILE, 'w') as f: json.dump(config, f, indent=4)
            print("Configuration saved.")
        except Exception as e: print(f"Error saving config: {e}"); messagebox.showwarning("Config Error", f"Could not save:\n{e}")

    def _load_config(self):
        defaults = {
            "color1_rgb": (204, 204, 204), "color2_rgb": (38, 120, 122),
            "area_tolerance": DEFAULT_AREA_TOLERANCE, "capture_box_size": DEFAULT_CAPTURE_BOX_SIZE,
            "click_method": DEFAULT_CLICK_METHOD, "click_button": DEFAULT_CLICK_BUTTON,
            "clicking_enabled": DEFAULT_CLICKING_ENABLED,
            "pause_hotkey": DEFAULT_PAUSE_HOTKEY, "toggle_click_hotkey": DEFAULT_TOGGLE_CLICK_HOTKEY }
        cfg = defaults.copy()
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r') as f: loaded = json.load(f)
                cfg.update(loaded); print("Config loaded.")
            except Exception as e: print(f"Config load error ({e}), using defaults."); messagebox.showwarning("Config Load Warning", f"Error reading config:\n{e}\nUsing defaults.")
        else: print("Config file not found, using defaults.")

        with self.config_lock:
            self.color1_rgb = tuple(cfg["color1_rgb"]) if isinstance(cfg["color1_rgb"], list) else defaults["color1_rgb"]
            self.color2_rgb = tuple(cfg["color2_rgb"]) if isinstance(cfg["color2_rgb"], list) else defaults["color2_rgb"]
            self.area_tolerance = int(cfg["area_tolerance"])
            self.capture_box_size = int(cfg["capture_box_size"])
            self.click_method = cfg["click_method"]
            self.click_button = cfg["click_button"]
            self.clicking_enabled = bool(cfg["clicking_enabled"])
            self.pause_hotkey_str = cfg["pause_hotkey"]
            self.toggle_click_hotkey_str = cfg["toggle_click_hotkey"]
            self.color1_lab = self._rgb_to_lab(self.color1_rgb)
            self.color2_lab = self._rgb_to_lab(self.color2_rgb)

        print(f"Initial Settings: Click={self.clicking_enabled}, Pause='{self.pause_hotkey_str}', ToggleClick='{self.toggle_click_hotkey_str}'")
        self._update_ui_from_config()
        self._register_hotkeys()

    def _update_ui_from_config(self):
         if not self.root.winfo_exists(): return
         try:
            self.color1_rgb_entry.delete(0, tk.END); self.color1_rgb_entry.insert(0, f"{self.color1_rgb[0]},{self.color1_rgb[1]},{self.color1_rgb[2]}")
            self._update_color_label_bg(self.color1_label, self.color1_rgb); self.color1_label.config(text=f"{self.color1_rgb[0]},{self.color1_rgb[1]},{self.color1_rgb[2]}")
            self.color2_rgb_entry.delete(0, tk.END); self.color2_rgb_entry.insert(0, f"{self.color2_rgb[0]},{self.color2_rgb[1]},{self.color2_rgb[2]}")
            self._update_color_label_bg(self.color2_label, self.color2_rgb); self.color2_label.config(text=f"{self.color2_rgb[0]},{self.color2_rgb[1]},{self.color2_rgb[2]}")

            self.area_tolerance_var.set(self.area_tolerance); self.area_tolerance_value_label.config(text=str(self.area_tolerance))
            self.box_size_var.set(self.capture_box_size); self.box_size_value_label.config(text=str(self.capture_box_size))

            self.clicking_enabled_var.set(self.clicking_enabled)
            self.click_method_var.set(self.click_method); self.click_button_var.set(self.click_button)

            self.pause_hotkey_label.config(text=self.pause_hotkey_str)
            self.toggle_click_hotkey_label.config(text=self.toggle_click_hotkey_str)

            self._update_status_ui()
         except Exception as e: print(f"Warning: UI update from config failed: {e}"); traceback.print_exc()

    def _on_closing(self):
        print("Closing application...")
        if self.is_listening_for_hotkey: self._stop_hotkey_listener()
        elif messagebox.askyesno("Save Config?", "Save current settings before closing?"): self._save_config()

        self.stop_event.set()
        self._unregister_hotkeys()

        if self.worker_thread and self.worker_thread.is_alive():
            print("Waiting for worker thread..."); self.worker_thread.join(timeout=0.5)
            if self.worker_thread.is_alive(): print("Worker thread join timed out.")

        if self.sct_main: 
            try: self.sct_main.close(); print("Main MSS closed.") 
            except Exception as e: print(f"Error closing main MSS: {e}")


        self.root.destroy()
        print("Application closed.")

if __name__ == "__main__":
    try:
        try: from ctypes import windll; windll.shcore.SetProcessDpiAwareness(1)
        except Exception: pass

        root = tk.Tk()
        app = ColorBotApp(root)
        root.mainloop()
    except Exception as main_err:
         print("\n--- UNHANDLED ERROR IN MAIN EXECUTION ---")
         traceback.print_exc()
         print("------------------------------------------")
         try: messagebox.showerror("Fatal Error", f"A critical error occurred:\n\n{main_err}\n\nSee console for details.")
         except: pass
         input("Press Enter to exit...")