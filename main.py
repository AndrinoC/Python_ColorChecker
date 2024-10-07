import tkinter as tk
from PIL import Image, ImageTk, ImageGrab
import cv2
import numpy as np
import os
import json
import keyboard

CONFIG_FILE = "config.json"
COLOR_TOLERANCE = 15
PROCESS_DELAY = 50

color1 = (204, 204, 204)
color2 = (38, 120, 122)
is_paused = True

root = tk.Tk()
root.title("Color detection")
root.geometry("650x300")
root.resizable(False, False)
root.configure(bg='black')

status_frame = tk.Frame(root, bg="#333333")
status_frame.place(relx=0.25, rely=0.05, anchor=tk.N)

main_frame = tk.Frame(root, bg="#333333")
main_frame.place(relx=0.36, rely=0.2, anchor=tk.N)

grayscale_preview_label = tk.Label(root)
grayscale_preview_label.place(relx=0.82, rely=0.25, anchor=tk.CENTER)

preview_label = tk.Label(root)
preview_label.place(relx=0.82, rely=0.7, anchor=tk.CENTER)

status_label = tk.Label(status_frame, text="Bot Status:", bg="#333333", fg="white", font=("Arial", 12))
status_label.grid(row=0, column=0, padx=5, pady=2)

status = tk.StringVar(value="Paused")
status_value_label = tk.Label(status_frame, textvariable=status, bg="#333333", fg="white", font=("Arial", 12))
status_value_label.grid(row=0, column=1, padx=5, pady=2)

color1_label = tk.Label(main_frame, text="Color 1 (% / px)", bg=f'#{color1[0]:02x}{color1[1]:02x}{color1[2]:02x}', fg="white", font=("Arial", 12))
color1_label.grid(row=0, column=0, padx=5, pady=2)

color1_rgb_entry = tk.Entry(main_frame, width=12)
color1_rgb_entry.grid(row=0, column=1, padx=5, pady=2)
color1_rgb_entry.insert(0, f"{color1[0]}, {color1[1]}, {color1[2]}")

color1_value = tk.StringVar(value="0% / 0 px")
color1_value_label = tk.Label(main_frame, textvariable=color1_value, bg="#333333", fg="white", font=("Arial", 12))
color1_value_label.grid(row=0, column=2, padx=5, pady=2)

color2_label = tk.Label(main_frame, text="Color 2 (% / px)", bg=f'#{color2[0]:02x}{color2[1]:02x}{color2[2]:02x}', fg="white", font=("Arial", 12))
color2_label.grid(row=1, column=0, padx=5, pady=2)

color2_rgb_entry = tk.Entry(main_frame, width=12)
color2_rgb_entry.grid(row=1, column=1, padx=5, pady=2)
color2_rgb_entry.insert(0, f"{color2[0]}, {color2[1]}, {color2[2]}")

color2_value = tk.StringVar(value="0% / 0 px")
color2_value_label = tk.Label(main_frame, textvariable=color2_value, bg="#333333", fg="white", font=("Arial", 12))
color2_value_label.grid(row=1, column=2, padx=5, pady=2)

tolerance_label = tk.Label(main_frame, text="Tolerance:", bg="#333333", fg="white", font=("Arial", 12))
tolerance_label.grid(row=2, column=0, padx=5, pady=2)

tolerance_value = tk.StringVar(value=str(COLOR_TOLERANCE))
tolerance_label_value = tk.Label(main_frame, textvariable=tolerance_value, bg="#333333", fg="white", font=("Arial", 12))
tolerance_label_value.grid(row=2, column=1, padx=5, pady=2)

tolerance_slider = tk.Scale(main_frame, from_=0, to=100, orient=tk.HORIZONTAL, command=lambda val: update_tolerance(val), bg="#333333")
tolerance_slider.set(COLOR_TOLERANCE)
tolerance_slider.grid(row=3, column=0, columnspan=3, padx=5, pady=10)

toggle_label = tk.Label(main_frame, text="Press CTRL+Shift+X to pause/resume the script.", bg="#333333", fg="white", font=("Arial", 12))
toggle_label.grid(row=4, column=0, columnspan=3, padx=5, pady=10)

def is_ui_alive():
    return root.winfo_exists()

def calculate_color_statistics(np_image, color1, color2):
    total_pixels = np_image.size // 3
    if total_pixels == 0:
        return 0, 0, 0, 0

    lower_color1 = np.array([max(0, c - COLOR_TOLERANCE) for c in color1])
    upper_color1 = np.array([min(255, c + COLOR_TOLERANCE) for c in color1])
    mask_color1 = cv2.inRange(np_image, lower_color1, upper_color1)
    count_color1 = cv2.countNonZero(mask_color1)

    lower_color2 = np.array([max(0, c - COLOR_TOLERANCE) for c in color2])
    upper_color2 = np.array([min(255, c + COLOR_TOLERANCE) for c in color2])
    mask_color2 = cv2.inRange(np_image, lower_color2, upper_color2)
    count_color2 = cv2.countNonZero(mask_color2)

    color1_percentage = (count_color1 / total_pixels) * 100
    color2_percentage = (count_color2 / total_pixels) * 100

    return color1_percentage, count_color1, color2_percentage, count_color2

def set_color(color_label, rgb_color, color_var):
    global color1, color2
    
    if color_var == "color1":
        color1 = tuple(int(c) for c in rgb_color.split(','))
        color1_rgb_entry.delete(0, tk.END)
        color1_rgb_entry.insert(0, f"{color1[0]}, {color1[1]}, {color1[2]}")
    elif color_var == "color2":
        color2 = tuple(int(c) for c in rgb_color.split(','))
        color2_rgb_entry.delete(0, tk.END)
        color2_rgb_entry.insert(0, f"{color2[0]}, {color2[1]}, {color2[2]}")

    color_label.config(bg=f'#{color1[0]:02x}{color1[1]:02x}{color1[2]:02x}' if color_var == "color1" else 
                      f'#{color2[0]:02x}{color2[1]:02x}{color2[2]:02x}')
    
    save_config(color1=color1, color2=color2)

def update_color_from_entry(event, color_var):
    rgb_values = color1_rgb_entry.get() if color_var == "color1" else color2_rgb_entry.get()
    set_color(color1_label if color_var == "color1" else color2_label, rgb_values, color_var)

color1_rgb_entry.bind("<Return>", lambda event: update_color_from_entry(event, "color1"))
color2_rgb_entry.bind("<Return>", lambda event: update_color_from_entry(event, "color2"))

def update_tolerance(val):
    global COLOR_TOLERANCE
    COLOR_TOLERANCE = int(val)
    tolerance_value.set(val)
    save_config(color1=color1, color2=color2, color_tolerance=COLOR_TOLERANCE)

def pick_color_from_screen(label):
    global color1, color2

    screenshot = ImageGrab.grab()
    screenshot_np = np.array(screenshot)

    screen_window = tk.Toplevel(root)
    screen_window.title("Click a color")

    max_size = 1600
    img_width, img_height = screenshot.size
    if img_width > img_height:
        scale_factor = max_size / img_width
    else:
        scale_factor = max_size / img_height

    resized_width = int(img_width * scale_factor)
    resized_height = int(img_height * scale_factor)

    screenshot_resized = screenshot.resize((resized_width, resized_height))
    img = ImageTk.PhotoImage(screenshot_resized)

    screen_label = tk.Label(screen_window, image=img)
    screen_label.image = img
    screen_label.pack()

    def on_click(event):
        scale_x = img_width / resized_width
        scale_y = img_height / resized_height

        x = int(event.x * scale_x)
        y = int(event.y * scale_y)

        x = min(max(x, 0), img_width - 1)
        y = min(max(y, 0), img_height - 1)

        rgb_color = tuple(screenshot_np[y, x].astype(int))

        if label == color1_label:
            set_color(color1_label, f"{rgb_color[0]}, {rgb_color[1]}, {rgb_color[2]}", "color1")
        elif label == color2_label:
            set_color(color2_label, f"{rgb_color[0]}, {rgb_color[1]}, {rgb_color[2]}", "color2")

        screen_window.destroy()

    screen_label.bind("<Button-1>", on_click)

color1_label.bind("<Button-1>", lambda event: pick_color_from_screen(color1_label))
color2_label.bind("<Button-1>", lambda event: pick_color_from_screen(color2_label))

def process_screen():
    if not is_paused:
        screenshot = ImageGrab.grab(bbox=(960 - 60, 540 - 60, 960 + 60, 540 + 60))
        screenshot_np = np.array(screenshot)

        color1_percentage, count_color1, color2_percentage, count_color2 = calculate_color_statistics(screenshot_np, color1, color2)

        color1_value.set(f"{color1_percentage:.2f}% / {count_color1} px")
        color2_value.set(f"{color2_percentage:.2f}% / {count_color2} px")

        gray_image = cv2.cvtColor(screenshot_np, cv2.COLOR_BGR2GRAY)

        lower_color1 = np.array([max(0, c - COLOR_TOLERANCE) for c in color1])
        upper_color1 = np.array([min(255, c + COLOR_TOLERANCE) for c in color1])
        mask_color1 = cv2.inRange(screenshot_np, lower_color1, upper_color1)

        lower_color2 = np.array([max(0, c - COLOR_TOLERANCE) for c in color2])
        upper_color2 = np.array([min(255, c + COLOR_TOLERANCE) for c in color2])
        mask_color2 = cv2.inRange(screenshot_np, lower_color2, upper_color2)

        overlay = cv2.merge([gray_image, gray_image, gray_image])
        overlay[mask_color1 != 0] = (0, 0, 255)
        overlay[mask_color2 != 0] = (255, 0, 0)

        screenshot_with_text = Image.fromarray(screenshot_np)
        screenshot_with_text.thumbnail((400, 300))
        img = ImageTk.PhotoImage(screenshot_with_text)
        preview_label.config(image=img)
        preview_label.image = img

        gray_with_overlay_image = Image.fromarray(overlay)
        gray_with_overlay_image.thumbnail((400, 300))
        gray_img = ImageTk.PhotoImage(gray_with_overlay_image)
        grayscale_preview_label.config(image=gray_img)
        grayscale_preview_label.image = gray_img

    root.after(PROCESS_DELAY, process_screen)

def toggle_script():
    global is_paused
    is_paused = not is_paused
    status.set("Running" if not is_paused else "Paused")

keyboard.add_hotkey('ctrl+shift+x', toggle_script)

def save_config(color1=None, color2=None, color_tolerance=None):
    config = {
        "color1": color1 if color1 is not None else (204, 204, 204),
        "color2": color2 if color2 is not None else (38, 120, 122),
        "color_tolerance": color_tolerance if color_tolerance is not None else COLOR_TOLERANCE,
    }

    with open(CONFIG_FILE, 'w') as f:
        json.dump(config, f, indent=4)

def load_config():
    global color1, color2, COLOR_TOLERANCE
    default_color1 = (204, 204, 204)
    default_color2 = (38, 120, 122)
    default_color_tolerance = 15

    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r') as f:
                config = json.load(f)

            color1 = tuple(config.get("color1", default_color1))
            color2 = tuple(config.get("color2", default_color2))
            COLOR_TOLERANCE = config.get("color_tolerance", default_color_tolerance)

        except (json.JSONDecodeError, TypeError) as e:
            color1 = default_color1
            color2 = default_color2
            COLOR_TOLERANCE = default_color_tolerance
    else:
        color1 = default_color1
        color2 = default_color2
        COLOR_TOLERANCE = default_color_tolerance

    color1_label.config(bg=f'#{color1[0]:02x}{color1[1]:02x}{color1[2]:02x}')
    color2_label.config(bg=f'#{color2[0]:02x}{color2[1]:02x}{color2[2]:02x}')
    tolerance_slider.set(COLOR_TOLERANCE)

def on_closing():
    root.destroy()

root.protocol("WM_DELETE_WINDOW", on_closing)

load_config()
process_screen()
root.mainloop()
