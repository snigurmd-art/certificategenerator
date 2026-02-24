#!/usr/bin/env python
#Это свободная программа: вы можете перераспространять ее и/или изменять ее на условиях Стандартной общественной лицензии GNU в том виде, в каком она была опубликована Фондом свободного программного обеспечения; либо версии 3 лицензии, либо (по вашему выбору) любой более поздней версии.
#Эта программа распространяется в надежде, что она будет полезной, но БЕЗО ВСЯКИХ ГАРАНТИЙ; даже без неявной гарантии ТОВАРНОГО ВИДА или ПРИГОДНОСТИ ДЛЯ ОПРЕДЕЛЕННЫХ ЦЕЛЕЙ. Подробнее см. в Стандартной общественной лицензии GNU.
#Вы должны были получить копию Стандартной общественной лицензии GNU вместе с этой программой. Если это не так, см. <https://www.gnu.org/licenses/>.
#Для связи: snigurmd@yandex.ru

import os
import sys
import json
import platform
import zipfile
import subprocess
import re
import tkinter.ttk as ttk
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

try:
    from fontTools.ttLib import TTFont as FontTool
    FONTTOOLS_AVAILABLE = True
except ImportError:
    FONTTOOLS_AVAILABLE = False

import tkinter as tk
from tkinter import filedialog, messagebox, colorchooser
import customtkinter as ctk
import fitz

# === Константы ===
BASE_DPI = 150
A4_WIDTH_MM = 210
A4_HEIGHT_MM = 297

def mm_to_px(mm, dpi):
    return int(round((mm / 25.4) * dpi))

def safe_int(value, default=0):
    try:
        return int(float(value))
    except (ValueError, TypeError):
        return default

def is_valid_hex(hex_str):
    return bool(re.fullmatch(r"#[0-9A-Fa-f]{6}", hex_str))

def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def open_file_or_dir(path):
    path = os.path.abspath(path)
    system = platform.system()
    try:
        if system == "Windows":
            os.startfile(path)
        elif system == "Darwin":
            subprocess.run(["open", path], check=True)
        else:
            subprocess.run(["xdg-open", path], check=True)
    except Exception as e:
        messagebox.showwarning("⚠️ Внимание", f"Не удалось открыть: {e}")

def get_font_full_name(font_path):
    if not FONTTOOLS_AVAILABLE:
        return Path(font_path).stem.replace("_", " ").title()
    try:
        font = FontTool(font_path)
        name_records = font["name"].names
        for record in name_records:
            if record.nameID == 4 and record.platformID == 3 and record.langID == 0x409:
                return record.toStr()
        for record in name_records:
            if record.nameID == 4:
                return record.toStr()
    except Exception:
        pass
    return Path(font_path).stem.replace("_", " ").title()

def get_system_font_dirs():
    system = platform.system()
    paths = []
    if system == "Windows":
        paths = [Path(os.environ.get("WINDIR", "C:\\Windows")) / "Fonts"]
    elif system == "Darwin":
        paths = [
            Path("/System/Library/Fonts"),
            Path("/Library/Fonts"),
            Path.home() / "Library/Fonts"
        ]
    elif system == "Linux":
        paths = [
            Path.home() / ".fonts",
            Path.home() / ".local/share/fonts",
            Path("/usr/share/fonts"),
            Path("/usr/local/share/fonts")
        ]
    return [p for p in paths if p.exists()]

def scan_font_files():
    font_map = {}
    seen_names = set()
    FONTS_DIR = resource_path("fonts")
    if os.path.exists(FONTS_DIR):
        for f in os.listdir(FONTS_DIR):
            if f.lower().endswith((".ttf", ".otf")):
                path = os.path.join(FONTS_DIR, f)
                full_name = get_font_full_name(path)
                counter = 1
                display_name = full_name
                while display_name in seen_names:
                    display_name = f"{full_name} ({counter})"
                    counter += 1
                seen_names.add(display_name)
                font_map[display_name] = path

    for font_dir in get_system_font_dirs():
        try:
            for font_path in font_dir.rglob("*"):
                if font_path.suffix.lower() in (".ttf", ".otf"):
                    full_name = get_font_full_name(str(font_path))
                    counter = 1
                    display_name = full_name
                    while display_name in seen_names:
                        display_name = f"{full_name} ({counter})"
                        counter += 1
                    seen_names.add(display_name)
                    font_map[display_name] = str(font_path)
        except (PermissionError, OSError):
            continue
    return font_map

FONT_PATHS = scan_font_files()
BUILTIN_FONTS = ["Helvetica", "Helvetica-Bold", "Times-Roman", "Courier"]
FONT_CHOICES = {name: name for name in FONT_PATHS.keys()}
FONT_CHOICES.update({f: f for f in BUILTIN_FONTS})

OUTPUT_DIR = "output"
SETTINGS_FILE = "settings.json"
os.makedirs(OUTPUT_DIR, exist_ok=True)

def infer_dpi_from_a4_size(img: Image.Image) -> float:
    """Предполагает, что изображение — A4, и вычисляет DPI."""
    w_px, h_px = img.width, img.height
    # Попробуем обе ориентации и выберем ту, где DPI ближе к целому
    dpi_portrait_w = w_px * 25.4 / A4_WIDTH_MM
    dpi_portrait_h = h_px * 25.4 / A4_HEIGHT_MM
    dpi_landscape_w = w_px * 25.4 / A4_HEIGHT_MM
    dpi_landscape_h = h_px * 25.4 / A4_WIDTH_MM

    error_portrait = abs(dpi_portrait_w - dpi_portrait_h)
    error_landscape = abs(dpi_landscape_w - dpi_landscape_h)

    if error_portrait <= error_landscape:
        return (dpi_portrait_w + dpi_portrait_h) / 2
    else:
        return (dpi_landscape_w + dpi_landscape_h) / 2

def load_template_image(template_path, base_dpi=BASE_DPI):
    path = Path(template_path)
    if path.suffix.lower() in (".pdf",):
        doc = fitz.open(template_path)
        page = doc.load_page(0)
        mat = fitz.Matrix(base_dpi / 72, base_dpi / 72)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        doc.close()
        orientation = "landscape" if img.width > img.height else "portrait"
        return img, base_dpi, orientation

    elif path.suffix.lower() in (".png", ".jpg", ".jpeg"):
        img = Image.open(template_path).convert("RGB")
        inferred_dpi = infer_dpi_from_a4_size(img)
        orientation = "landscape" if img.width > img.height else "portrait"
        return img, inferred_dpi, orientation
    else:
        raise ValueError("Поддерживаются только PDF, PNG, JPG, JPEG")

def get_pil_font(font_name, size_px):
    if font_name in FONT_PATHS:
        try:
            return ImageFont.truetype(FONT_PATHS[font_name], size_px)
        except Exception:
            pass
    try:
        return ImageFont.truetype("arial.ttf", size_px)
    except:
        try:
            return ImageFont.truetype("DejaVuSans.ttf", size_px)
        except:
            return ImageFont.load_default()

def draw_preview(template_img, fields, preview_scale=1.0):
    img = template_img.copy()
    if preview_scale != 1.0:
        img = img.resize(
            (int(img.width * preview_scale), int(img.height * preview_scale)),
            Image.LANCZOS
        )
    draw = ImageDraw.Draw(img)
    for field in fields:
        try:
            x = int(field["x"])
            y = int(field["y"])
            text = str(field["text"])
            size = int(field["size"])
            align = field["align"]
            color = field.get("color", "#000000")
            if not is_valid_hex(color):
                color = "#000000"
            fill = hex_to_rgb(color)
            x_s = int(x * preview_scale)
            y_s = int(y * preview_scale)
            size_s = max(6, int(size * preview_scale))
            pil_font = get_pil_font(field.get("font", "Helvetica"), size_s)
            try:
                ascent, _ = pil_font.getmetrics()
            except Exception:
                ascent = int(size_s * 0.8)
            y_draw = y_s - ascent
            if align == "center":
                bbox = pil_font.getbbox(text)
                text_width = bbox[2] - bbox[0]
                x_draw = x_s - text_width // 2
            else:
                x_draw = x_s
            draw.text((x_draw, y_draw), text, fill=fill, font=pil_font)
            cx, cy = x_s, y_s
            draw.line([(cx - 3, cy), (cx + 3, cy)], fill=fill, width=1)
            draw.line([(cx, cy - 3), (cx, cy + 3)], fill=fill, width=1)
        except Exception as e:
            print(f"⚠️ draw_preview ошибка: {e}")
            continue
    return img

def create_certificate_png(bg_img, fields, output_path, dpi):
    draw = ImageDraw.Draw(bg_img)
    for field in fields:
        try:
            x = int(field["x"])
            y = int(field["y"])
            text = str(field["text"])
            size = int(field["size"])
            align = field["align"]
            color = field.get("color", "#000000")
            if not is_valid_hex(color):
                color = "#000000"
            fill = hex_to_rgb(color)
            pil_font = get_pil_font(field.get("font", "Helvetica"), size)
            try:
                ascent, _ = pil_font.getmetrics()
            except Exception:
                ascent = int(size * 0.8)
            y_draw = y - ascent
            if align == "center":
                bbox = pil_font.getbbox(text)
                text_width = bbox[2] - bbox[0]
                x_draw = x - text_width // 2
            else:
                x_draw = x
            draw.text((x_draw, y_draw), text, fill=fill, font=pil_font)
        except Exception as e:
            print(f"⚠️ create_certificate_png ошибка: {e}")
            continue
    bg_img.save(output_path, "PNG", dpi=(dpi, dpi))

def generate_single_png(
    template_path, font_name,
    name_txt, name_x_mm, name_y_mm, name_size_mm, name_align, name_color,
    start_num, num_x_mm, num_y_mm, num_size_mm, num_align, num_color,
    fix_txt, fix_x_mm, fix_y_mm, fix_size_mm, fix_align, fix_color
):
    if not template_path:
        return None, "❌ Загрузите шаблон"
    try:
        start_num = safe_int(start_num, 1)
    except:
        return None, "❌ Номер — целое число"

    bg_img, actual_dpi, orientation = load_template_image(template_path, base_dpi=BASE_DPI)

    fields = [
        {"text": name_txt, "x": mm_to_px(name_x_mm, actual_dpi), "y": mm_to_px(name_y_mm, actual_dpi),
         "size": mm_to_px(name_size_mm, actual_dpi), "align": name_align, "color": name_color, "font": font_name},
        {"text": f"{start_num:04d}", "x": mm_to_px(num_x_mm, actual_dpi), "y": mm_to_px(num_y_mm, actual_dpi),
         "size": mm_to_px(num_size_mm, actual_dpi), "align": num_align, "color": num_color, "font": font_name},
        {"text": fix_txt, "x": mm_to_px(fix_x_mm, actual_dpi), "y": mm_to_px(fix_y_mm, actual_dpi),
         "size": mm_to_px(fix_size_mm, actual_dpi), "align": fix_align, "color": fix_color, "font": font_name},
    ]

    out_path = os.path.join(OUTPUT_DIR, "single.png")
    try:
        create_certificate_png(bg_img, fields, out_path, dpi=actual_dpi)
        return out_path, "✅ PNG готов!"
    except Exception as e:
        return None, f"❌ Ошибка: {e}"

def generate_batch_png(
    template_path, names_file, font_name,
    start_num,
    name_x_mm, name_y_mm, name_size_mm, name_align, name_color,
    num_x_mm, num_y_mm, num_size_mm, num_align, num_color,
    fix_txt, fix_x_mm, fix_y_mm, fix_size_mm, fix_align, fix_color
):
    if not template_path or not names_file:
        return None, "❌ Загрузите шаблон и список имён"
    try:
        start_num = safe_int(start_num, 1)
        with open(names_file, encoding="utf-8") as f:
            names = [n.strip() for n in f if n.strip()]
        if not names:
            return None, "❌ Список пуст"
    except Exception as e:
        return None, f"❌ Ошибка чтения: {e}"

    bg_img_orig, actual_dpi, orientation = load_template_image(template_path, base_dpi=BASE_DPI)
    zip_path = os.path.join(OUTPUT_DIR, "batch_png.zip")

    with zipfile.ZipFile(zip_path, "w") as zf:
        for i, name in enumerate(names):
            bg_img = bg_img_orig.copy()
            fields = [
                {"text": name, "x": mm_to_px(name_x_mm, actual_dpi), "y": mm_to_px(name_y_mm, actual_dpi),
                 "size": mm_to_px(name_size_mm, actual_dpi), "align": name_align, "color": name_color, "font": font_name},
                {"text": f"{start_num + i:04d}", "x": mm_to_px(num_x_mm, actual_dpi), "y": mm_to_px(num_y_mm, actual_dpi),
                 "size": mm_to_px(num_size_mm, actual_dpi), "align": num_align, "color": num_color, "font": font_name},
                {"text": fix_txt, "x": mm_to_px(fix_x_mm, actual_dpi), "y": mm_to_px(fix_y_mm, actual_dpi),
                 "size": mm_to_px(fix_size_mm, actual_dpi), "align": fix_align, "color": fix_color, "font": font_name},
            ]
            safe_name = "".join(c if c.isalnum() or c in " _-" else "_" for c in name)
            png_name = f"cert_{safe_name}_{start_num + i:04d}.png"
            png_path = os.path.join(OUTPUT_DIR, png_name)
            create_certificate_png(bg_img, fields, png_path, dpi=actual_dpi)
            zf.write(png_path, arcname=png_name)
            os.remove(png_path)
    return zip_path, f"✅ {len(names)} PNG в ZIP"

# ========== SpinboxFrame ==========
class SpinboxFrame(ctk.CTkFrame):
    def __init__(self, parent, initial=0, min_val=0, max_val=5000, width=70, command=None, **kwargs):
        super().__init__(parent, fg_color="transparent", **kwargs)
        self.min_val = min_val
        self.max_val = max_val
        self.command = command or (lambda: None)
        self.entry = ctk.CTkEntry(self, width=width - 14, justify="center")
        self.entry.insert(0, str(int(initial)))
        self.entry.pack(side="left", padx=(0, 2))
        self.entry.bind("<KeyRelease>", lambda e: self.command())
        self.entry.bind("<FocusOut>", lambda e: self._on_focus_out())
        arrow_frame = ctk.CTkFrame(self, fg_color="transparent", width=12, height=24)
        arrow_frame.pack(side="left")
        arrow_frame.pack_propagate(False)
        self.up_arrow = ctk.CTkLabel(arrow_frame, text="▲", width=12, height=10, font=ctk.CTkFont(size=8), cursor="hand2")
        self.down_arrow = ctk.CTkLabel(arrow_frame, text="▼", width=12, height=10, font=ctk.CTkFont(size=8), cursor="hand2")
        self.up_arrow.bind("<Button-1>", lambda e: self.increment())
        self.down_arrow.bind("<Button-1>", lambda e: self.decrement())
        self.up_arrow.pack(side="top", fill="x", pady=(0, 1))
        self.down_arrow.pack(side="bottom", fill="x", pady=(1, 0))

    def get(self):
        return self.entry.get()

    def set(self, value):
        self.entry.delete(0, tk.END)
        self.entry.insert(0, str(int(value)))

    def _on_focus_out(self):
        val = self.entry.get().strip()
        if val == "":
            self.set(0)
        else:
            try:
                num = int(float(val))
                num = max(self.min_val, min(self.max_val, num))
                self.set(num)
            except ValueError:
                self.set(0)
        self.command()

    def increment(self):
        try:
            val = int(float(self.entry.get()))
            val = min(self.max_val, val + 1)
            self.set(val)
            self.command()
        except ValueError:
            self.set(0)
            self.command()

    def decrement(self):
        try:
            val = int(float(self.entry.get()))
            val = max(self.min_val, val - 1)
            self.set(val)
            self.command()
        except ValueError:
            self.set(0)
            self.command()

    def update_range(self, min_val, max_val):
        self.min_val = min_val
        self.max_val = max_val
        try:
            val = int(float(self.entry.get()))
            val = max(min_val, min(max_val, val))
            self.set(val)
        except ValueError:
            self.set(min_val)

# ========== Основное приложение ==========
class CertificateApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("🎓 Генератор сертификатов")
        self.geometry("1200x900")
        self.template_path = None
        self.names_file_path = None
        self.preview_img_tk = None
        # Кэш для быстрого предпросмотра
        self.cached_template_img = None
        self.cached_dpi = None
        self.cached_orientation = None
        self.setup_ui()
        self.load_settings_on_start()
        self.bind_mousewheel()

    def bind_mousewheel(self):
        self.preview_label.bind("<MouseWheel>", self.on_mousewheel)
        self.preview_label.bind("<Button-4>", self.on_mousewheel)
        self.preview_label.bind("<Button-5>", self.on_mousewheel)

    def on_mousewheel(self, event):
        if not self.template_path:
            return
        if event.num == 4 or event.delta > 0:
            step = +0.1
        elif event.num == 5 or event.delta < 0:
            step = -0.1
        else:
            return
        current = self.preview_scale.get()
        new_scale = round(current + step, 1)
        new_scale = max(0.3, min(3.0, new_scale))
        if abs(new_scale - current) < 0.01:
            return
        self.preview_scale.set(new_scale)
        self.update_preview()

    def update_spinbox_ranges_for_orientation(self, orientation):
        w_max = A4_HEIGHT_MM if orientation == "landscape" else A4_WIDTH_MM
        h_max = A4_WIDTH_MM if orientation == "landscape" else A4_HEIGHT_MM
        self.name_x.update_range(0, w_max)
        self.name_y.update_range(0, h_max)
        self.num_x.update_range(0, w_max)
        self.num_y.update_range(0, h_max)
        self.fix_x.update_range(0, w_max)
        self.fix_y.update_range(0, h_max)

    def setup_ui(self):
        control_frame = ctk.CTkScrollableFrame(self)
        control_frame.pack(side="left", fill="y", padx=10, pady=10, expand=False)

        ctk.CTkLabel(control_frame, text="🔤 Шрифт").pack(pady=(10, 0))
        self.font_var = ctk.StringVar(value="Helvetica")
        self.font_combobox = ttk.Combobox(
            control_frame,
            textvariable=self.font_var,
            values=sorted(FONT_CHOICES.keys()),
            state="readonly",
            height=15
        )
        self.font_combobox.bind("<<ComboboxSelected>>", lambda e: self.update_preview())
        self.font_combobox.pack(pady=5, padx=10, fill="x")

        ctk.CTkLabel(control_frame, text="🔍 Масштаб предпросмотра").pack(pady=(15, 5))
        self.preview_scale = ctk.DoubleVar(value=1.0)
        self.scale_slider = ctk.CTkSlider(
            control_frame,
            from_=0.3,
            to=3.0,
            number_of_steps=50,
            variable=self.preview_scale,
            command=lambda _: self.update_preview()
        )
        self.scale_slider.pack(pady=5, padx=10, fill="x")

        # === Поле 1: Имя ===
        ctk.CTkLabel(control_frame, text="🧑 Имя (в мм)").pack(pady=(15, 5))
        self.name_text = ctk.CTkEntry(control_frame, placeholder_text="Иванов Иван")
        self.name_text.insert(0, "Иванов Иван")
        self.name_text.pack(padx=10, fill="x")
        self.name_text.bind("<KeyRelease>", lambda e: self.update_preview())
        self.name_text.bind("<FocusOut>", lambda e: self.update_preview())

        coord_frame1 = ctk.CTkFrame(control_frame, fg_color="transparent")
        ctk.CTkLabel(coord_frame1, text="X:", width=20).pack(side="left", padx=(0, 5))
        self.name_x = SpinboxFrame(coord_frame1, initial=105, min_val=0, max_val=A4_WIDTH_MM, command=self.update_preview)
        self.name_x.pack(side="left", padx=(0, 10))
        ctk.CTkLabel(coord_frame1, text="Y:", width=20).pack(side="left", padx=(0, 5))
        self.name_y = SpinboxFrame(coord_frame1, initial=100, min_val=0, max_val=A4_HEIGHT_MM, command=self.update_preview)
        self.name_y.pack(side="left")
        coord_frame1.pack(pady=2)

        size_frame1 = ctk.CTkFrame(control_frame, fg_color="transparent")
        ctk.CTkLabel(size_frame1, text="Размер текста:", width=100).pack(side="left", padx=(0, 5))
        self.name_size = SpinboxFrame(size_frame1, initial=8, min_val=2, max_val=30, width=60, command=self.update_preview)
        self.name_size.pack(side="left")
        size_frame1.pack(pady=2)

        self.name_align = ctk.StringVar(value="center")
        align_frame1 = ctk.CTkFrame(control_frame, fg_color="transparent")
        ctk.CTkRadioButton(align_frame1, text="Слева", variable=self.name_align, value="left", command=self.update_preview).pack(side="left")
        ctk.CTkRadioButton(align_frame1, text="По центру", variable=self.name_align, value="center", command=self.update_preview).pack(side="left")
        align_frame1.pack(pady=2)

        self.name_color = "#000000"
        color_frame1 = ctk.CTkFrame(control_frame, fg_color="transparent")
        ctk.CTkButton(color_frame1, text="🎨", command=self.pick_name_color, width=36, height=24).pack(side="left", padx=(0, 5))
        self.name_color_indicator = ctk.CTkLabel(color_frame1, text="", width=20, height=20, corner_radius=4, fg_color=self.name_color)
        self.name_color_indicator.pack(side="left", padx=(0, 5))
        self.name_hex_entry = ctk.CTkEntry(color_frame1, placeholder_text="#000000", width=80)
        self.name_hex_entry.insert(0, self.name_color)
        self.name_hex_entry.pack(side="left")
        self.name_hex_entry.bind("<FocusOut>", self.on_name_hex_focus_out)
        self.name_hex_entry.bind("<Return>", self.on_name_hex_focus_out)
        color_frame1.pack(pady=5, padx=10, fill="x")

        # === Поле 2: Номер ===
        ctk.CTkLabel(control_frame, text="🔢 Номер (в мм)").pack(pady=(15, 5))
        self.num_start = ctk.CTkEntry(control_frame, placeholder_text="Начальный номер")
        self.num_start.insert(0, "1")
        self.num_start.pack(padx=10, fill="x")
        self.num_start.bind("<KeyRelease>", lambda e: self.update_preview())
        self.num_start.bind("<FocusOut>", lambda e: self.update_preview())

        coord_frame2 = ctk.CTkFrame(control_frame, fg_color="transparent")
        ctk.CTkLabel(coord_frame2, text="X:", width=20).pack(side="left", padx=(0, 5))
        self.num_x = SpinboxFrame(coord_frame2, initial=105, min_val=0, max_val=A4_WIDTH_MM, command=self.update_preview)
        self.num_x.pack(side="left", padx=(0, 10))
        ctk.CTkLabel(coord_frame2, text="Y:", width=20).pack(side="left", padx=(0, 5))
        self.num_y = SpinboxFrame(coord_frame2, initial=120, min_val=0, max_val=A4_HEIGHT_MM, command=self.update_preview)
        self.num_y.pack(side="left")
        coord_frame2.pack(pady=2)

        size_frame2 = ctk.CTkFrame(control_frame, fg_color="transparent")
        ctk.CTkLabel(size_frame2, text="Размер текста:", width=100).pack(side="left", padx=(0, 5))
        self.num_size = SpinboxFrame(size_frame2, initial=6, min_val=2, max_val=20, width=60, command=self.update_preview)
        self.num_size.pack(side="left")
        size_frame2.pack(pady=2)

        self.num_align = ctk.StringVar(value="center")
        align_frame2 = ctk.CTkFrame(control_frame, fg_color="transparent")
        ctk.CTkRadioButton(align_frame2, text="Слева", variable=self.num_align, value="left", command=self.update_preview).pack(side="left")
        ctk.CTkRadioButton(align_frame2, text="По центру", variable=self.num_align, value="center", command=self.update_preview).pack(side="left")
        align_frame2.pack(pady=2)

        self.num_color = "#000000"
        color_frame2 = ctk.CTkFrame(control_frame, fg_color="transparent")
        ctk.CTkButton(color_frame2, text="🎨", command=self.pick_num_color, width=36, height=24).pack(side="left", padx=(0, 5))
        self.num_color_indicator = ctk.CTkLabel(color_frame2, text="", width=20, height=20, corner_radius=4, fg_color=self.num_color)
        self.num_color_indicator.pack(side="left", padx=(0, 5))
        self.num_hex_entry = ctk.CTkEntry(color_frame2, placeholder_text="#000000", width=80)
        self.num_hex_entry.insert(0, self.num_color)
        self.num_hex_entry.pack(side="left")
        self.num_hex_entry.bind("<FocusOut>", self.on_num_hex_focus_out)
        self.num_hex_entry.bind("<Return>", self.on_num_hex_focus_out)
        color_frame2.pack(pady=5, padx=10, fill="x")

        # === Поле 3: Фикс. текст ===
        ctk.CTkLabel(control_frame, text="📝 Фикс. текст (в мм)").pack(pady=(15, 5))
        self.fix_text = ctk.CTkEntry(control_frame, placeholder_text="Текст")
        self.fix_text.insert(0, "28 декабря 2025")
        self.fix_text.pack(padx=10, fill="x")
        self.fix_text.bind("<KeyRelease>", lambda e: self.update_preview())
        self.fix_text.bind("<FocusOut>", lambda e: self.update_preview())

        coord_frame3 = ctk.CTkFrame(control_frame, fg_color="transparent")
        ctk.CTkLabel(coord_frame3, text="X:", width=20).pack(side="left", padx=(0, 5))
        self.fix_x = SpinboxFrame(coord_frame3, initial=105, min_val=0, max_val=A4_WIDTH_MM, command=self.update_preview)
        self.fix_x.pack(side="left", padx=(0, 10))
        ctk.CTkLabel(coord_frame3, text="Y:", width=20).pack(side="left", padx=(0, 5))
        self.fix_y = SpinboxFrame(coord_frame3, initial=140, min_val=0, max_val=A4_HEIGHT_MM, command=self.update_preview)
        self.fix_y.pack(side="left")
        coord_frame3.pack(pady=2)

        size_frame3 = ctk.CTkFrame(control_frame, fg_color="transparent")
        ctk.CTkLabel(size_frame3, text="Размер текста:", width=100).pack(side="left", padx=(0, 5))
        self.fix_size = SpinboxFrame(size_frame3, initial=5, min_val=2, max_val=15, width=60, command=self.update_preview)
        self.fix_size.pack(side="left")
        size_frame3.pack(pady=2)

        self.fix_align = ctk.StringVar(value="center")
        align_frame3 = ctk.CTkFrame(control_frame, fg_color="transparent")
        ctk.CTkRadioButton(align_frame3, text="Слева", variable=self.fix_align, value="left", command=self.update_preview).pack(side="left")
        ctk.CTkRadioButton(align_frame3, text="По центру", variable=self.fix_align, value="center", command=self.update_preview).pack(side="left")
        align_frame3.pack(pady=2)

        self.fix_color = "#000000"
        color_frame3 = ctk.CTkFrame(control_frame, fg_color="transparent")
        ctk.CTkButton(color_frame3, text="🎨", command=self.pick_fix_color, width=36, height=24).pack(side="left", padx=(0, 5))
        self.fix_color_indicator = ctk.CTkLabel(color_frame3, text="", width=20, height=20, corner_radius=4, fg_color=self.fix_color)
        self.fix_color_indicator.pack(side="left", padx=(0, 5))
        self.fix_hex_entry = ctk.CTkEntry(color_frame3, placeholder_text="#000000", width=80)
        self.fix_hex_entry.insert(0, self.fix_color)
        self.fix_hex_entry.pack(side="left")
        self.fix_hex_entry.bind("<FocusOut>", self.on_fix_hex_focus_out)
        self.fix_hex_entry.bind("<Return>", self.on_fix_hex_focus_out)
        color_frame3.pack(pady=5, padx=10, fill="x")

        ctk.CTkButton(control_frame, text="📄 Загрузить шаблон (PDF/PNG)", command=self.load_template).pack(pady=10, padx=10, fill="x")
        ctk.CTkButton(control_frame, text="📂 Загрузить список имён (.txt)", command=self.load_names_file).pack(pady=5, padx=10, fill="x")
        ctk.CTkButton(control_frame, text="💾 Сохранить настройки", command=self.save_settings).pack(pady=10, padx=10, fill="x")
        ctk.CTkButton(control_frame, text="🖼️ Один PNG", command=self.generate_single_png).pack(pady=5, padx=10, fill="x")
        ctk.CTkButton(control_frame, text="🖼️ Пачка PNG (ZIP)", command=self.generate_batch_png).pack(pady=5, padx=10, fill="x")

        self.status_label = ctk.CTkLabel(control_frame, text="", height=20)
        self.status_label.pack(pady=10)

        copyright_label = ctk.CTkLabel(
            control_frame,
            text="© 2025 Владимир Сергеевич Снигур \n Это свободное ПО \n Лицензия GPL 3.0 \n \n v. 1.3 \n snigurmd@yandex.ru",
            text_color="gray70",
            font=ctk.CTkFont(size=10),
            justify="left"
        )
        copyright_label.pack(side="bottom", pady=(20, 10))

        preview_frame = ctk.CTkFrame(self)
        preview_frame.pack(side="right", fill="both", expand=True, padx=10, pady=10)
        self.preview_label = ctk.CTkLabel(preview_frame, text="👀 Предпросмотр\n(загрузите шаблон)")
        self.preview_label.pack(expand=True)

    def pick_name_color(self):
        color = colorchooser.askcolor(initialcolor=self.name_color, title="Цвет имени")[1]
        if color:
            self.name_color = color.upper()
            self.name_color_indicator.configure(fg_color=self.name_color)
            self.name_hex_entry.delete(0, tk.END)
            self.name_hex_entry.insert(0, self.name_color)
            self.update_preview()

    def pick_num_color(self):
        color = colorchooser.askcolor(initialcolor=self.num_color, title="Цвет номера")[1]
        if color:
            self.num_color = color.upper()
            self.num_color_indicator.configure(fg_color=self.num_color)
            self.num_hex_entry.delete(0, tk.END)
            self.num_hex_entry.insert(0, self.num_color)
            self.update_preview()

    def pick_fix_color(self):
        color = colorchooser.askcolor(initialcolor=self.fix_color, title="Цвет фикс. текста")[1]
        if color:
            self.fix_color = color.upper()
            self.fix_color_indicator.configure(fg_color=self.fix_color)
            self.fix_hex_entry.delete(0, tk.END)
            self.fix_hex_entry.insert(0, self.fix_color)
            self.update_preview()

    def _validate_and_apply_hex(self, entry, indicator, color_attr):
        raw = entry.get().strip()
        if not raw:
            return
        raw = raw.upper()
        if not raw.startswith("#"):
            raw = "#" + raw
        if len(raw) > 7:
            raw = raw[:7]
        if is_valid_hex(raw):
            setattr(self, color_attr, raw)
            indicator.configure(fg_color=raw)
            entry.delete(0, tk.END)
            entry.insert(0, raw)
            self.update_preview()
        else:
            current = getattr(self, color_attr)
            entry.delete(0, tk.END)
            entry.insert(0, current)

    def on_name_hex_focus_out(self, event=None): self._validate_and_apply_hex(self.name_hex_entry, self.name_color_indicator, 'name_color')
    def on_num_hex_focus_out(self, event=None): self._validate_and_apply_hex(self.num_hex_entry, self.num_color_indicator, 'num_color')
    def on_fix_hex_focus_out(self, event=None): self._validate_and_apply_hex(self.fix_hex_entry, self.fix_color_indicator, 'fix_color')

    def load_template(self):
        path = filedialog.askopenfilename(filetypes=[
            ("Шаблоны", "*.pdf *.png *.jpg *.jpeg"),
            ("PDF", "*.pdf"),
            ("PNG", "*.png"),
            ("JPEG", "*.jpg *.jpeg")
        ])
        if not path:
            return
        self.template_path = path
        try:
            img, dpi, orientation = load_template_image(path, base_dpi=BASE_DPI)
            self.cached_template_img = img
            self.cached_dpi = dpi
            self.cached_orientation = orientation
            self.update_spinbox_ranges_for_orientation(orientation)
            self.status_label.configure(text=f"({img.width}×{img.height}, DPI≈{dpi:.0f}, {'альбомная' if orientation == 'landscape' else 'книжная'})")
            self.update_preview()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить шаблон: {e}")

    def load_names_file(self):
        path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        if path:
            self.names_file_path = path

    def get_fields_mm(self):
        return {
            "name": {
                "text": self.name_text.get(),
                "x": safe_int(self.name_x.get(), 105),
                "y": safe_int(self.name_y.get(), 100),
                "size": safe_int(self.name_size.get(), 8),
                "align": self.name_align.get(),
                "color": self.name_color
            },
            "num": {
                "start": safe_int(self.num_start.get(), 1),
                "x": safe_int(self.num_x.get(), 105),
                "y": safe_int(self.num_y.get(), 120),
                "size": safe_int(self.num_size.get(), 6),
                "align": self.num_align.get(),
                "color": self.num_color
            },
            "fix": {
                "text": self.fix_text.get(),
                "x": safe_int(self.fix_x.get(), 105),
                "y": safe_int(self.fix_y.get(), 140),
                "size": safe_int(self.fix_size.get(), 5),
                "align": self.fix_align.get(),
                "color": self.fix_color
            }
        }

    def update_preview(self, *args):
        if self.cached_template_img is None:
            self.preview_label.configure(image=None, text="👀 Предпросмотр\n(загрузите шаблон)")
            return
        try:
            f = self.get_fields_mm()
            fields = [
                {"text": f["name"]["text"], "x": mm_to_px(f["name"]["x"], self.cached_dpi), "y": mm_to_px(f["name"]["y"], self.cached_dpi),
                 "size": mm_to_px(f["name"]["size"], self.cached_dpi), "align": f["name"]["align"], "color": f["name"]["color"], "font": self.font_var.get()},
                {"text": f"{f['num']['start']:04d}", "x": mm_to_px(f["num"]["x"], self.cached_dpi), "y": mm_to_px(f["num"]["y"], self.cached_dpi),
                 "size": mm_to_px(f["num"]["size"], self.cached_dpi), "align": f["num"]["align"], "color": f["num"]["color"], "font": self.font_var.get()},
                {"text": f["fix"]["text"], "x": mm_to_px(f["fix"]["x"], self.cached_dpi), "y": mm_to_px(f["fix"]["y"], self.cached_dpi),
                 "size": mm_to_px(f["fix"]["size"], self.cached_dpi), "align": f["fix"]["align"], "color": f["fix"]["color"], "font": self.font_var.get()},
            ]
            preview_img = draw_preview(self.cached_template_img, fields, preview_scale=self.preview_scale.get())
            self.preview_img_tk = ctk.CTkImage(dark_image=preview_img, size=preview_img.size)
            self.preview_label.configure(image=self.preview_img_tk, text="")
        except Exception as e:
            self.status_label.configure(text=f"❌ Ошибка предпросмотра: {e}")

    def save_settings(self):
        data = {
            "font_name": self.font_var.get(),
            "preview_scale": self.preview_scale.get(),
            "field_1": {
                "x": self.name_x.get(), "y": self.name_y.get(),
                "size": int(self.name_size.get()), "align": self.name_align.get(),
                "preview_text": self.name_text.get(),
                "color": self.name_color
            },
            "field_2": {
                "start_num": self.num_start.get(),
                "x": self.num_x.get(), "y": self.num_y.get(),
                "size": int(self.num_size.get()), "align": self.num_align.get(),
                "color": self.num_color
            },
            "field_3": {
                "text": self.fix_text.get(),
                "x": self.fix_x.get(), "y": self.fix_y.get(),
                "size": int(self.fix_size.get()), "align": self.fix_align.get(),
                "color": self.fix_color
            }
        }
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        self.status_label.configure(text=f"✅ Сохранено в {SETTINGS_FILE}")

    def load_settings_on_start(self):
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, encoding="utf-8") as f:
                    d = json.load(f)
                self.font_var.set(d.get("font_name", "Helvetica"))
                self.preview_scale.set(d.get("preview_scale", 1.0))
                f1 = d.get("field_1", {})
                self.name_text.delete(0, tk.END); self.name_text.insert(0, f1.get("preview_text", "Иванов Иван"))
                self.name_x.set(f1.get("x", "105")); self.name_y.set(f1.get("y", "100"))
                self.name_size.set(f1.get("size", "8")); self.name_align.set(f1.get("align", "center"))
                self.name_color = f1.get("color", "#000000")
                self.name_color_indicator.configure(fg_color=self.name_color)
                self.name_hex_entry.delete(0, tk.END); self.name_hex_entry.insert(0, self.name_color)
                f2 = d.get("field_2", {})
                self.num_start.delete(0, tk.END); self.num_start.insert(0, f2.get("start_num", "1"))
                self.num_x.set(f2.get("x", "105")); self.num_y.set(f2.get("y", "120"))
                self.num_size.set(f2.get("size", "6")); self.num_align.set(f2.get("align", "center"))
                self.num_color = f2.get("color", "#000000")
                self.num_color_indicator.configure(fg_color=self.num_color)
                self.num_hex_entry.delete(0, tk.END); self.num_hex_entry.insert(0, self.num_color)
                f3 = d.get("field_3", {})
                self.fix_text.delete(0, tk.END); self.fix_text.insert(0, f3.get("text", "28 декабря 2025"))
                self.fix_x.set(f3.get("x", "105")); self.fix_y.set(f3.get("y", "140"))
                self.fix_size.set(f3.get("size", "5")); self.fix_align.set(f3.get("align", "center"))
                self.fix_color = f3.get("color", "#000000")
                self.fix_color_indicator.configure(fg_color=self.fix_color)
                self.fix_hex_entry.delete(0, tk.END); self.fix_hex_entry.insert(0, self.fix_color)
            except Exception as e:
                self.status_label.configure(text=f"⚠️ Настройки не загружены: {e}")

    def generate_single_png(self):
        if not self.template_path:
            messagebox.showerror("Ошибка", "Сначала загрузите шаблон")
            return
        try:
            f = self.get_fields_mm()
            result = generate_single_png(
                self.template_path, self.font_var.get(),
                f["name"]["text"], f["name"]["x"], f["name"]["y"], f["name"]["size"], f["name"]["align"], f["name"]["color"],
                f["num"]["start"],
                f["num"]["x"], f["num"]["y"], f["num"]["size"], f["num"]["align"], f["num"]["color"],
                f["fix"]["text"], f["fix"]["x"], f["fix"]["y"], f["fix"]["size"], f["fix"]["align"], f["fix"]["color"]
            )
            if result[0]:
                open_file_or_dir(result[0])
                self.status_label.configure(text="✅ Один PNG готов")
            else:
                messagebox.showerror("Ошибка", result[1])
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def generate_batch_png(self):
        if not self.template_path or not self.names_file_path:
            messagebox.showerror("Ошибка", "Загрузите шаблон и список имён")
            return
        try:
            f = self.get_fields_mm()
            result = generate_batch_png(
                self.template_path, self.names_file_path, self.font_var.get(),
                f["num"]["start"],
                f["name"]["x"], f["name"]["y"], f["name"]["size"], f["name"]["align"], f["name"]["color"],
                f["num"]["x"], f["num"]["y"], f["num"]["size"], f["num"]["align"], f["num"]["color"],
                f["fix"]["text"], f["fix"]["x"], f["fix"]["y"], f["fix"]["size"], f["fix"]["align"], f["fix"]["color"]
            )
            if result[0]:
                open_file_or_dir(os.path.dirname(result[0]))
                self.status_label.configure(text=result[1])
            else:
                messagebox.showerror("Ошибка", result[1])
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

if __name__ == "__main__":
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")
    app = CertificateApp()
    app.mainloop()