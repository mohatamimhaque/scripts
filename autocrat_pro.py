#!/usr/bin/env python3
"""
AutoCrat Pro -- Enterprise Document Automation Suite
Unified standalone CLI & GUI runner for high-volume document generation with optional QR code and image support.
"""

# ==============================================================================
# Cell 2: AutoCrat Pro Production Core Engine (with QR & Image Support)
# ==============================================================================
import os
import re
import sys
import copy
import time
import urllib.request
import concurrent.futures
from typing import List, Dict, Optional, Tuple, Any
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pypdf import PdfWriter
from PIL import Image, ImageDraw, ImageColor, ImageOps

try:
    from tqdm import tqdm
except ImportError:
    def tqdm(iterable, **kwargs):
        return iterable


# ------------------------------------------------------------------------------
# 1. PowerPoint to PDF COM Converter
# ------------------------------------------------------------------------------
class PPTXToPDFConverter:
    """
    Manages a persistent Microsoft PowerPoint COM application instance for 
    ultra-fast batch PDF rendering on Windows.
    """
    def __init__(self):
        self.powerpoint = None
        self.com_initialized = False
        self.backend = None

    def start(self) -> bool:
        if self.com_initialized and self.powerpoint is not None:
            return True

        # Attempt 1: win32com.client
        try:
            import win32com.client
            self.powerpoint = win32com.client.Dispatch("Powerpoint.Application")
            self.com_initialized = True
            self.backend = "win32com"
            print("  [OK] PowerPoint COM engine initialized successfully (win32com).")
            return True
        except Exception:
            pass

        # Attempt 2: comtypes.client
        try:
            import comtypes.client
            self.powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
            self.com_initialized = True
            self.backend = "comtypes"
            print("  [OK] PowerPoint COM engine initialized successfully (comtypes).")
            return True
        except Exception:
            pass

        print("  [WARN] Could not initialize PowerPoint COM. Native PDF conversion disabled.")
        print("         Ensure Microsoft Office PowerPoint is installed on this Windows machine.")
        return False

    def convert(self, pptx_path: str, pdf_path: str) -> bool:
        if not self.com_initialized or self.powerpoint is None:
            return False
        
        abs_pptx = os.path.abspath(pptx_path)
        abs_pdf = os.path.abspath(pdf_path)

        deck = None
        try:
            deck = self.powerpoint.Presentations.Open(abs_pptx, WithWindow=False)
            deck.SaveAs(abs_pdf, 32)
            deck.Close()
            deck = None
            return True
        except Exception as err:
            print(f"    [ERR] COM conversion error for {os.path.basename(pptx_path)}: {err}")
            if deck is not None:
                try: deck.Close()
                except Exception: pass
            return False

    def stop(self):
        if self.powerpoint is not None:
            try: self.powerpoint.Quit()
            except Exception: pass
            self.powerpoint = None
            self.com_initialized = False
            self.backend = None
            print("  [OK] PowerPoint COM engine released.")


# ------------------------------------------------------------------------------
# 2. Intelligent Placeholder Replacement Engine
# ------------------------------------------------------------------------------
class PlaceholderEngine:
    """
    Advanced placeholder replacer supporting:
    - Fuzzy space and case normalization
    - Regular text frames
    - Table cells (rows & columns)
    - Grouped shapes (recursive traversal)
    - Multi-run formatting preservation
    """
    @staticmethod
    def normalize_key(key: Any) -> str:
        return re.sub(r"[\s_]+", "", str(key)).lower()

    @classmethod
    def build_mapping(cls, row: pd.Series) -> Dict[str, str]:
        mapping = {}
        for col_name, val in row.items():
            str_val = "" if pd.isna(val) else str(val).strip()
            if re.match(r"^\d+\.0$", str_val):
                str_val = str_val[:-2]
            raw_key = str(col_name).strip()
            norm_key = cls.normalize_key(raw_key)
            mapping[raw_key] = str_val
            mapping[norm_key] = str_val
        return mapping

    @classmethod
    def replace_text(cls, text: str, mapping: Dict[str, str]) -> str:
        if not text or "{{" not in text:
            return text

        def replacer(match):
            raw_token = match.group(1).strip()
            norm_token = cls.normalize_key(raw_token)
            if raw_token in mapping:
                return mapping[raw_token]
            if norm_token in mapping:
                return mapping[norm_token]
            return match.group(0)

        return re.sub(r"\{\{\s*([^}]+)\s*\}\}", replacer, text)

    @classmethod
    def process_text_frame(cls, text_frame, mapping: Dict[str, str]):
        for para in text_frame.paragraphs:
            full_text = "".join(run.text for run in para.runs)
            # Skip dedicated QR and Image placeholder shapes; handled by their respective engines
            if re.search(r"^\{\{\s*(qr|image|photo|picture)\s*\}\}$", full_text.strip(), re.IGNORECASE):
                continue
            if "{{" in full_text:
                new_text = cls.replace_text(full_text, mapping)
                for i in range(len(para.runs) - 1, 0, -1):
                    try: para.runs[i].text = ""
                    except Exception: pass
                if len(para.runs) > 0:
                    para.runs[0].text = new_text
                else:
                    para.add_run().text = new_text

    @classmethod
    def process_shape(cls, shape, mapping: Dict[str, str]):
        if shape.has_text_frame:
            cls.process_text_frame(shape.text_frame, mapping)
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    cls.process_text_frame(cell.text_frame, mapping)
        if hasattr(shape, "shapes"):
            for child_shape in shape.shapes:
                cls.process_shape(child_shape, mapping)

    @classmethod
    def process_slide(cls, slide, mapping: Dict[str, str]):
        for shape in slide.shapes:
            cls.process_shape(shape, mapping)

    @classmethod
    def extract_placeholders(cls, prs: Presentation) -> List[str]:
        found = set()

        def scan_text(txt):
            if txt and "{{" in txt:
                for m in re.finditer(r"\{\{\s*([^}]+)\s*\}\}", txt):
                    found.add(m.group(1).strip())

        def scan_shape(sh):
            if sh.has_text_frame:
                for p in sh.text_frame.paragraphs:
                    scan_text(p.text)
            if sh.has_table:
                for r in sh.table.rows:
                    for c in r.cells:
                        for p in c.text_frame.paragraphs:
                            scan_text(p.text)
            if hasattr(sh, "shapes"):
                for child in sh.shapes:
                    scan_shape(child)

        for s in prs.slides:
            for sh in s.shapes:
                scan_shape(sh)

        return sorted(list(found))


# ------------------------------------------------------------------------------
# 3. Dynamic QR Code Engine (Config & Handler)
# ------------------------------------------------------------------------------
class QRCodeConfig:
    def __init__(
        self,
        enabled: bool = False,
        data_column: Optional[str] = None,
        data_template: Optional[str] = None,
        position_mode: str = "placeholder",
        size_mode: str = "auto",
        custom_size_inches: float = 1.5,
        custom_left_inches: float = 0.0,
        custom_top_inches: float = 0.0,
        margin_inches: float = 0.5,
        box_size: int = 10,
        border: int = 1
    ):
        self.enabled = enabled
        self.data_column = data_column
        self.data_template = data_template
        self.position_mode = str(position_mode).lower()
        self.size_mode = str(size_mode).lower()
        self.custom_size_inches = float(custom_size_inches)
        self.custom_left_inches = float(custom_left_inches)
        self.custom_top_inches = float(custom_top_inches)
        self.margin_inches = float(margin_inches)
        self.box_size = box_size
        self.border = border

    def resolve_payload(self, row_map: Dict[str, str], raw_row: pd.Series) -> str:
        if not self.enabled:
            return ""
        if self.data_template:
            return PlaceholderEngine.replace_text(self.data_template, row_map).strip()
        if self.data_column and self.data_column in raw_row:
            val = raw_row[self.data_column]
            return "" if pd.isna(val) else str(val).strip()
        
        for key in ["serial", "id", "roll", "registration_id", "cert_id", "url", "qr"]:
            for col, val in row_map.items():
                if PlaceholderEngine.normalize_key(col) == key and val:
                    return str(val).strip()
        return ""


class QRCodeHandler:
    @staticmethod
    def generate_image(payload: str, out_path: str, box_size: int = 10, border: int = 1) -> bool:
        if not payload:
            return False
        try:
            import qrcode
            qr = qrcode.QRCode(version=1, box_size=box_size, border=border)
            qr.add_data(payload)
            qr.make(fit=True)
            img = qr.make_image(fill_color="black", back_color="white")
            img.save(out_path)
            return True
        except Exception as e:
            print(f"    [WARN] QR generation error: {e}")
            return False

    @classmethod
    def embed_in_slide(cls, slide, prs: Presentation, qr_path: str, config: QRCodeConfig):
        if not config.enabled or not qr_path or not os.path.exists(qr_path):
            return

        placeholder_shape = None
        for shape in slide.shapes:
            if shape.has_text_frame and re.search(r"\{\{\s*qr\s*\}\}", shape.text_frame.text, re.IGNORECASE):
                placeholder_shape = shape
                break

        slide_w = prs.slide_width
        slide_h = prs.slide_height

        # Determine target QR size:
        # If user specified a custom size (or default), ALWAYS strictly follow user input!
        if config.size_mode == "custom" and config.custom_size_inches > 0:
            sq_size = int(config.custom_size_inches * 914400)
        elif placeholder_shape is not None and config.size_mode == "auto":
            # Auto-fit to placeholder shape, but guard against giant / full-slide container boxes!
            raw_size = max(placeholder_shape.width, placeholder_shape.height) if min(placeholder_shape.width, placeholder_shape.height) < 914400 * 0.8 else min(placeholder_shape.width, placeholder_shape.height)
            # Never let auto-fit exceed 3.5 inches or 35% of slide dimensions
            max_limit = min(int(3.5 * 914400), int(min(slide_w, slide_h) * 0.35))
            sq_size = min(raw_size, max_limit)
            if sq_size < int(0.8 * 914400):  # Never make QR smaller than 0.8 inches
                sq_size = int(config.custom_size_inches * 914400 if config.custom_size_inches > 0 else 1.5 * 914400)
        else:
            sq_size = int((config.custom_size_inches if config.custom_size_inches > 0 else 1.5) * 914400)

        # Cap QR code so it never exceeds 50% of the slide
        sq_size = min(sq_size, int(min(slide_w, slide_h) * 0.5))

        if placeholder_shape is not None and config.position_mode in ["placeholder", "auto"]:
            orig_left = placeholder_shape.left
            orig_top = placeholder_shape.top
            orig_width = placeholder_shape.width
            orig_height = placeholder_shape.height

            left = orig_left + (orig_width - sq_size) // 2
            top = orig_top + (orig_height - sq_size) // 2

            # Clamp within slide boundaries
            left = max(0, min(left, slide_w - sq_size))
            top = max(0, min(top, slide_h - sq_size))

            slide.shapes.add_picture(qr_path, left, top, sq_size, sq_size)
            try:
                elem = placeholder_shape._element
                elem.getparent().remove(elem)
            except Exception:
                try: placeholder_shape.text_frame.text = ""
                except Exception: pass
            return

        margin = int(config.margin_inches * 914400)
        pos = config.position_mode
        if pos == "bottom-right":
            left = slide_w - sq_size - margin
            top = slide_h - sq_size - margin
        elif pos == "bottom-left":
            left = margin
            top = slide_h - sq_size - margin
        elif pos == "top-right":
            left = slide_w - sq_size - margin
            top = margin
        elif pos == "top-left":
            left = margin
            top = margin
        elif pos == "center":
            left = (slide_w - sq_size) // 2
            top = (slide_h - sq_size) // 2
        elif pos == "custom":
            left = int(config.custom_left_inches * 914400)
            top = int(config.custom_top_inches * 914400)
        else:
            left = slide_w - sq_size - margin
            top = slide_h - sq_size - margin

        left = max(0, min(left, slide_w - sq_size))
        top = max(0, min(top, slide_h - sq_size))

        slide.shapes.add_picture(qr_path, left, top, sq_size, sq_size)


# ------------------------------------------------------------------------------
# 4. Dynamic Image & Photo Engine (Config & Handler)
# ------------------------------------------------------------------------------
class ImageConfig:
    """
    Configuration for optional dynamic image / photo insertion:
    - enabled: True / False (NOT mandatory)
    - data_column: Column name containing file path or web URL (e.g. 'Photo', 'Image_URL')
    - image_folder: Optional local directory where images live (e.g. images/{Serial}.jpg)
    - id_column: Column name to match against image files in image_folder
    - position_mode: 'placeholder' (replaces {{image}} or {{photo}} box), 'top-right', 'top-left', 'bottom-right', 'bottom-left', 'custom'
    - size_mode: 'auto' (fits placeholder box preserving aspect ratio), 'custom'
    - custom_width_inches / custom_height_inches: Dimensions if size_mode == 'custom'
    - keep_aspect_ratio: True (prevents photo distortion)
    - shape_mode: 'square' (sharp corners), 'rounded' (smooth curves), 'circle' (circular avatar/oval)
    - corner_radius: Corner radius in pixels if shape_mode == 'rounded'
    - border_width: Border stroke width in pixels (0 for none)
    - border_color: Hex color string (e.g. '#FFFFFF', '#000000', '#0066CC')
    """
    def __init__(
        self,
        enabled: bool = False,
        data_column: Optional[str] = None,
        image_folder: Optional[str] = None,
        id_column: Optional[str] = None,
        position_mode: str = "placeholder",
        size_mode: str = "auto",
        custom_width_inches: float = 1.5,
        custom_height_inches: float = 2.0,
        custom_left_inches: float = 7.5,
        custom_top_inches: float = 1.5,
        margin_inches: float = 0.5,
        keep_aspect_ratio: bool = True,
        shape_mode: str = "square",
        corner_radius: int = 25,
        border_width: int = 0,
        border_color: str = "#FFFFFF"
    ):
        self.enabled = enabled
        self.data_column = data_column
        self.image_folder = image_folder
        self.id_column = id_column
        self.position_mode = str(position_mode).lower()
        self.size_mode = str(size_mode).lower()
        self.custom_width_inches = float(custom_width_inches)
        self.custom_height_inches = float(custom_height_inches)
        self.custom_left_inches = float(custom_left_inches)
        self.custom_top_inches = float(custom_top_inches)
        self.margin_inches = float(margin_inches)
        self.keep_aspect_ratio = keep_aspect_ratio
        self.shape_mode = str(shape_mode).lower()
        self.corner_radius = int(corner_radius)
        self.border_width = int(border_width)
        self.border_color = str(border_color).strip()


class ImageHandler:
    """Resolves image paths/URLs, applies optional corner radius & border styling, and embeds them accurately onto slides."""
    @staticmethod
    def resolve_image_path(config: ImageConfig, row: pd.Series, row_map: Dict[str, str], temp_dir: str) -> Tuple[Optional[str], bool]:
        """
        Returns (image_file_path, is_temporary)
        """
        if not config.enabled:
            return None, False

        target_val = None
        if config.data_column and config.data_column in row:
            v = row[config.data_column]
            if pd.notna(v) and str(v).strip():
                target_val = str(v).strip()

        # Heuristic detection for columns like photo, image, picture, avatar
        if not target_val:
            for c, v in row_map.items():
                if PlaceholderEngine.normalize_key(c) in ["photo", "image", "picture", "avatar", "photourl", "imageurl"] and v:
                    target_val = str(v).strip()
                    break

        # Check if URL
        if target_val and target_val.startswith(("http://", "https://")):
            try:
                ext = os.path.splitext(target_val.split("?")[0])[1] or ".jpg"
                dl_path = os.path.join(temp_dir, f"_temp_img_{int(time.time() * 1000)}{ext}")
                urllib.request.urlretrieve(target_val, dl_path)
                return dl_path, True
            except Exception as e:
                print(f"    [WARN] Failed to download image from {target_val}: {e}")
                return None, False

        # Check local file
        if target_val and os.path.isfile(target_val):
            return os.path.abspath(target_val), False

        # Check image_folder matching
        if config.image_folder and os.path.isdir(config.image_folder):
            id_val = ""
            if config.id_column and config.id_column in row:
                id_val = str(row[config.id_column]).strip()
            elif target_val:
                id_val = target_val
            else:
                for c in ["serial", "id", "roll"]:
                    if c in row_map and row_map[c]:
                        id_val = row_map[c]
                        break

            if id_val:
                for ext in [".jpg", ".jpeg", ".png", ".webp", ".bmp"]:
                    cand = os.path.join(config.image_folder, f"{id_val}{ext}")
                    if os.path.isfile(cand):
                        return cand, False

        return None, False

    @classmethod
    def apply_styling(cls, src_path: str, config: ImageConfig, dest_path: str) -> str:
        """
        Applies corner radius (square, rounded, or circle) and outside border stroke (width & color)
        using anti-aliased alpha supersampling in Pillow without clipping into the image.
        """
        if config.shape_mode == "square" and config.border_width <= 0:
            return src_path

        try:
            from PIL import Image, ImageDraw, ImageColor
            with Image.open(src_path) as orig:
                im = orig.convert("RGBA")
            
            w, h = im.size
            scale = 2
            w_s, h_s = w * scale, h * scale
            bw = max(0, int(config.border_width))
            bw_s = int(bw * scale)
            im_s = im.resize((w_s, h_s), Image.Resampling.LANCZOS)
            
            # Mask for the inner image content
            mask = Image.new("L", (w_s, h_s), 0)
            draw_mask = ImageDraw.Draw(mask)
            
            # Canvas with space for outside border
            total_w_s = w_s + 2 * bw_s
            total_h_s = h_s + 2 * bw_s
            out = Image.new("RGBA", (total_w_s, total_h_s), (0, 0, 0, 0))
            
            if bw_s > 0 and config.border_color:
                b_draw = ImageDraw.Draw(out)
                try:
                    c = ImageColor.getrgb(config.border_color)
                except Exception:
                    c = (255, 255, 255)
                c_rgba = (c[0], c[1], c[2], 255)

                if config.shape_mode == "circle":
                    b_draw.ellipse((0, 0, total_w_s - 1, total_h_s - 1), fill=c_rgba)
                    draw_mask.ellipse((0, 0, w_s - 1, h_s - 1), fill=255)
                elif config.shape_mode == "rounded":
                    r = min(int(config.corner_radius * scale), min(w_s, h_s) // 2)
                    r_outer = r + bw_s
                    b_draw.rounded_rectangle((0, 0, total_w_s - 1, total_h_s - 1), radius=r_outer, fill=c_rgba)
                    draw_mask.rounded_rectangle((0, 0, w_s - 1, h_s - 1), radius=r, fill=255)
                else:
                    b_draw.rectangle((0, 0, total_w_s - 1, total_h_s - 1), fill=c_rgba)
                    draw_mask.rectangle((0, 0, w_s - 1, h_s - 1), fill=255)
            else:
                if config.shape_mode == "circle":
                    draw_mask.ellipse((0, 0, w_s - 1, h_s - 1), fill=255)
                elif config.shape_mode == "rounded":
                    r = min(int(config.corner_radius * scale), min(w_s, h_s) // 2)
                    draw_mask.rounded_rectangle((0, 0, w_s - 1, h_s - 1), radius=r, fill=255)
                else:
                    draw_mask.rectangle((0, 0, w_s - 1, h_s - 1), fill=255)
                
            # Paste the complete inner image without any border covering it
            out.paste(im_s, (bw_s, bw_s), mask=mask)
                    
            final_img = out.resize((w + 2 * bw, h + 2 * bw), Image.Resampling.LANCZOS)
            final_img.save(dest_path, "PNG")
            return dest_path
        except Exception as e:
            print(f"    [WARN] Failed to apply image styling: {e}")
            return src_path

    @classmethod
    def embed_in_slide(cls, slide, prs: Presentation, image_path: str, config: ImageConfig):
        if not config.enabled or not image_path or not os.path.isfile(image_path):
            return

        temp_styled = None
        actual_img_path = image_path
        if config.shape_mode in ["rounded", "circle"] or config.border_width > 0:
            dir_name = os.path.dirname(image_path) or "."
            base_name = os.path.splitext(os.path.basename(image_path))[0]
            temp_styled = os.path.join(dir_name, f"_styled_{int(time.time() * 1000)}_{base_name}.png")
            actual_img_path = cls.apply_styling(image_path, config, temp_styled)

        try:
            with Image.open(actual_img_path) as im:
                img_w, img_h = im.size
            img_aspect = img_w / img_h
        except Exception:
            img_aspect = 1.0

        try:
            # Search for {{image}} or {{photo}} or {{picture}}
            placeholder_shape = None
            for shape in slide.shapes:
                if shape.has_text_frame and re.search(r"\{\{\s*(image|photo|picture)\s*\}\}", shape.text_frame.text, re.IGNORECASE):
                    placeholder_shape = shape
                    break

            if placeholder_shape is not None and config.position_mode in ["placeholder", "auto"]:
                orig_left = placeholder_shape.left
                orig_top = placeholder_shape.top
                orig_width = placeholder_shape.width
                orig_height = placeholder_shape.height

                if config.size_mode == "custom":
                    box_w = int(config.custom_width_inches * 914400)
                    box_h = int(config.custom_height_inches * 914400)
                else:
                    box_w = orig_width
                    box_h = orig_height

                if config.keep_aspect_ratio:
                    box_aspect = box_w / box_h
                    if img_aspect > box_aspect:
                        new_w = box_w
                        new_h = int(box_w / img_aspect)
                    else:
                        new_h = box_h
                        new_w = int(box_h * img_aspect)
                else:
                    new_w, new_h = box_w, box_h

                left = orig_left + (box_w - new_w) // 2
                top = orig_top + (box_h - new_h) // 2

                slide.shapes.add_picture(actual_img_path, left, top, new_w, new_h)
                try:
                    elem = placeholder_shape._element
                    elem.getparent().remove(elem)
                except Exception:
                    try: placeholder_shape.text_frame.text = ""
                    except Exception: pass
                return

            # Preset or custom positioning
            box_w = int(config.custom_width_inches * 914400)
            box_h = int(config.custom_height_inches * 914400)
            margin = int(config.margin_inches * 914400)
            slide_w = prs.slide_width
            slide_h = prs.slide_height

            if config.keep_aspect_ratio:
                box_aspect = box_w / box_h
                if img_aspect > box_aspect:
                    new_w = box_w
                    new_h = int(box_w / img_aspect)
                else:
                    new_h = box_h
                    new_w = int(box_h * img_aspect)
            else:
                new_w, new_h = box_w, box_h

            pos = config.position_mode
            if pos == "top-right":
                left = slide_w - new_w - margin
                top = margin
            elif pos == "top-left":
                left = margin
                top = margin
            elif pos == "bottom-right":
                left = slide_w - new_w - margin
                top = slide_h - new_h - margin
            elif pos == "bottom-left":
                left = margin
                top = slide_h - new_h - margin
            elif pos == "center":
                left = (slide_w - new_w) // 2
                top = (slide_h - new_h) // 2
            elif pos == "custom":
                left = int(config.custom_left_inches * 914400)
                top = int(config.custom_top_inches * 914400)
            else:
                left = slide_w - new_w - margin
                top = margin

            slide.shapes.add_picture(actual_img_path, left, top, new_w, new_h)
        finally:
            if temp_styled and os.path.exists(temp_styled):
                try: os.remove(temp_styled)
                except Exception: pass


class SlideDuplicator:
    @staticmethod
    def duplicate_slide(prs: Presentation, source_slide):
        layout = source_slide.slide_layout
        new_slide = prs.slides.add_slide(layout)

        source_cSld = source_slide.element.cSld
        new_cSld = new_slide.element.cSld
        bg_element = None
        for child in source_cSld:
            if child.tag.endswith("}bg"):
                bg_element = child
                break

        if bg_element is not None:
            for child in list(new_cSld):
                if child.tag.endswith("}bg"):
                    new_cSld.remove(child)
            cloned_bg = copy.deepcopy(bg_element)
            new_cSld.insert(0, cloned_bg)

        for shape in list(new_slide.shapes):
            new_slide.shapes.element.remove(shape.element)

        for shape in source_slide.shapes:
            cloned_el = copy.deepcopy(shape.element)
            new_slide.shapes._spTree.insert_element_before(cloned_el, "p:extLst")

        for rel_id, rel in source_slide.part.rels.items():
            if "notesSlide" not in rel.reltype and "slideLayout" not in rel.reltype:
                try:
                    new_slide.part.rels.add_relationship(rel.reltype, rel._target, rel.rId)
                except Exception:
                    pass

        return new_slide


# ------------------------------------------------------------------------------
# 6. Data Loader & File Utilities
# ------------------------------------------------------------------------------
class DataLoader:
    @staticmethod
    def load(data_path: str, sheet_name: Optional[Any] = None) -> pd.DataFrame:
        clean_path = str(data_path).strip().strip("'\"")
        if not os.path.exists(clean_path):
            raise FileNotFoundError(f"Data file does not exist: {clean_path}")

        ext = os.path.splitext(clean_path)[1].lower()
        if ext == ".csv":
            df = pd.read_csv(clean_path)
        elif ext in [".xlsx", ".xls"]:
            # If sheet_name is None, default to 0 (first sheet) so pandas returns a DataFrame instead of a dict
            target_sheet = 0 if sheet_name is None else sheet_name
            df = pd.read_excel(clean_path, sheet_name=target_sheet)
            if isinstance(df, dict):
                df = list(df.values())[0]
        else:
            raise ValueError(f"Unsupported file format '{ext}'. Please provide a .csv or .xlsx file.")

        df.columns = [str(c).strip() for c in df.columns]
        return df

    @staticmethod
    def get_sheet_names(excel_path: str) -> List[str]:
        clean_path = str(excel_path).strip().strip("'\"")
        if clean_path.lower().endswith((".xlsx", ".xls")):
            try:
                xl = pd.ExcelFile(clean_path)
                return xl.sheet_names
            except Exception:
                return []
        return []


class FileHelper:
    @staticmethod
    def sanitize_filename(name: str) -> str:
        clean = re.sub(r'[\\/*?:"<>|]', "_", str(name)).strip()
        return clean or "document"

    @staticmethod
    def safe_delete(path: str, retries: int = 5, delay: float = 0.25) -> bool:
        for _ in range(retries):
            try:
                if os.path.exists(path):
                    os.remove(path)
                return True
            except Exception:
                time.sleep(delay)
        return not os.path.exists(path)

    @staticmethod
    def get_export_dir(template_path: str, subfolder: str = "output") -> str:
        parent = os.path.dirname(os.path.abspath(template_path))
        export_dir = os.path.join(parent, subfolder)
        os.makedirs(export_dir, exist_ok=True)
        return export_dir


# ------------------------------------------------------------------------------
# 7. Interactive GUI & Console Dynamic Pickers
# ------------------------------------------------------------------------------
class DynamicPicker:
    @staticmethod
    def select_file(title: str = "Select File", filetypes: Optional[List[Tuple[str, str]]] = None) -> Optional[str]:
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            selected = filedialog.askopenfilename(title=title, filetypes=filetypes or [("All Files", "*.*")])
            root.destroy()
            if selected:
                return os.path.abspath(selected)
        except Exception:
            pass

        val = input(f"\n{title} [Path]: ").strip().strip("'\"")
        return os.path.abspath(val) if val else None

    @staticmethod
    def select_directory(title: str = "Select Folder") -> Optional[str]:
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            selected = filedialog.askdirectory(title=title)
            root.destroy()
            if selected:
                return os.path.abspath(selected)
        except Exception:
            pass

        val = input(f"\n{title} [Folder Path]: ").strip().strip("'\"")
        return os.path.abspath(val) if val else None


def prompt_qr_options(df: pd.DataFrame, prs: Optional[Presentation] = None) -> QRCodeConfig:
    print("\n--- Optional QR Code Configuration ---")
    use_qr = input("Enable dynamic QR code generation? (y/n) [Default: n]: ").strip().lower()
    if use_qr not in ["y", "yes"]:
        return QRCodeConfig(enabled=False)

    cols = list(df.columns)
    detected_col = None
    for c in cols:
        if c.lower() in ["serial", "id", "roll", "registration_id", "url", "qr"]:
            detected_col = c
            break

    default_data = detected_col or (cols[0] if cols else "")
    print(f"Available columns: {cols}")
    data_input = input(f"Enter column name or verification URL template [Default: '{default_data}']: ").strip()
    if not data_input:
        data_input = default_data

    if "{{" in data_input:
        qr_cfg = QRCodeConfig(enabled=True, data_template=data_input)
    else:
        qr_cfg = QRCodeConfig(enabled=True, data_column=data_input)

    has_qr_placeholder = False
    if prs is not None:
        for s in prs.slides:
            for sh in s.shapes:
                if sh.has_text_frame and re.search(r"\{\{\s*qr\s*\}\}", sh.text_frame.text, re.IGNORECASE):
                    has_qr_placeholder = True
                    break

    print("\nChoose QR Code Position:")
    if has_qr_placeholder:
        print("  [1] Replace {{qr}} placeholder shape in template (Recommended - Detected!)")
    else:
        print("  [1] Replace {{qr}} placeholder shape in template (Recommended)")
    print("  [2] Bottom-Right corner")
    print("  [3] Bottom-Left corner")
    print("  [4] Top-Right corner")
    print("  [5] Top-Left corner")
    print("  [6] Center of slide")
    print("  [7] Custom X/Y coordinates")

    def_pos = "1" if has_qr_placeholder else "2"
    pos_choice = input(f"Select position (1-7) [Default: {def_pos}]: ").strip() or def_pos

    pos_map = {
        "1": "placeholder",
        "2": "bottom-right",
        "3": "bottom-left",
        "4": "top-right",
        "5": "top-left",
        "6": "center",
        "7": "custom"
    }
    qr_cfg.position_mode = pos_map.get(pos_choice, "placeholder")

    print("\nChoose QR Code Size:")
    sz_val = input("Enter QR code size in inches [Default: 1.5, or type 'auto' to fit box]: ").strip()
    if sz_val.lower() == "auto":
        qr_cfg.size_mode = "auto"
    elif sz_val:
        try:
            qr_cfg.custom_size_inches = float(sz_val)
            qr_cfg.size_mode = "custom"
        except ValueError:
            qr_cfg.custom_size_inches = 1.5
            qr_cfg.size_mode = "custom"
    else:
        qr_cfg.custom_size_inches = 1.5
        qr_cfg.size_mode = "custom"

    if qr_cfg.position_mode == "custom":
        left_val = input("Enter Left position in inches [Default: 8.0]: ").strip() or "8.0"
        top_val = input("Enter Top position in inches [Default: 5.5]: ").strip() or "5.5"
        try: qr_cfg.custom_left_inches = float(left_val)
        except: qr_cfg.custom_left_inches = 8.0
        try: qr_cfg.custom_top_inches = float(top_val)
        except: qr_cfg.custom_top_inches = 5.5

    print(f"[OK] QR Code configured: Mode='{qr_cfg.position_mode}', Size='{qr_cfg.size_mode}' ({qr_cfg.custom_size_inches} in)")
    return qr_cfg


def prompt_image_options(df: pd.DataFrame, prs: Optional[Presentation] = None) -> ImageConfig:
    print("\n--- Optional Dynamic Image / Photo Configuration ---")
    use_img = input("Enable dynamic image / photo insertion? (y/n) [Default: n]: ").strip().lower()
    if use_img not in ["y", "yes"]:
        return ImageConfig(enabled=False)

    cols = list(df.columns)
    detected_col = None
    for c in cols:
        if c.lower() in ["photo", "image", "picture", "avatar", "photourl", "imageurl", "photo_url"]:
            detected_col = c
            break

    print("\nChoose Image Data Source:")
    print("  [1] From spreadsheet column (containing local file path or web URL)")
    print("  [2] From a folder matching by ID/Serial column (e.g. images/{Serial}.jpg)")
    src_choice = input("Select source (1-2) [Default: 1]: ").strip() or "1"

    img_col = None
    img_folder = None
    id_col = None

    if src_choice == "2":
        img_folder = DynamicPicker.select_directory("Select Folder containing Image files")
        id_col = input(f"Enter column name matching image filenames [Default: 'Serial']: ").strip() or "Serial"
    else:
        def_c = detected_col or (cols[0] if cols else "")
        print(f"Available columns: {cols}")
        img_col = input(f"Enter image column name [Default: '{def_c}']: ").strip() or def_c

    has_img_placeholder = False
    if prs is not None:
        for s in prs.slides:
            for sh in s.shapes:
                if sh.has_text_frame and re.search(r"\{\{\s*(image|photo|picture)\s*\}\}", sh.text_frame.text, re.IGNORECASE):
                    has_img_placeholder = True
                    break

    print("\nChoose Image Position:")
    if has_img_placeholder:
        print("  [1] Replace {{image}} / {{photo}} placeholder box in template (Recommended - Detected!)")
    else:
        print("  [1] Replace {{image}} / {{photo}} placeholder box in template (Recommended)")
    print("  [2] Top-Right corner")
    print("  [3] Top-Left corner")
    print("  [4] Bottom-Right corner")
    print("  [5] Bottom-Left corner")
    print("  [6] Custom X/Y coordinates")

    def_p = "1" if has_img_placeholder else "2"
    p_choice = input(f"Select position (1-6) [Default: {def_p}]: ").strip() or def_p
    p_map = {
        "1": "placeholder",
        "2": "top-right",
        "3": "top-left",
        "4": "bottom-right",
        "5": "bottom-left",
        "6": "custom"
    }
    pos_mode = p_map.get(p_choice, "placeholder")

    print("\nChoose Image Shape / Corner Radius:")
    print("  [1] Square / Rectangle (Sharp 90-degree corners - Standard)")
    print("  [2] Rounded Rectangle (Smooth rounded corners)")
    print("  [3] Circle / Oval (Circular avatar / passport style)")
    s_choice = input("Select shape (1-3) [Default: 1]: ").strip() or "1"
    shape_map = {"1": "square", "2": "rounded", "3": "circle"}
    shape_mode = shape_map.get(s_choice, "square")

    corner_radius = 25
    if shape_mode == "rounded":
        r_val = input("Enter corner radius in pixels (e.g. 15, 25, 40) [Default: 25]: ").strip() or "25"
        try: corner_radius = int(r_val)
        except: corner_radius = 25

    print("\nChoose Image Border / Stroke:")
    b_val = input("Enter border stroke width in pixels (0 for none) [Default: 0]: ").strip() or "0"
    try: border_width = int(b_val)
    except: border_width = 0

    border_color = "#FFFFFF"
    if border_width > 0:
        print("  Color presets: [1] White (#FFFFFF), [2] Black (#000000), [3] Blue (#0066CC), [4] Navy (#1E3A8A), [5] Gold (#D97706), [6] Red (#E11D48)")
        c_choice = input("Select color (1-6) or enter Hex [Default: 1]: ").strip() or "1"
        color_map = {
            "1": "#FFFFFF",
            "2": "#000000",
            "3": "#0066CC",
            "4": "#1E3A8A",
            "5": "#D97706",
            "6": "#E11D48"
        }
        if c_choice in color_map:
            border_color = color_map[c_choice]
        elif c_choice.startswith("#") or len(c_choice) in [3, 6]:
            border_color = c_choice if c_choice.startswith("#") else f"#{c_choice}"
        else:
            border_color = "#FFFFFF"

    img_cfg = ImageConfig(
        enabled=True,
        data_column=img_col,
        image_folder=img_folder,
        id_column=id_col,
        position_mode=pos_mode,
        size_mode="auto" if pos_mode == "placeholder" else "custom",
        shape_mode=shape_mode,
        corner_radius=corner_radius,
        border_width=border_width,
        border_color=border_color
    )

    if pos_mode == "custom":
        l_val = input("Enter Left position in inches [Default: 7.5]: ").strip() or "7.5"
        t_val = input("Enter Top position in inches [Default: 1.5]: ").strip() or "1.5"
        try: img_cfg.custom_left_inches = float(l_val)
        except: img_cfg.custom_left_inches = 7.5
        try: img_cfg.custom_top_inches = float(t_val)
        except: img_cfg.custom_top_inches = 1.5

    style_summary = f"Shape={img_cfg.shape_mode.capitalize()}"
    if img_cfg.shape_mode == "rounded":
        style_summary += f" (Radius={img_cfg.corner_radius}px)"
    if img_cfg.border_width > 0:
        style_summary += f", Border={img_cfg.border_width}px ({img_cfg.border_color})"
    else:
        style_summary += ", No Border"

    print(f"[OK] Image engine configured: Mode='{img_cfg.position_mode}', {style_summary}")
    return img_cfg


print("[+] AutoCrat Pro Core Engine (with QR & Image Support) loaded successfully!")


# ==============================================================================
# Cell 3: Interactive Universal Dashboard (All-in-One Console)
# ==============================================================================
# Run this cell to launch the interactive AutoCrat Pro menu.
# You can choose any mode, browse files via Windows GUI dialogs, and monitor execution.

# ------------------------------------------------------------------------------
# Concurrency Worker Engine for Individual Document Processing (Mode 1)
# ------------------------------------------------------------------------------
def process_single_individual_pptx(item: Dict[str, Any]) -> Tuple[bool, int, str, str, str]:
    """
    Worker task: Concurrently generates 1 individual PPTX file (Text + QR + Image).
    Executed across concurrent worker threads without process spawn overhead or COM crashes.
    Returns (success, idx, temp_pptx, pdf_path, error_message).
    """
    idx = item["idx"]
    row_dict = item["row_dict"]
    pptx_path = item["pptx_path"]
    export_dir = item["export_dir"]
    id_cols = item.get("id_cols", [])
    qr_config = item.get("qr_config")
    img_config = item.get("img_config")

    temp_pptx = ""
    temp_qr_img = ""
    try:
        row = pd.Series(row_dict)
        mapping = PlaceholderEngine.build_mapping(row)

        if id_cols:
            parts = [str(row[c]).strip() for c in id_cols if c in row and str(row[c]).strip()]
            base_name = FileHelper.sanitize_filename("-".join(parts) if parts else f"doc_{idx+1}")
        else:
            base_name = f"doc_{idx+1}"

        pdf_path = os.path.join(export_dir, f"{base_name}.pdf")
        temp_pptx = os.path.join(export_dir, f"{base_name}_temp_{idx+1}.pptx")
        temp_qr_img = os.path.join(export_dir, f"{base_name}_qr_{idx+1}.png")

        prs = Presentation(pptx_path)

        # 1. Fill text placeholders
        for slide in prs.slides:
            PlaceholderEngine.process_slide(slide, mapping)

        # 2. Embed QR code if enabled
        if qr_config and qr_config.enabled:
            payload = qr_config.resolve_payload(mapping, row)
            if payload and QRCodeHandler.generate_image(payload, temp_qr_img, qr_config.box_size, qr_config.border):
                for slide in prs.slides:
                    QRCodeHandler.embed_in_slide(slide, prs, temp_qr_img, qr_config)
                FileHelper.safe_delete(temp_qr_img)

        # 3. Embed Image / Photo if enabled
        if img_config and img_config.enabled:
            img_path, is_temp = ImageHandler.resolve_image_path(img_config, row, mapping, export_dir)
            if img_path:
                for slide in prs.slides:
                    ImageHandler.embed_in_slide(slide, prs, img_path, img_config)
                if is_temp:
                    FileHelper.safe_delete(img_path)

        prs.save(temp_pptx)
        return (True, idx, temp_pptx, pdf_path, "")
    except Exception as err:
        if temp_pptx:
            FileHelper.safe_delete(temp_pptx)
        if temp_qr_img:
            FileHelper.safe_delete(temp_qr_img)
        return (False, idx, "", "", str(err))


def run_mode_1_individual():
    print("\n" + "=" * 60)
    print("  MODE 1: Individual Document Generator (1 PDF per Row)")
    print("=" * 60)
    
    # 1. Data File
    data_path = DynamicPicker.select_file("Select Excel (.xlsx) or CSV file", [("Spreadsheets", "*.xlsx;*.xls;*.csv"), ("All Files", "*.*")])
    if not data_path or not os.path.exists(data_path):
        print("[!] No valid data file selected. Aborting.")
        return
        
    sheets = DataLoader.get_sheet_names(data_path)
    selected_sheet = sheets[0] if sheets else None
    if len(sheets) > 1:
        print(f"\nFound sheets: {sheets}")
        choice = input(f"Enter sheet name or index [Default: {sheets[0]}]: ").strip()
        if choice in sheets:
            selected_sheet = choice
        elif choice.isdigit() and 1 <= int(choice) <= len(sheets):
            selected_sheet = sheets[int(choice) - 1]
            
    df = DataLoader.load(data_path, selected_sheet)
    print(f"[OK] Loaded {len(df)} rows. Columns: {list(df.columns)}")

    # 2. Template
    pptx_path = DynamicPicker.select_file("Select PowerPoint Template (.pptx)", [("PowerPoint Presentation", "*.pptx")])
    if not pptx_path or not os.path.exists(pptx_path):
        print("[!] No valid template selected. Aborting.")
        return

    prs_inspect = Presentation(pptx_path)

    # 3. Filename Columns
    print(f"\nColumns available: {list(df.columns)}")
    cols_input = input("Enter column name(s) for output filename (comma-separated, Enter for row_1): ").strip()
    id_cols = [c.strip() for c in cols_input.split(",") if c.strip() in df.columns]
    
    # 4. Optional QR Code Configuration
    qr_config = prompt_qr_options(df, prs_inspect)

    # 5. Optional Image / Photo Configuration
    img_config = prompt_image_options(df, prs_inspect)

    # 6. Confirmation & Output
    export_dir = FileHelper.get_export_dir(pptx_path)
    print(f"\n[DIR] Files will be saved in: {export_dir}")
    proceed = input(f"Generate {len(df)} individual PDF files? (y/n) [Default: y]: ").strip().lower()
    if proceed not in ["", "y", "yes"]:
        print("Cancelled.")
        return

    num_workers = min(8, max(1, len(df)))
    success_count = 0
    start_time = time.time()
    print(f"\n[*] Processing {len(df)} documents with {num_workers} concurrent workers...")

    items = []
    for idx, row in df.iterrows():
        items.append({
            "idx": idx,
            "row_dict": row.to_dict(),
            "pptx_path": pptx_path,
            "export_dir": export_dir,
            "id_cols": id_cols,
            "qr_config": qr_config,
            "img_config": img_config
        })

    converter = PPTXToPDFConverter()
    if not converter.start():
        print("[!] PowerPoint COM engine unavailable; PDF generation aborted.")
        return

    try:
        with concurrent.futures.ThreadPoolExecutor(max_workers=num_workers) as executor:
            futures = [executor.submit(process_single_individual_pptx, item) for item in items]
            for fut in tqdm(concurrent.futures.as_completed(futures), total=len(items), desc=f"Generating PDFs ({num_workers} Workers)"):
                ok, idx, temp_pptx, pdf_path, err = fut.result()
                if ok and temp_pptx:
                    if converter.convert(temp_pptx, pdf_path):
                        success_count += 1
                    FileHelper.safe_delete(temp_pptx)
                elif err:
                    print(f"\n    [ERR] Row {idx+1} failed: {err}")
    finally:
        converter.stop()

    elapsed = round(time.time() - start_time, 2)
    print("\n" + "=" * 60)
    print(f"  [COMPLETED] {success_count}/{len(df)} PDFs successfully generated in {elapsed}s")
    print(f"  [SAVED IN]  {export_dir}")
    print("=" * 60)


def run_mode_2_merged():
    print("\n" + "=" * 60)
    print("  MODE 2: Merged Master Document Generator (All Rows in 1 File)")
    print("=" * 60)
    
    data_path = DynamicPicker.select_file("Select Excel (.xlsx) or CSV file", [("Spreadsheets", "*.xlsx;*.xls;*.csv"), ("All Files", "*.*")])
    if not data_path or not os.path.exists(data_path):
        return
        
    sheets = DataLoader.get_sheet_names(data_path)
    selected_sheet = sheets[0] if sheets else None
    if len(sheets) > 1:
        print(f"\nFound sheets: {sheets}")
        choice = input(f"Enter sheet name [Default: {sheets[0]}]: ").strip()
        if choice in sheets:
            selected_sheet = choice
            
    df = DataLoader.load(data_path, selected_sheet)
    print(f"[OK] Loaded {len(df)} rows.")

    pptx_path = DynamicPicker.select_file("Select PowerPoint Template (.pptx)", [("PowerPoint Presentation", "*.pptx")])
    if not pptx_path or not os.path.exists(pptx_path):
        return

    prs_inspect = Presentation(pptx_path)

    default_name = f"{os.path.splitext(os.path.basename(pptx_path))[0]}_merged"
    output_name = input(f"\nEnter final output file name [Default: {default_name}]: ").strip()
    if not output_name:
        output_name = default_name
    output_name = FileHelper.sanitize_filename(output_name)

    qr_config = prompt_qr_options(df, prs_inspect)
    img_config = prompt_image_options(df, prs_inspect)

    export_pdf_choice = input("\nExport as PDF in addition to PPTX? (y/n) [Default: y]: ").strip().lower()
    export_pdf = export_pdf_choice in ["", "y", "yes"]

    export_dir = FileHelper.get_export_dir(pptx_path)
    prs = Presentation(pptx_path)
    slides_per_row = len(prs.slides)
    num_rows = len(df)

    print(f"\n[*] Duplicating template slides for {num_rows} recipients ({num_rows * slides_per_row} slides total)...")
    template_slides = [prs.slides[i] for i in range(slides_per_row)]
    
    for r in tqdm(range(1, num_rows), desc="Cloning Slides"):
        for t_slide in template_slides:
            SlideDuplicator.duplicate_slide(prs, t_slide)

    print("[*] Populating placeholders, QR codes & photos in merged presentation...")
    for idx, row in tqdm(df.iterrows(), total=num_rows, desc="Filling Data"):
        mapping = PlaceholderEngine.build_mapping(row)
        temp_qr_img = os.path.join(export_dir, f"_temp_qr_{idx}.png")
        qr_ready = False

        if qr_config.enabled:
            payload = qr_config.resolve_payload(mapping, row)
            if payload:
                qr_ready = QRCodeHandler.generate_image(payload, temp_qr_img, qr_config.box_size, qr_config.border)

        img_path, is_temp = ImageHandler.resolve_image_path(img_config, row, mapping, export_dir) if img_config.enabled else (None, False)

        for offset in range(slides_per_row):
            target_slide = prs.slides[idx * slides_per_row + offset]
            PlaceholderEngine.process_slide(target_slide, mapping)
            if qr_ready:
                QRCodeHandler.embed_in_slide(target_slide, prs, temp_qr_img, qr_config)
            if img_path:
                ImageHandler.embed_in_slide(target_slide, prs, img_path, img_config)

        if qr_ready:
            FileHelper.safe_delete(temp_qr_img)
        if is_temp and img_path:
            FileHelper.safe_delete(img_path)

    out_pptx = os.path.join(export_dir, f"{output_name}.pptx")
    prs.save(out_pptx)
    print(f"  [OK] Saved merged PowerPoint: {output_name}.pptx")

    if export_pdf:
        out_pdf = os.path.join(export_dir, f"{output_name}.pdf")
        converter = PPTXToPDFConverter()
        if converter.start():
            print("  [*] Converting merged presentation to master PDF...")
            converted = converter.convert(out_pptx, out_pdf)
            converter.stop()
            if converted:
                print(f"  [OK] Saved master PDF: {output_name}.pdf")
            else:
                print("  [!] PDF conversion failed.")
        else:
            print("  [!] PowerPoint COM unavailable for PDF conversion.")

    print("\n" + "=" * 60)
    print(f"  [COMPLETED] Master documents saved in: {export_dir}")
    print("=" * 60)


def run_mode_3_batch_directory():
    print("\n" + "=" * 60)
    print("  MODE 3: Batch Directory Auto-Processor (Paired XLSX + PPTX)")
    print("=" * 60)
    
    target_dir = DynamicPicker.select_directory("Select directory containing matching XLSX and PPTX files")
    if not target_dir or not os.path.isdir(target_dir):
        print("[!] No directory selected. Aborting.")
        return

    export_dir = os.path.join(target_dir, "output")
    os.makedirs(export_dir, exist_ok=True)

    files = os.listdir(target_dir)
    xlsx_files = [f for f in files if f.endswith((".xlsx", ".xls")) and not f.startswith("~") and not f.startswith("all_")]

    pairs = []
    for xf in xlsx_files:
        base = os.path.splitext(xf)[0]
        pf = f"{base}.pptx"
        if os.path.exists(os.path.join(target_dir, pf)):
            pairs.append((xf, pf, base))

    if not pairs:
        print(f"[!] No matching .xlsx and .pptx pairs found in {target_dir}")
        return

    print(f"\nFound {len(pairs)} matching pair(s):")
    for xf, pf, b in pairs:
        print(f"  - {xf} <--> {pf}")

    # Inspect first batch pair to guide QR and Image options
    sample_df = DataLoader.load(os.path.join(target_dir, pairs[0][0]))
    sample_prs = Presentation(os.path.join(target_dir, pairs[0][1]))

    qr_cfg = prompt_qr_options(sample_df, sample_prs)
    img_cfg = prompt_image_options(sample_df, sample_prs)

    export_pdf_choice = input("\nAlso convert each batch to PDF? (y/n) [Default: y]: ").strip().lower()
    export_pdf = export_pdf_choice in ["", "y", "yes"]

    converter = PPTXToPDFConverter() if export_pdf else None
    if converter and not converter.start():
        print("[WARN] PowerPoint COM unavailable; proceeding with PPTX-only.")
        converter = None

    try:
        for xf, pf, base_name in tqdm(pairs, desc="Processing Batches"):
            xlsx_path = os.path.join(target_dir, xf)
            pptx_path = os.path.join(target_dir, pf)
            
            df = DataLoader.load(xlsx_path)
            if len(df) == 0:
                continue

            prs = Presentation(pptx_path)
            slides_per_row = len(prs.slides)
            template_slides = [prs.slides[i] for i in range(slides_per_row)]

            for r in range(1, len(df)):
                for t_slide in template_slides:
                    SlideDuplicator.duplicate_slide(prs, t_slide)

            for idx, row in df.iterrows():
                mapping = PlaceholderEngine.build_mapping(row)
                temp_qr = os.path.join(export_dir, f"_temp_qr_{idx}.png")
                qr_ready = False

                if qr_cfg.enabled:
                    payload = qr_cfg.resolve_payload(mapping, row)
                    if payload:
                        qr_ready = QRCodeHandler.generate_image(payload, temp_qr, qr_cfg.box_size, qr_cfg.border)

                img_path, is_temp = ImageHandler.resolve_image_path(img_cfg, row, mapping, export_dir) if img_cfg.enabled else (None, False)

                for offset in range(slides_per_row):
                    target_slide = prs.slides[idx * slides_per_row + offset]
                    PlaceholderEngine.process_slide(target_slide, mapping)
                    if qr_ready:
                        QRCodeHandler.embed_in_slide(target_slide, prs, temp_qr, qr_cfg)
                    if img_path:
                        ImageHandler.embed_in_slide(target_slide, prs, img_path, img_cfg)

                if qr_ready:
                    FileHelper.safe_delete(temp_qr)
                if is_temp and img_path:
                    FileHelper.safe_delete(img_path)

            out_pptx = os.path.join(export_dir, f"{base_name}_merged.pptx")
            prs.save(out_pptx)
            
            if converter:
                out_pdf = os.path.join(export_dir, f"{base_name}_merged.pdf")
                converter.convert(out_pptx, out_pdf)
    finally:
        if converter:
            converter.stop()

    print("\n" + "=" * 60)
    print(f"  [COMPLETED] All batch files saved in: {export_dir}")
    print("=" * 60)


def run_mode_4_pdf_merger():
    print("\n" + "=" * 60)
    print("  MODE 4: Standalone PDF Merger Utility")
    print("=" * 60)
    
    folder_path = DynamicPicker.select_directory("Select folder containing PDF files to merge")
    if not folder_path or not os.path.isdir(folder_path):
        print("[!] Invalid folder path.")
        return

    pdf_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]
    pdf_files.sort(key=lambda x: [int(t) if t.isdigit() else t.lower() for t in re.split(r'(\d+)', os.path.basename(x))])

    if not pdf_files:
        print("[!] No PDF files found in the selected folder.")
        return

    print(f"\nFound {len(pdf_files)} PDF file(s).")
    default_out = "Merged_Document.pdf"
    out_name = input(f"Enter merged output filename [Default: {default_out}]: ").strip()
    if not out_name:
        out_name = default_out
    if not out_name.lower().endswith(".pdf"):
        out_name += ".pdf"

    output_path = os.path.join(folder_path, out_name)
    merger = PdfWriter()

    print("[*] Merging PDF files...")
    for pdf_path in tqdm(pdf_files, desc="Merging PDFs"):
        if os.path.abspath(pdf_path) == os.path.abspath(output_path):
            continue
        merger.append(pdf_path)

    with open(output_path, "wb") as f_out:
        merger.write(f_out)
    merger.close()

    print("\n" + "=" * 60)
    print(f"  [OK] Successfully merged {len(pdf_files)} PDFs!")
    print(f"  [SAVED AS] {output_path}")
    print("=" * 60)


def run_mode_5_validator():
    print("\n" + "=" * 60)
    print("  MODE 5: Pre-Flight Template & Data Validator")
    print("=" * 60)
    
    data_path = DynamicPicker.select_file("Select Excel or CSV Data File", [("Spreadsheets", "*.xlsx;*.xls;*.csv")])
    if not data_path: return
    pptx_path = DynamicPicker.select_file("Select PowerPoint Template (.pptx)", [("PowerPoint", "*.pptx")])
    if not pptx_path: return

    sheets = DataLoader.get_sheet_names(data_path)
    selected_sheet = sheets[0] if sheets else None
    if len(sheets) > 1:
        print(f"\nFound sheets: {sheets}")
        choice = input(f"Enter sheet name or index [Default: {sheets[0]}]: ").strip()
        if choice in sheets:
            selected_sheet = choice
        elif choice.isdigit() and 1 <= int(choice) <= len(sheets):
            selected_sheet = sheets[int(choice) - 1]

    df = DataLoader.load(data_path, selected_sheet)
    prs = Presentation(pptx_path)

    template_tokens = PlaceholderEngine.extract_placeholders(prs)
    excel_cols = list(df.columns)
    excel_norm = {PlaceholderEngine.normalize_key(c): c for c in excel_cols}

    print("\n" + "-" * 50)
    print(f"  Spreadsheet Columns ({len(excel_cols)}): {excel_cols}")
    print(f"  Template Placeholders ({len(template_tokens)}): {template_tokens}")
    print("-" * 50)

    matched = []
    missing_in_data = []

    for token in template_tokens:
        tok_low = token.lower()
        if tok_low in ["qr", "image", "photo", "picture"]:
            continue
        norm_t = PlaceholderEngine.normalize_key(token)
        if token in df.columns:
            matched.append((token, token, "Exact Match"))
        elif norm_t in excel_norm:
            matched.append((token, excel_norm[norm_t], "Fuzzy Normalized Match"))
        else:
            missing_in_data.append(token)

    print(f"\n[+] Successfully Matched ({len(matched)}):")
    for t, col, match_type in matched:
        print(f"    - '{{{{{t}}}}}' --> Column '{col}' ({match_type})")

    tok_set = [t.lower() for t in template_tokens]
    if "qr" in tok_set:
        print("    - '{{qr}}' --> Dynamic QR Code Placeholder Shape (Detected!)")
    if any(t in tok_set for t in ["image", "photo", "picture"]):
        print("    - '{{image}}' / '{{photo}}' --> Dynamic Photo Placeholder Shape (Detected!)")

    if missing_in_data:
        print(f"\n[!] Missing in Spreadsheet ({len(missing_in_data)}):")
        for m in missing_in_data:
            print(f"    - '{{{{{m}}}}}' (No matching column in dataset)")
    else:
        print("\n[+] All template placeholders are 100% covered by the dataset!")


# Main Dashboard (Runs Once & Stops Immediately)
def main_dashboard():
    print("\n" + "=" * 65)
    print("  *** AutoCrat Pro -- Universal Document Automation Suite ***")
    print("=" * 65)
    print("  [1] Mode 1: Individual Documents (1 PDF per Row, with QR & Photo)")
    print("  [2] Mode 2: Merged Master Document (All Rows in 1 PPTX & PDF)")
    print("  [3] Mode 3: Batch Directory Auto-Processor (Paired XLSX + PPTX)")
    print("  [4] Mode 4: Standalone PDF Merger Utility")
    print("  [5] Mode 5: Pre-Flight Template & Data Validator (Dry Run)")
    print("  [0] Exit")
    print("=" * 65)
    
    choice = input("Select an option (0-5): ").strip()
    if choice == "1":
        run_mode_1_individual()
    elif choice == "2":
        run_mode_2_merged()
    elif choice == "3":
        run_mode_3_batch_directory()
    elif choice == "4":
        run_mode_4_pdf_merger()
    elif choice == "5":
        run_mode_5_validator()
    elif choice in ["0", "q", "exit"]:
        print("\nExiting AutoCrat Pro. Have a productive day!")
    else:
        print("[!] Invalid option. Exiting.")
    return

# Launch Dashboard
if __name__ == "__main__":
    main_dashboard()

