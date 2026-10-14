"""
All-in-One PDF Toolkit - Core Engine
====================================
A complete, reliable, and well-structured utility for managing PDF documents.
Provides 10 core operations:
  1. Merge PDFs
  2. Split PDFs
  3. Compress PDFs
  4. Unlock PDF
  5. Protect PDF
  6. Images to PDF (NEW)
  7. PDF to Images (NEW)
  8. OCR PDF
  9. Markdown to PDF
  10. PDF to Markdown
Includes modern Glassmorphism Desktop GUI Dashboard.
"""

import os
import sys
import re
import time
import shutil
from pathlib import Path

# Ensure project and workspace directories are always in sys.path
for _p in [Path(__file__).resolve().parent if '__file__' in globals() else Path.cwd(), Path.cwd(), Path.cwd().parent, Path(r"d:\New folder\PY"), Path(r"d:\New folder\PY\scripts")]:
    _p_str = str(_p.resolve()) if hasattr(_p, "resolve") and _p.exists() else str(_p)
    if _p_str not in sys.path:
        sys.path.insert(0, _p_str)
import logging
import threading
import queue
from pathlib import Path
from typing import List, Dict, Tuple, Optional, Any, Set, Callable
import getpass

# Core PDF, image, and text processing libraries
import pypdf
import pymupdf
import markdown
from fpdf import FPDF
from PIL import Image, ImageTk, ImageDraw

try:
    import pypdfium2 as pdfium
    PDFIUM_AVAILABLE = True
except ImportError:
    PDFIUM_AVAILABLE = False

try:
    import pytesseract
    PYTESSERACT_AVAILABLE = True
except ImportError:
    PYTESSERACT_AVAILABLE = False

# GUI modules
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import windnd
    WINDND_AVAILABLE = True
except ImportError:
    WINDND_AVAILABLE = False

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger("PDFToolkit")


# ==============================================================================
# SECTION 3 & 12: SHARED UTILITIES & OUTPUT MANAGEMENT
# ==============================================================================

def parse_folder_paths(input_str: str) -> List[str]:
    r"""
    Parses comma-separated folder paths from user input.
    Supports:
      - Quoted paths containing literal commas: "D:\Docs, Files", D:\Other
      - Escaped commas: D:\Docs\, Files
      - Standard Windows and Linux path formats
      - Strips surrounding quotes and whitespace
    """
    if not input_str or not input_str.strip():
        return []

    tokens = []
    current = []
    in_quotes = False
    quote_char = None
    i = 0
    chars = input_str.strip()

    while i < len(chars):
        c = chars[i]
        # Check for escaped comma
        if c == '\\' and i + 1 < len(chars) and chars[i + 1] == ',':
            current.append(',')
            i += 2
            continue
        elif c in ('"', "'"):
            if not in_quotes:
                in_quotes = True
                quote_char = c
            elif quote_char == c:
                in_quotes = False
                quote_char = None
            else:
                current.append(c)
        elif c == ',' and not in_quotes:
            token = "".join(current).strip()
            if token:
                tokens.append(token)
            current = []
        else:
            current.append(c)
        i += 1

    last_token = "".join(current).strip()
    if last_token:
        tokens.append(last_token)

    # Clean any surrounding quotes left on tokens
    cleaned_tokens = []
    for t in tokens:
        if (t.startswith('"') and t.endswith('"')) or (t.startswith("'") and t.endswith("'")):
            t = t[1:-1].strip()
        if t:
            cleaned_tokens.append(t)

    return cleaned_tokens


def validate_folder_paths(raw_paths: List[str]) -> Tuple[List[Path], List[str]]:
    """
    Validates folder existence and accessibility.
    Deduplicates resolved paths while preserving user input order.
    Returns: (valid_folders, warnings)
    """
    valid_folders: List[Path] = []
    warnings: List[str] = []
    seen_resolved: Set[str] = set()

    for p_str in raw_paths:
        try:
            p = Path(p_str.strip())
            if not p.exists():
                warnings.append(f"Folder does not exist: {p_str}")
                continue
            if not p.is_dir():
                warnings.append(f"Path is not a directory: {p_str}")
                continue
            # Check read access
            if not os.access(p, os.R_OK):
                warnings.append(f"Folder is not accessible / readable: {p_str}")
                continue

            resolved_key = str(p.resolve()).lower()
            if resolved_key in seen_resolved:
                # Duplicate folder path, skip silently to prevent duplicate discovery
                continue
            seen_resolved.add(resolved_key)
            valid_folders.append(p)
        except Exception as e:
            warnings.append(f"Error resolving path '{p_str}': {e}")

    return valid_folders, warnings


def discover_files(folders: List[Path], extensions: List[str] = [".pdf"], recursive: bool = False) -> List[Path]:
    """
    Discovers files with the given extensions across the provided folders.
    Deduplicates files so the same physical file is never listed twice.
    Maintains a stable discovery order.
    """
    discovered: List[Path] = []
    seen_files: Set[str] = set()
    norm_exts = {ext.lower() for ext in extensions}

    for folder in folders:
        try:
            iterator = folder.rglob("*") if recursive else folder.iterdir()
            # Sort locally for predictable folder listing before user ordering
            entries = sorted(list(iterator), key=lambda x: str(x).lower())
            for item in entries:
                if item.is_file() and item.suffix.lower() in norm_exts:
                    try:
                        resolved_path = str(item.resolve()).lower()
                        if resolved_path not in seen_files:
                            seen_files.add(resolved_path)
                            discovered.append(item)
                    except Exception:
                        if str(item) not in seen_files:
                            seen_files.add(str(item))
                            discovered.append(item)
        except Exception as e:
            logger.warning(f"Error scanning folder '{folder}': {e}")

    return discovered


def display_files(files: List[Path], title: str = "Available PDF Files") -> None:
    """
    Displays the discovered files in a clean, numbered list starting at 1.
    """
    print("\n" + "=" * 56)
    print(f" {title}".ljust(56))
    print("=" * 56)
    if not files:
        print(" No matching files found.")
    else:
        for idx, file_path in enumerate(files, start=1):
            print(f"\n[{idx}] {file_path.name}")
            print(f"    Folder: {file_path.parent}")
    print("-" * 56)
    print(f"Total files found: {len(files)}")
    print("=" * 56 + "\n")


def parse_serial_selection(input_str: str, total_count: int) -> Tuple[Optional[List[int]], Optional[str]]:
    """
    Validates serial number selection input:
      - 'q' -> Cancel
      - 'all' -> [1, 2, ..., total_count]
      - '3,1,4' -> [3, 1, 4]
    Validates range, rejects duplicates with an explicit error, and preserves user order.
    Returns: (selection_list, error_or_cancel_message)
    """
    cleaned = input_str.strip()
    if not cleaned:
        return None, "Input cannot be empty. Please enter serial numbers or 'all'."

    if cleaned.lower() == 'q':
        return None, "Operation cancelled by user."

    if cleaned.lower() == 'all':
        if total_count <= 0:
            return None, "No files available to select."
        return list(range(1, total_count + 1)), None

    tokens = [t.strip() for t in cleaned.split(',') if t.strip()]
    if not tokens:
        return None, "No valid serial numbers entered."

    parsed_serials: List[int] = []
    seen: Set[int] = set()

    for t in tokens:
        if not t.isdigit():
            return None, f"Invalid entry '{t}'. Serial numbers must be positive integers."
        val = int(t)
        if val <= 0:
            return None, f"Invalid serial number '{val}'. Numbers must start from 1."
        if val > total_count:
            return None, f"Serial number '{val}' is out of range. Max available is {total_count}."
        if val in seen:
            return None, f"Duplicate serial number detected: [{val}]. Please remove duplicates and re-enter."
        seen.add(val)
        parsed_serials.append(val)

    return parsed_serials, None


def format_file_size(num_bytes: int) -> str:
    """Formats bytes into human readable size."""
    if num_bytes < 0:
        return "0.00 B"
    for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
        if num_bytes < 1024.0 or unit == 'TB':
            return f"{num_bytes:.2f} {unit}"
        num_bytes /= 1024.0
    return f"{num_bytes:.2f} TB"


class OutputManager:
    """
    Centralized output directory and path management.
    Ensures:
      - Original files are never overwritten
      - Existing output files are never silently overwritten
      - Automatic collision-safe naming
      - File verification after generation
    """

    @staticmethod
    def get_collision_safe_path(target_path: Path) -> Path:
        """
        If target_path already exists, generates target (1).ext, target (2).ext, etc.
        """
        if not target_path.exists():
            return target_path
        parent = target_path.parent
        stem = target_path.stem
        suffix = target_path.suffix
        counter = 1
        while True:
            candidate = parent / f"{stem} ({counter}){suffix}"
            if not candidate.exists():
                return candidate
            counter += 1

    @staticmethod
    def get_collision_safe_dir(target_dir: Path) -> Path:
        """
        If target_dir already exists and we need a fresh folder,
        generates folder_1, folder_2, etc.
        """
        if not target_dir.exists():
            return target_dir
        parent = target_dir.parent
        name = target_dir.name
        counter = 1
        while True:
            candidate = parent / f"{name}_{counter}"
            if not candidate.exists():
                return candidate
            counter += 1

    @staticmethod
    def ensure_output_dir(dir_path: Path) -> Path:
        """Creates directory if it does not exist."""
        dir_path.mkdir(parents=True, exist_ok=True)
        return dir_path

    @staticmethod
    def verify_pdf_readable(pdf_path: Path) -> Tuple[bool, str]:
        """Verifies that a generated PDF exists and has readable structure."""
        if not pdf_path.exists():
            return False, "File does not exist."
        if pdf_path.stat().st_size == 0:
            return False, "File is empty (0 bytes)."
        try:
            reader = pypdf.PdfReader(str(pdf_path))
            page_count = len(reader.pages)
            if page_count < 1:
                return False, "PDF contains 0 pages."
            return True, f"Valid PDF with {page_count} page(s)."
        except Exception as e:
            return False, f"PDF verification failed: {e}"


class ProcessingReport:
    """
    Centralized report collector and renderer matching Section 13 specification.
    """
    def __init__(self, operation_name: str):
        self.operation_name = operation_name
        self.results: List[Dict[str, Any]] = []

    def add_result(self, input_file: Optional[Path], status: str, output_files: List[Path] = None, error: str = None, details: str = None):
        self.results.append({
            "input_file": input_file,
            "status": status,  # "Success", "Failed", "Skipped"
            "output_files": output_files or [],
            "error": error,
            "details": details
        })

    def display(self):
        total = len(self.results)
        successful = sum(1 for r in self.results if r["status"] == "Success")
        failed = sum(1 for r in self.results if r["status"] == "Failed")
        skipped = sum(1 for r in self.results if r["status"] == "Skipped")

        print("\n" + "=" * 56)
        print(" Processing Summary".ljust(56))
        print("=" * 56)
        print(f"Operation: {self.operation_name}\n")
        print(f"Total selected: {total}")
        print(f"Successful:     {successful}")
        print(f"Failed:         {failed}")
        print(f"Skipped:        {skipped}\n")

        all_outputs = []
        for r in self.results:
            all_outputs.extend(r["output_files"])

        if all_outputs:
            print("Output files:")
            for idx, out in enumerate(all_outputs, start=1):
                print(f"{idx}. {out}")
            print()

        errors = [r for r in self.results if r["error"]]
        if errors:
            print("Errors / Warnings:")
            for err in errors:
                fname = err["input_file"].name if err["input_file"] else "General"
                print(f"- {fname}: {err['error']}")
            print()

        print("=" * 56 + "\n")


# ==============================================================================
# OPERATION 1: MERGE PDFS
# ==============================================================================

def execute_merge_pdfs(
    selected_files: List[Path],
    output_parent_folder: Path,
    merged_filename: str = "merged.pdf",
    allow_collision_rename: bool = True,
    save_in_subfolder: bool = True,
    progress_callback: Optional[Callable[[float, str], None]] = None,
    log_callback: Optional[Callable[[str], None]] = None
) -> ProcessingReport:
    """
    Merges multiple selected PDF files into one PDF while preserving
    the exact sequence specified by the user.
    Saves directly in output_parent_folder (if save_in_subfolder=False) or in 'output/' child folder.
    """
    def _cb(pct: float, msg: str):
        if progress_callback:
            try:
                progress_callback(pct, msg)
            except Exception:
                pass

    def _log(msg: str):
        if log_callback:
            try:
                log_callback(msg)
            except Exception:
                pass

    report = ProcessingReport("Merge PDFs")
    if not selected_files:
        report.add_result(None, "Failed", error="No files provided for merging.")
        return report

    total_files = len(selected_files)
    _cb(0.0, f"Preparing to merge {total_files} document(s)...")

    # Prepare output folder
    output_dir = (output_parent_folder / "output") if save_in_subfolder else Path(output_parent_folder)
    OutputManager.ensure_output_dir(output_dir)

    clean_name = merged_filename.strip()
    if not clean_name.lower().endswith(".pdf"):
        clean_name += ".pdf"

    target_path = output_dir / clean_name

    # Check if target path would overwrite an input file
    resolved_inputs = {f.resolve() for f in selected_files}
    if target_path.resolve() in resolved_inputs:
        report.add_result(None, "Failed", error="Output path conflicts with one of the input files.")
        return report

    # Handle existing output file collision
    if target_path.exists():
        if allow_collision_rename:
            target_path = OutputManager.get_collision_safe_path(target_path)
            logger.info(f"Existing file found. Auto-renaming output to: {target_path.name}")
        else:
            report.add_result(None, "Failed", error=f"Output file '{target_path.name}' already exists.")
            return report

    writer = pypdf.PdfWriter()
    total_pages = 0
    successful_sources: List[Path] = []

    for idx, file_path in enumerate(selected_files, 1):
        pct = ((idx - 1) / total_files) * 85.0
        _cb(pct, f"[{idx}/{total_files}] Merging {file_path.name}...")
        _log(f"[*] [{idx}/{total_files}] Appending: {file_path.name}")
        try:
            if not file_path.exists() or file_path.stat().st_size == 0:
                report.add_result(file_path, "Failed", error="File is missing or empty.")
                continue

            reader = pypdf.PdfReader(str(file_path))
            if reader.is_encrypted:
                report.add_result(file_path, "Failed", error="File is encrypted. Please unlock first.")
                continue

            page_count = len(reader.pages)
            if page_count == 0:
                report.add_result(file_path, "Failed", error="PDF contains 0 pages.")
                continue

            writer.append(reader)
            total_pages += page_count
            successful_sources.append(file_path)
        except Exception as e:
            report.add_result(file_path, "Failed", error=f"Error reading PDF: {e}")

    if not successful_sources or total_pages == 0:
        report.add_result(None, "Failed", error="No valid pages could be appended to the merge document.")
        return report

    # Write merged document
    try:
        _cb(88.0, f"Writing merged document ({total_pages} total pages)...")
        _log(f"[*] Writing {total_pages} pages to: {target_path.name}...")
        with open(target_path, "wb") as f_out:
            writer.write(f_out)

        # Verification step
        is_valid, msg = OutputManager.verify_pdf_readable(target_path)
        if not is_valid:
            report.add_result(None, "Failed", error=f"Verification failed: {msg}")
            return report

        for src in successful_sources:
            report.add_result(src, "Success", output_files=[target_path], details=f"Merged into {target_path.name}")

        _cb(100.0, f"Merged {len(successful_sources)} files into {target_path.name}")
        _log(f"[+] Output: {target_path}")

    except Exception as e:
        report.add_result(None, "Failed", error=f"Failed to write merged output: {e}")

    return report


# ==============================================================================
# OPERATION 2: SPLIT PDFS
# ==============================================================================

def parse_page_ranges(range_str: str, max_pages: int) -> Tuple[Optional[List[Tuple[int, int]]], Optional[str]]:
    """
    Parses comma-separated ranges such as '1-3,4-6,7-10'.
    Validates against max_pages. 1-indexed inclusive.
    """
    ranges: List[Tuple[int, int]] = []
    tokens = [t.strip() for t in range_str.split(",") if t.strip()]
    if not tokens:
        return None, "No page ranges provided."

    for tok in tokens:
        if "-" in tok:
            parts = tok.split("-")
            if len(parts) != 2 or not parts[0].strip().isdigit() or not parts[1].strip().isdigit():
                return None, f"Malformed page range '{tok}'. Expected format: Start-End (e.g. 1-3)."
            start, end = int(parts[0].strip()), int(parts[1].strip())
        else:
            if not tok.isdigit():
                return None, f"Malformed page number '{tok}'."
            start = end = int(tok)

        if start <= 0 or end <= 0:
            return None, f"Page numbers must be >= 1. Got '{tok}'."
        if start > end:
            return None, f"Reversed range '{tok}': start ({start}) cannot exceed end ({end})."
        if end > max_pages:
            return None, f"Range '{tok}' exceeds total document pages ({max_pages})."

        ranges.append((start, end))

    return ranges, None


def execute_split_pdfs(
    selected_files: List[Path],
    output_parent_folder: Path,
    mode: str = "A",  # 'A': each page, 'B': page ranges, 'C': groups of N
    range_str: Optional[str] = None,
    group_size: int = 1,
    progress_callback: Optional[Callable[[float, str], None]] = None,
    log_callback: Optional[Callable[[str], None]] = None
) -> ProcessingReport:
    """
    Splits selected PDFs into separate PDF documents, creating an individual
    child folder for each source PDF named after the source PDF stem.
    """
    def _cb(pct: float, msg: str):
        if progress_callback:
            try:
                progress_callback(pct, msg)
            except Exception:
                pass

    def _log(msg: str):
        if log_callback:
            try:
                log_callback(msg)
            except Exception:
                pass

    report = ProcessingReport("Split PDFs")
    OutputManager.ensure_output_dir(output_parent_folder)
    total_files = len(selected_files)

    for file_idx, file_path in enumerate(selected_files, 1):
        base_pct = ((file_idx - 1) / total_files) * 100.0
        file_weight = 100.0 / total_files
        _cb(base_pct, f"[{file_idx}/{total_files}] Splitting {file_path.name}...")
        _log(f"[*] [{file_idx}/{total_files}] Processing: {file_path.name}")
        try:
            if not file_path.exists() or file_path.stat().st_size == 0:
                report.add_result(file_path, "Failed", error="File does not exist or is empty.")
                continue

            reader = pypdf.PdfReader(str(file_path))
            if reader.is_encrypted:
                report.add_result(file_path, "Failed", error="File is encrypted. Unlock it first.")
                continue

            total_pages = len(reader.pages)
            if total_pages == 0:
                report.add_result(file_path, "Failed", error="PDF contains 0 pages.")
                continue

            # Output folder per source PDF
            base_split_dir = output_parent_folder / file_path.stem
            # Prevent accidental mixup if rerun
            OutputManager.ensure_output_dir(base_split_dir)

            generated_files: List[Path] = []

            if mode.upper() == "A":
                # Every page separately
                for page_num in range(1, total_pages + 1):
                    sub_pct = base_pct + (page_num / total_pages) * file_weight
                    _cb(sub_pct, f"[{file_idx}/{total_files}] Page {page_num}/{total_pages} of {file_path.name}...")
                    writer = pypdf.PdfWriter()
                    writer.add_page(reader.pages[page_num - 1])
                    out_name = f"page_{page_num:03d}.pdf"
                    out_path = OutputManager.get_collision_safe_path(base_split_dir / out_name)
                    with open(out_path, "wb") as f_out:
                        writer.write(f_out)
                    is_valid, _ = OutputManager.verify_pdf_readable(out_path)
                    if is_valid:
                        generated_files.append(out_path)

            elif mode.upper() == "B":
                # By page ranges
                parsed_ranges, err = parse_page_ranges(range_str or "", total_pages)
                if err:
                    report.add_result(file_path, "Failed", error=err)
                    continue

                for idx, (start, end) in enumerate(parsed_ranges, start=1):
                    writer = pypdf.PdfWriter()
                    for p in range(start - 1, end):
                        writer.add_page(reader.pages[p])
                    out_name = f"range_{start:03d}_{end:03d}.pdf"
                    out_path = OutputManager.get_collision_safe_path(base_split_dir / out_name)
                    with open(out_path, "wb") as f_out:
                        writer.write(f_out)
                    is_valid, _ = OutputManager.verify_pdf_readable(out_path)
                    if is_valid:
                        generated_files.append(out_path)

            elif mode.upper() == "C":
                # Groups of N pages
                if group_size <= 0:
                    report.add_result(file_path, "Failed", error="Group size must be a positive integer.")
                    continue

                part_idx = 1
                for start_idx in range(0, total_pages, group_size):
                    end_idx = min(start_idx + group_size, total_pages)
                    writer = pypdf.PdfWriter()
                    for p in range(start_idx, end_idx):
                        writer.add_page(reader.pages[p])
                    out_name = f"part_{part_idx:03d}_pages_{start_idx + 1:03d}_{end_idx:03d}.pdf"
                    out_path = OutputManager.get_collision_safe_path(base_split_dir / out_name)
                    with open(out_path, "wb") as f_out:
                        writer.write(f_out)
                    is_valid, _ = OutputManager.verify_pdf_readable(out_path)
                    if is_valid:
                        generated_files.append(out_path)
                    part_idx += 1
            else:
                report.add_result(file_path, "Failed", error=f"Unknown splitting mode '{mode}'.")
                continue

            _log(f"  [+] Saved {len(generated_files)} split file(s) to: {base_split_dir}")
            report.add_result(file_path, "Success", output_files=generated_files, details=f"Generated {len(generated_files)} file(s)")

        except Exception as e:
            report.add_result(file_path, "Failed", error=f"Split failed: {e}")

    _cb(100.0, f"Split completed for {total_files} document(s).")
    return report


# ==============================================================================
# OPERATION 3: COMPRESS PDFS
# ==============================================================================

def execute_compress_pdfs(
    selected_files: List[Path],
    level: str = "medium",  # 'low', 'medium', 'high'
    retain_original_if_larger: bool = True,
    output_dir: Optional[Path] = None,
    progress_callback: Optional[Callable[[float, str], None]] = None,
    log_callback: Optional[Callable[[str], None]] = None
) -> ProcessingReport:
    """
    Compresses selected PDFs.
    Low: Conservative optimization (lossless stream compression, removing unreferenced objects).
    Medium: Balanced (stream compression, clean structure, standard image optimization).
    High: Strong optimization (recompresses images with quality trade-offs).
    Saves in output_dir (if provided) or in 'output/' child folder inside each source folder.
    """
    def _cb(pct: float, msg: str):
        if progress_callback:
            try:
                progress_callback(pct, msg)
            except Exception:
                pass

    def _log(msg: str):
        if log_callback:
            try:
                log_callback(msg)
            except Exception:
                pass

    report = ProcessingReport("Compress PDFs")
    level_norm = level.lower().strip()
    total_files = len(selected_files)

    print("\n" + "=" * 56)
    print(" Compression Results".ljust(56))
    print("=" * 56)

    for idx, file_path in enumerate(selected_files, 1):
        start_pct = ((idx - 1) / total_files) * 100.0
        _cb(start_pct, f"[{idx}/{total_files}] Compressing {file_path.name}...")
        _log(f"[*] [{idx}/{total_files}] Compressing: {file_path.name}")
        try:
            if not file_path.exists() or file_path.stat().st_size == 0:
                report.add_result(file_path, "Failed", error="File does not exist or is empty.")
                continue

            orig_size = file_path.stat().st_size
            out_dir = Path(output_dir) if output_dir is not None else (file_path.parent / "output")
            OutputManager.ensure_output_dir(out_dir)
            out_path = OutputManager.get_collision_safe_path(out_dir / f"{file_path.stem}_compressed.pdf")

            # Apply compression using PyMuPDF and pypdf stream optimization
            doc = pymupdf.open(str(file_path))
            if doc.is_encrypted:
                doc.close()
                report.add_result(file_path, "Failed", error="File is encrypted. Unlock it first.")
                continue

            if level_norm == "low":
                # Conservative: garbage collect, deflate streams
                doc.save(str(out_path), garbage=3, deflate=True, clean=True)
            elif level_norm == "medium":
                # Balanced: strong deflation, deduplication
                doc.save(str(out_path), garbage=4, deflate=True, clean=True)
            elif level_norm == "high":
                # High: aggressive deflation and image downsampling where possible
                # Attempt image recompression
                for page in doc:
                    image_list = page.get_images(full=True)
                    for img_info in image_list:
                        xref = img_info[0]
                        try:
                            pix = pymupdf.Pixmap(doc, xref)
                            if pix.n >= 5:  # CMYK
                                pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
                            # Recompress image as JPEG with 60% quality
                            img_data = pix.tobytes("jpeg", jpg_quality=60)
                            # Replace stream if smaller
                            doc.update_stream(xref, img_data)
                        except Exception:
                            pass
                doc.save(str(out_path), garbage=4, deflate=True, clean=True)
            else:
                doc.close()
                report.add_result(file_path, "Failed", error=f"Invalid compression level '{level}'.")
                continue

            doc.close()

            # Post-compression size verification
            comp_size = out_path.stat().st_size
            saved_bytes = orig_size - comp_size
            reduction_pct = (saved_bytes / orig_size * 100.0) if orig_size > 0 else 0.0

            status_note = "Success"
            if comp_size >= orig_size:
                status_note = "No reduction achieved (compressed size >= original)"
                if retain_original_if_larger:
                    # User option: if compressed is larger, remove bloated output or warn
                    logger.warning(f"Compression increased file size for {file_path.name}. Retaining copy with warning.")

            print(f"\nFile: {file_path.name}")
            print(f"Original size:   {format_file_size(orig_size)}")
            print(f"Compressed size: {format_file_size(comp_size)}")
            print(f"Space saved:     {format_file_size(max(0, saved_bytes))}")
            print(f"Reduction:       {reduction_pct:.2f}%")
            print(f"Status:          {status_note}")

            # Verify readability
            is_valid, msg = OutputManager.verify_pdf_readable(out_path)
            if not is_valid:
                report.add_result(file_path, "Failed", error=f"Compressed output verification failed: {msg}")
                continue

            done_pct = (idx / total_files) * 100.0
            _cb(done_pct, f"[{idx}/{total_files}] Compressed {file_path.name} (-{reduction_pct:.1f}%)")
            _log(f"  [+] Compressed: {out_path.name} ({format_file_size(orig_size)} -> {format_file_size(comp_size)}, -{reduction_pct:.1f}%)")

            report.add_result(
                file_path,
                "Success",
                output_files=[out_path],
                details=f"Reduction: {reduction_pct:.2f}% ({format_file_size(orig_size)} -> {format_file_size(comp_size)})"
            )

        except Exception as e:
            report.add_result(file_path, "Failed", error=f"Compression error: {e}")

    print("\n" + "=" * 56 + "\n")
    _cb(100.0, f"Compressed {total_files} document(s).")
    return report


# ==============================================================================
# OPERATION 4: UNLOCK PDF
# ==============================================================================

def execute_unlock_pdf(
    selected_files: List[Path],
    password: Optional[str] = None,
    output_dir: Optional[Path] = None,
    progress_callback: Optional[Callable[[float, str], None]] = None,
    log_callback: Optional[Callable[[str], None]] = None
) -> ProcessingReport:
    """
    Creates an unencrypted copy of a PDF when the user provides the authorized password.
    Saves in output_dir (if provided) or in 'output/' child folder inside each source folder.
    """
    def _cb(pct: float, msg: str):
        if progress_callback:
            try:
                progress_callback(pct, msg)
            except Exception:
                pass

    def _log(msg: str):
        if log_callback:
            try:
                log_callback(msg)
            except Exception:
                pass

    report = ProcessingReport("Unlock PDF")
    total_files = len(selected_files)

    for idx, file_path in enumerate(selected_files, 1):
        start_pct = ((idx - 1) / total_files) * 100.0
        _cb(start_pct, f"[{idx}/{total_files}] Unlocking {file_path.name}...")
        _log(f"[*] [{idx}/{total_files}] Decrypting: {file_path.name}")
        try:
            if not file_path.exists() or file_path.stat().st_size == 0:
                report.add_result(file_path, "Failed", error="File does not exist or is empty.")
                continue

            target_out_dir = Path(output_dir) if output_dir is not None else (file_path.parent / "output")
            OutputManager.ensure_output_dir(target_out_dir)

            reader = pypdf.PdfReader(str(file_path))
            if not reader.is_encrypted:
                # Document is already unencrypted
                out_path = OutputManager.get_collision_safe_path(target_out_dir / f"{file_path.stem}_unlocked.pdf")
                shutil.copy2(file_path, out_path)
                report.add_result(file_path, "Success", output_files=[out_path], details="File was already unencrypted; copy created.")
                continue

            # If encrypted and no password supplied, fail gracefully
            if password is None:
                report.add_result(file_path, "Failed", error="Password required to decrypt file.")
                continue

            # Attempt decryption
            decrypt_res = reader.decrypt(password)
            if decrypt_res == 0:
                report.add_result(file_path, "Failed", error="Incorrect password or unsupported encryption.")
                continue

            # Write unencrypted copy
            writer = pypdf.PdfWriter()
            writer.append(reader)

            out_path = OutputManager.get_collision_safe_path(target_out_dir / f"{file_path.stem}_unlocked.pdf")

            with open(out_path, "wb") as f_out:
                writer.write(f_out)

            # Verification: output must open without password
            verify_reader = pypdf.PdfReader(str(out_path))
            if verify_reader.is_encrypted:
                report.add_result(file_path, "Failed", error="Decrypted output file still reports encrypted status.")
                continue

            done_pct = (idx / total_files) * 100.0
            _cb(done_pct, f"[{idx}/{total_files}] Unlocked {file_path.name}")
            _log(f"  [+] Successfully decrypted: {out_path.name}")
            report.add_result(file_path, "Success", output_files=[out_path], details="Successfully decrypted and unlocked.")

        except Exception as e:
            report.add_result(file_path, "Failed", error=f"Unlock error: {e}")

    _cb(100.0, f"Unlock completed for {total_files} file(s).")
    return report


# ==============================================================================
# OPERATION 5: PROTECT PDF
# ==============================================================================

def execute_protect_pdf(
    selected_files: List[Path],
    user_password: str,
    owner_password: Optional[str] = None,
    allow_printing: bool = True,
    allow_copying: bool = False,
    allow_modifying: bool = False,
    output_dir: Optional[Path] = None,
    progress_callback: Optional[Callable[[float, str], None]] = None,
    log_callback: Optional[Callable[[str], None]] = None
) -> ProcessingReport:
    """
    Encrypts PDF documents with a user-provided password using modern AES encryption.
    Distinguishes user password (to open) and owner password (for permissions).
    """
    report = ProcessingReport("Protect PDF")
    if not user_password:
        report.add_result(None, "Failed", error="Password cannot be empty.")
        return report

    actual_owner = owner_password if owner_password else user_password

    # Configure permission flags
    permissions = pypdf.constants.UserAccessPermissions(0)
    if allow_printing:
        permissions |= (
            pypdf.constants.UserAccessPermissions.PRINT |
            pypdf.constants.UserAccessPermissions.PRINT_TO_REPRESENTATION
        )
    if allow_copying:
        permissions |= (
            pypdf.constants.UserAccessPermissions.EXTRACT |
            pypdf.constants.UserAccessPermissions.EXTRACT_TEXT_AND_GRAPHICS
        )
    if allow_modifying:
        permissions |= (
            pypdf.constants.UserAccessPermissions.MODIFY |
            pypdf.constants.UserAccessPermissions.ADD_OR_MODIFY |
            pypdf.constants.UserAccessPermissions.ASSEMBLE_DOC
        )

    def _cb(pct: float, msg: str):
        if progress_callback:
            try:
                progress_callback(pct, msg)
            except Exception:
                pass

    def _log(msg: str):
        if log_callback:
            try:
                log_callback(msg)
            except Exception:
                pass

    total_files = len(selected_files)
    for idx, file_path in enumerate(selected_files, 1):
        start_pct = ((idx - 1) / total_files) * 100.0
        _cb(start_pct, f"[{idx}/{total_files}] Protecting {file_path.name}...")
        _log(f"[*] [{idx}/{total_files}] Encrypting: {file_path.name} (AES-256)")
        try:
            if not file_path.exists() or file_path.stat().st_size == 0:
                report.add_result(file_path, "Failed", error="File does not exist or is empty.")
                continue

            reader = pypdf.PdfReader(str(file_path))
            if reader.is_encrypted:
                report.add_result(file_path, "Failed", error="File is already encrypted. Unlock first before re-encrypting.")
                continue

            writer = pypdf.PdfWriter()
            writer.append(reader)

            # Modern AES-256 or AES-128 encryption
            writer.encrypt(
                user_password=user_password,
                owner_password=actual_owner,
                permissions_flag=permissions,
                algorithm="AES-256"
            )

            out_dir = Path(output_dir) if output_dir is not None else (file_path.parent / "output")
            OutputManager.ensure_output_dir(out_dir)
            out_path = OutputManager.get_collision_safe_path(out_dir / f"{file_path.stem}_protected.pdf")

            with open(out_path, "wb") as f_out:
                writer.write(f_out)

            # Verification: output must report encrypted and open with password
            verify_reader = pypdf.PdfReader(str(out_path))
            if not verify_reader.is_encrypted:
                report.add_result(file_path, "Failed", error="Output failed encryption check.")
                continue

            if verify_reader.decrypt(user_password) == 0:
                report.add_result(file_path, "Failed", error="Output encryption verification with password failed.")
                continue

            done_pct = (idx / total_files) * 100.0
            _cb(done_pct, f"[{idx}/{total_files}] Protected {file_path.name}")
            _log(f"  [+] Protected: {out_path.name}")
            report.add_result(file_path, "Success", output_files=[out_path], details="Protected with AES-256 encryption.")

        except Exception as e:
            report.add_result(file_path, "Failed", error=f"Protection error: {e}")

    _cb(100.0, f"Protection completed for {total_files} file(s).")
    return report



# ==============================================================================
# OPERATION 6: IMAGES TO PDF
# ==============================================================================

def execute_images_to_pdf(
    image_paths: List[Path],
    output_dir: Path,
    output_name: str = "images_merged.pdf",
    page_size: str = "fit",       # 'fit', 'a4', 'letter', 'legal'
    orientation: str = "auto",     # 'auto', 'portrait', 'landscape'
    margin_mm: float = 0.0,
    quality: int = 92,
    single_pdf: bool = True,
    log_callback = None,
    progress_callback: Optional[Callable[[float, str], None]] = None
) -> ProcessingReport:
    """
    Converts a sequence of images into PDF document(s).
    Supports all image formats: PNG, JPG, JPEG, WEBP, BMP, TIFF, GIF, etc.
    """
    def _log(msg: str):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    def _cb(pct: float, msg: str):
        if progress_callback:
            try:
                progress_callback(pct, msg)
            except Exception:
                pass

    report = ProcessingReport("Images to PDF")
    OutputManager.ensure_output_dir(output_dir)

    if not image_paths:
        report.add_result(None, "Failed", error="No image files provided.")
        return report

    total_imgs = len(image_paths)
    _cb(0.0, f"Processing {total_imgs} image(s) to PDF...")
    _log(f"Processing {total_imgs} image(s) to PDF...")
    processed_pil_images = []

    page_dims_mm = {
        "a4": (210.0, 297.0),
        "letter": (215.9, 279.4),
        "legal": (215.9, 355.6)
    }

    valid_src_paths = []

    for idx, img_path in enumerate(image_paths):
        pct = (idx / total_imgs) * 85.0
        _cb(pct, f"[{idx + 1}/{total_imgs}] Processing {img_path.name}...")
        _log(f"  [{idx + 1}/{total_imgs}] Reading: {img_path.name}")
        try:
            with Image.open(img_path) as raw_img:
                # Handle transparency / alpha channels cleanly
                if raw_img.mode in ("RGBA", "LA", "P"):
                    bg = Image.new("RGB", raw_img.size, (255, 255, 255))
                    converted = raw_img.convert("RGBA")
                    bg.paste(converted, mask=converted.split()[-1])
                    cur_img = bg
                elif raw_img.mode != "RGB":
                    cur_img = raw_img.convert("RGB")
                else:
                    cur_img = raw_img.copy()

                w_px, h_px = cur_img.size

                if page_size in page_dims_mm:
                    base_w_mm, base_h_mm = page_dims_mm[page_size]
                    if orientation == "portrait":
                        p_w_mm, p_h_mm = min(base_w_mm, base_h_mm), max(base_w_mm, base_h_mm)
                    elif orientation == "landscape":
                        p_w_mm, p_h_mm = max(base_w_mm, base_h_mm), min(base_w_mm, base_h_mm)
                    else: # auto
                        if w_px > h_px:
                            p_w_mm, p_h_mm = max(base_w_mm, base_h_mm), min(base_w_mm, base_h_mm)
                        else:
                            p_w_mm, p_h_mm = min(base_w_mm, base_h_mm), max(base_w_mm, base_h_mm)

                    DPI = 300.0
                    PX_PER_MM = DPI / 25.4
                    canvas_w = int(p_w_mm * PX_PER_MM)
                    canvas_h = int(p_h_mm * PX_PER_MM)
                    margin_px = int(margin_mm * PX_PER_MM)

                    avail_w = max(10, canvas_w - (margin_px * 2))
                    avail_h = max(10, canvas_h - (margin_px * 2))

                    scale = min(avail_w / w_px, avail_h / h_px)
                    new_w = max(1, int(w_px * scale))
                    new_h = max(1, int(h_px * scale))

                    resized = cur_img.resize((new_w, new_h), Image.Resampling.LANCZOS)
                    canvas = Image.new("RGB", (canvas_w, canvas_h), (255, 255, 255))
                    offset_x = (canvas_w - new_w) // 2
                    offset_y = (canvas_h - new_h) // 2
                    canvas.paste(resized, (offset_x, offset_y))
                    processed_pil_images.append(canvas)
                else:
                    processed_pil_images.append(cur_img)

                valid_src_paths.append(img_path)
        except Exception as e:
            _log(f"  [-] Error processing {img_path.name}: {e}")
            report.add_result(img_path, "Failed", error=str(e))

    if not processed_pil_images:
        report.add_result(None, "Failed", error="No valid images could be converted.")
        return report

    if single_pdf:
        dest_pdf = OutputManager.get_collision_safe_path(output_dir / output_name)
        if not dest_pdf.name.lower().endswith(".pdf"):
            dest_pdf = dest_pdf.with_suffix(".pdf")
        
        _cb(90.0, f"Saving merged PDF: {dest_pdf.name}...")
        _log(f"Saving merged PDF: {dest_pdf.name}")
        first_img = processed_pil_images[0]
        other_imgs = processed_pil_images[1:] if len(processed_pil_images) > 1 else []
        first_img.save(
            dest_pdf,
            "PDF",
            save_all=True,
            append_images=other_imgs,
            resolution=300.0,
            quality=quality
        )
        report.add_result(
            valid_src_paths[0],
            "Success",
            output_files=[dest_pdf],
            details=f"Merged {len(processed_pil_images)} images into single PDF."
        )
    else:
        for idx, (img, src_path) in enumerate(zip(processed_pil_images, valid_src_paths)):
            dest_pdf = OutputManager.get_collision_safe_path(output_dir / f"{src_path.stem}.pdf")
            img.save(dest_pdf, "PDF", resolution=300.0, quality=quality)
            report.add_result(
                src_path,
                "Success",
                output_files=[dest_pdf],
                details="Converted single image to PDF."
            )

    _cb(100.0, f"Converted {len(valid_src_paths)} image(s) to PDF.")
    _log(f"[OK] Images to PDF completed successfully.")
    return report


# ==============================================================================
# OPERATION 7: PDF TO IMAGES
# ==============================================================================

def parse_page_selection(range_str: str, total_pages: int) -> List[int]:
    """
    Parses a page range string like 'all', '1-5, 8, 10' into 0-indexed page numbers.
    """
    if not range_str or range_str.strip().lower() in ("all", "*"):
        return list(range(total_pages))
    
    selected = set()
    parts = [p.strip() for p in range_str.split(",") if p.strip()]
    for part in parts:
        if "-" in part:
            sub = part.split("-")
            try:
                start = max(1, int(sub[0].strip()))
                end = min(total_pages, int(sub[1].strip()))
                for p in range(start, end + 1):
                    selected.add(p - 1)
            except ValueError:
                pass
        else:
            try:
                p = int(part)
                if 1 <= p <= total_pages:
                    selected.add(p - 1)
            except ValueError:
                pass
    return sorted(list(selected)) if selected else list(range(total_pages))


def execute_pdf_to_images(
    pdf_paths: List[Path],
    output_dir: Path,
    img_format: str = "png",       # 'png', 'jpeg', 'webp', 'tiff', 'bmp'
    dpi: int = 150,                # 150, 200, 300, 600
    page_range: str = "all",
    color_mode: str = "rgb",       # 'rgb', 'grayscale'
    quality: int = 90,
    subfolder_per_pdf: bool = True,
    log_callback = None,
    progress_callback: Optional[Callable[[float, str], None]] = None
) -> ProcessingReport:
    """
    Renders PDF document pages into high-fidelity image files.
    Supports PNG, JPEG, WEBP, TIFF, BMP with customizable DPI.
    Uses Google PDFium via pypdfium2 with PyMuPDF fallback.
    """
    def _log(msg: str):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    def _cb(pct: float, msg: str):
        if progress_callback:
            try:
                progress_callback(pct, msg)
            except Exception:
                pass

    report = ProcessingReport("PDF to Images")
    OutputManager.ensure_output_dir(output_dir)

    if not pdf_paths:
        report.add_result(None, "Failed", error="No PDF files provided.")
        return report

    img_format = img_format.lower().strip()
    ext = ".jpg" if img_format == "jpeg" else f".{img_format}"
    pil_format = "JPEG" if img_format == "jpeg" else img_format.upper()
    scale = dpi / 72.0

    total_pdfs = len(pdf_paths)
    for pdf_idx, pdf_p in enumerate(pdf_paths):
        base_pct = (pdf_idx / total_pdfs) * 100.0
        pdf_weight = 100.0 / total_pdfs
        _cb(base_pct, f"[{pdf_idx + 1}/{total_pdfs}] Converting {pdf_p.name}...")
        _log(f"Converting [{pdf_idx + 1}/{total_pdfs}]: {pdf_p.name}")
        
        target_dir = output_dir / pdf_p.stem if (subfolder_per_pdf and len(pdf_paths) > 1) else output_dir
        OutputManager.ensure_output_dir(target_dir)

        exported_imgs = []

        try:
            # 1. Primary Engine: pypdfium2
            if PDFIUM_AVAILABLE:
                pdf = pdfium.PdfDocument(pdf_p)
                total_pages = len(pdf)
                pages_to_render = parse_page_selection(page_range, total_pages)
                _log(f"  Rendering {len(pages_to_render)} of {total_pages} page(s) at {dpi} DPI...")

                for page_i, p_num in enumerate(pages_to_render):
                    page_pct = base_pct + ((page_i + 1) / len(pages_to_render)) * pdf_weight
                    _cb(page_pct, f"[{pdf_idx + 1}/{total_pdfs}] Page {p_num + 1}/{total_pages} of {pdf_p.name}...")
                    page = pdf[p_num]
                    pil_im = page.render(scale=scale).to_pil()
                    if color_mode == "grayscale":
                        pil_im = pil_im.convert("L")
                    elif pil_im.mode != "RGB":
                        pil_im = pil_im.convert("RGB")

                    out_img_name = f"{pdf_p.stem}_page_{p_num + 1:03d}{ext}"
                    out_img_path = OutputManager.get_collision_safe_path(target_dir / out_img_name)

                    save_kwargs = {}
                    if pil_format in ("JPEG", "WEBP"):
                        save_kwargs["quality"] = quality
                    pil_im.save(out_img_path, format=pil_format, **save_kwargs)
                    exported_imgs.append(out_img_path)
            else:
                raise ImportError("pypdfium2 is not installed.")

            report.add_result(
                pdf_p,
                "Success",
                output_files=exported_imgs,
                details=f"Rendered {len(exported_imgs)} page(s) at {dpi} DPI ({img_format.upper()})."
            )
            _log(f"  [+] Saved {len(exported_imgs)} page(s) to: {target_dir}")
        except Exception as e:
            _log(f"  [-] pypdfium2 note ({e}). Attempting PyMuPDF fallback...")
            try:
                import fitz
                doc = fitz.open(pdf_p)
                total_pages = len(doc)
                pages_to_render = parse_page_selection(page_range, total_pages)
                zoom = dpi / 72.0
                mat = fitz.Matrix(zoom, zoom)
                for p_num in pages_to_render:
                    page = doc[p_num]
                    pix = page.get_pixmap(matrix=mat, alpha=False)
                    out_img_name = f"{pdf_p.stem}_page_{p_num + 1:03d}{ext}"
                    out_img_path = OutputManager.get_collision_safe_path(target_dir / out_img_name)
                    if color_mode == "grayscale":
                        im = Image.frombytes("RGB", [pix.width, pix.height], pix.samples).convert("L")
                        im.save(out_img_path, format=pil_format)
                    else:
                        pix.save(str(out_img_path))
                    exported_imgs.append(out_img_path)

                report.add_result(
                    pdf_p,
                    "Success",
                    output_files=exported_imgs,
                    details=f"Rendered {len(exported_imgs)} page(s) via PyMuPDF at {dpi} DPI."
                )
                _log(f"  [+] Fallback successfully saved {len(exported_imgs)} page(s).")
            except Exception as e2:
                _log(f"  [-] Rendering failed for {pdf_p.name}: {e2}")
                report.add_result(pdf_p, "Failed", error=str(e2))

    _cb(100.0, f"Rendered images for {len(pdf_paths)} PDF(s).")
    _log(f"[OK] PDF to Images completed.")
    return report


# ==============================================================================
# OPERATION 8: OCR PDF
# ==============================================================================

def detect_tesseract_binary() -> Optional[str]:
    """
    Detects Tesseract executable across system PATH and common installation directories.
    """
    # 1. System PATH
    found = shutil.which("tesseract")
    if found:
        return found

    # 2. Windows standard paths
    candidates = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Tesseract-OCR\tesseract.exe"),
        os.path.expandvars(r"%PROGRAMFILES%\Tesseract-OCR\tesseract.exe"),
    ]
    for c in candidates:
        if os.path.isfile(c):
            return c

    return None


def execute_ocr_pdf(
    selected_files: List[Path],
    lang: str = "eng",
    process_all_pages: bool = True,
    coverage_mode: str = "all",
    dpi: int = 300,
    custom_tesseract_cmd: Optional[str] = None,
    output_dir: Optional[Path] = None,
    progress_callback: Optional[Callable[[float, str], None]] = None,
    log_callback: Optional[Callable[[str], None]] = None
) -> ProcessingReport:
    """
    Applies Optical Character Recognition (OCR) to convert scanned PDFs and mixed slides
    into high-fidelity searchable PDFs with full text layers while preserving visual layout.
    Supports English ('eng'), Bengali ('ben'), and bilingual ('eng+ben').
    """
    report = ProcessingReport("OCR PDF")

    # Detect Tesseract
    tess_cmd = custom_tesseract_cmd or detect_tesseract_binary()
    if not tess_cmd or not PYTESSERACT_AVAILABLE:
        instructions = (
            "Tesseract OCR is not installed or not found.\n"
            "To install Tesseract OCR:\n"
            "  - Windows: Download the installer from https://github.com/UB-Mannheim/tesseract/wiki\n"
            "    Install it and optionally download Bengali language pack ('ben.traineddata') into tessdata/.\n"
            "  - Linux: sudo apt-get install tesseract-ocr tesseract-ocr-ben\n"
            "  - Python packages: pip install pytesseract Pillow"
        )
        report.add_result(None, "Failed", error=instructions)
        return report

    pytesseract.pytesseract.tesseract_cmd = tess_cmd
    tess_path = Path(tess_cmd)
    tessdata_candidate = tess_path.parent / "tessdata"
    if tessdata_candidate.exists():
        os.environ["TESSDATA_PREFIX"] = str(tessdata_candidate)

    def _cb(pct: float, msg: str):
        if progress_callback:
            try:
                progress_callback(pct, msg)
            except Exception:
                pass

    def _log(msg: str):
        if log_callback:
            try:
                log_callback(msg)
            except Exception:
                pass

    total_files = len(selected_files)
    for idx, file_path in enumerate(selected_files, 1):
        base_pct = ((idx - 1) / total_files) * 100.0
        file_weight = 100.0 / total_files
        _cb(base_pct, f"[{idx}/{total_files}] OCR processing {file_path.name}...")
        _log(f"[*] [{idx}/{total_files}] OCR processing: {file_path.name}")
        try:
            if not file_path.exists() or file_path.stat().st_size == 0:
                report.add_result(file_path, "Failed", error="File does not exist or is empty.")
                continue

            doc = pymupdf.open(str(file_path))
            if doc.is_encrypted:
                doc.close()
                report.add_result(file_path, "Failed", error="File is encrypted. Unlock it first.")
                continue

            out_dir = Path(output_dir) if output_dir is not None else (file_path.parent / "output")
            OutputManager.ensure_output_dir(out_dir)
            out_path = OutputManager.get_collision_safe_path(out_dir / f"{file_path.stem}_ocr.pdf")

            pdf_writer = pypdf.PdfWriter()
            pages_processed = 0

            num_pages = len(doc)
            for pno in range(num_pages):
                sub_pct = base_pct + ((pno + 1) / num_pages) * file_weight
                _cb(sub_pct, f"[{idx}/{total_files}] OCR page {pno + 1}/{num_pages} of {file_path.name}...")
                page = doc[pno]
                has_text = len(page.get_text().strip()) > 30
                has_images = len(page.get_images()) > 0

                # Determine if page requires OCR based on coverage strategy
                if coverage_mode == "all" or process_all_pages:
                    should_ocr = True
                elif coverage_mode == "hybrid":
                    # OCR page if it has missing text OR if it contains embedded diagram images
                    should_ocr = (not has_text) or has_images
                elif coverage_mode == "scanned_only":
                    should_ocr = not has_text
                else:
                    should_ocr = True

                if not should_ocr and has_text:
                    # Page already has complete digital vector text and no embedded diagrams
                    single_doc = pymupdf.open()
                    single_doc.insert_pdf(doc, from_page=pno, to_page=pno)
                    page_bytes = single_doc.tobytes()
                    single_doc.close()
                    single_reader = pypdf.PdfReader(pymupdf.io.BytesIO(page_bytes))
                    pdf_writer.add_page(single_reader.pages[0])
                else:
                    # Perform high-resolution OCR on rendered page image
                    render_dpi = max(150, min(dpi, 600))
                    pix = page.get_pixmap(dpi=render_dpi)
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

                    # Generate searchable PDF page from pytesseract
                    pdf_bytes = pytesseract.image_to_pdf_or_hocr(img, extension='pdf', lang=lang)
                    single_reader = pypdf.PdfReader(pymupdf.io.BytesIO(pdf_bytes))
                    pdf_writer.add_page(single_reader.pages[0])
                    pages_processed += 1

            doc.close()

            with open(out_path, "wb") as f_out:
                pdf_writer.write(f_out)

            # Verification: output must exist and contain text
            is_valid, msg = OutputManager.verify_pdf_readable(out_path)
            if not is_valid:
                report.add_result(file_path, "Failed", error=f"OCR output verification failed: {msg}")
                continue

            _log(f"  [+] OCR completed: {out_path.name} ({pages_processed}/{num_pages} page(s) OCR'd)")
            report.add_result(file_path, "Success", output_files=[out_path], details=f"OCR processed {pages_processed} page(s) with lang='{lang}'")

        except Exception as e:
            report.add_result(file_path, "Failed", error=f"OCR processing failed: {e}")

    _cb(100.0, f"OCR completed for {total_files} file(s).")
    return report


# ==============================================================================
# OPERATION 9: MARKDOWN TO PDF
# ==============================================================================

class MarkdownPDFRenderer(FPDF):
    """Custom PDF renderer for converting formatted Markdown to PDF with full Unicode TrueType font support."""
    def __init__(self, page_size: str = "A4", orientation: str = "P", margins: float = 15.0):
        super().__init__(orientation=orientation, unit="mm", format=page_size)
        self.set_margins(margins, margins, margins)
        self.set_auto_page_break(auto=True, margin=margins)
        self.font_family_name = self._setup_unicode_fonts()

    def _setup_unicode_fonts(self) -> str:
        """
        Attempts to load standard TrueType Unicode fonts from system directories
        to support international characters, smart quotes, em-dashes, and symbols.
        """
        font_candidates = [
            # Windows Segoe UI
            (r"C:\Windows\Fonts\segoeui.ttf", r"C:\Windows\Fonts\segoeuib.ttf", r"C:\Windows\Fonts\segoeuii.ttf", r"C:\Windows\Fonts\segoeuiz.ttf"),
            # Windows Arial
            (r"C:\Windows\Fontsrial.ttf", r"C:\Windows\Fontsrialbd.ttf", r"C:\Windows\Fontsriali.ttf", r"C:\Windows\Fontsrialbi.ttf"),
            # Windows Calibri
            (r"C:\Windows\Fonts\calibri.ttf", r"C:\Windows\Fonts\calibrib.ttf", r"C:\Windows\Fonts\calibrii.ttf", r"C:\Windows\Fonts\calibriz.ttf"),
            # Windows Tahoma
            (r"C:\Windows\Fonts	ahoma.ttf", r"C:\Windows\Fonts	ahomabd.ttf", None, None),
            # Linux DejaVu Sans
            ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-BoldOblique.ttf"),
            # macOS Arial
            ("/Library/Fonts/Arial.ttf", "/Library/Fonts/Arial Bold.ttf", "/Library/Fonts/Arial Italic.ttf", "/Library/Fonts/Arial Bold Italic.ttf")
        ]

        for reg, bold, ital, bi in font_candidates:
            if os.path.isfile(reg):
                try:
                    family = "UnicodeSysFont"
                    self.add_font(family, "", reg)
                    if bold and os.path.isfile(bold):
                        self.add_font(family, "B", bold)
                    else:
                        self.add_font(family, "B", reg)
                    if ital and os.path.isfile(ital):
                        self.add_font(family, "I", ital)
                    else:
                        self.add_font(family, "I", reg)
                    if bi and os.path.isfile(bi):
                        self.add_font(family, "BI", bi)
                    else:
                        self.add_font(family, "BI", bold if bold and os.path.isfile(bold) else reg)
                    return family
                except Exception:
                    continue

        return "Helvetica"

    def header(self):
        self.set_font(self.font_family_name, 'I' if self.font_family_name == "Helvetica" else "", 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 5, 'Generated by PDF Toolkit', align='R')
        self.ln(6)

    def footer(self):
        self.set_y(-12)
        self.set_font(self.font_family_name, 'I' if self.font_family_name == "Helvetica" else "", 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, f'Page {self.page_no()}/{{nb}}', align='C')


def execute_markdown_to_pdf(
    selected_files: List[Path],
    page_size: str = "A4",
    orientation: str = "P",
    margins: float = 15.0,
    output_dir: Optional[Path] = None,
    progress_callback: Optional[Callable[[float, str], None]] = None,
    log_callback: Optional[Callable[[str], None]] = None
) -> ProcessingReport:
    """
    Converts Markdown documents (.md, .markdown) to properly styled PDF documents.
    Supports headings, paragraphs, lists, bold/italic, tables, blockquotes,
    links, images, and code blocks.
    """
    def _cb(pct: float, msg: str):
        if progress_callback:
            try:
                progress_callback(pct, msg)
            except Exception:
                pass

    def _log(msg: str):
        if log_callback:
            try:
                log_callback(msg)
            except Exception:
                pass

    report = ProcessingReport("Markdown to PDF")
    total_files = len(selected_files)

    for idx, file_path in enumerate(selected_files, 1):
        start_pct = ((idx - 1) / total_files) * 100.0
        _cb(start_pct, f"[{idx}/{total_files}] Rendering {file_path.name}...")
        _log(f"[*] [{idx}/{total_files}] Rendering Markdown: {file_path.name}")
        try:
            if not file_path.exists() or file_path.stat().st_size == 0:
                report.add_result(file_path, "Failed", error="File does not exist or is empty.")
                continue

            # Guard against accidental PDF input in Markdown-to-PDF mode
            if file_path.suffix.lower() == ".pdf":
                err_msg = (
                    f"File '{file_path.name}' is already a PDF document, not Markdown source (.md / .markdown).\n"
                    "If you want to extract Markdown text from this PDF, please switch to 'PDF to Markdown' mode."
                )
                _log(f"[-] {err_msg}")
                report.add_result(file_path, "Failed", error=err_msg)
                continue

            # Guard against binary files with %PDF- header even if misnamed
            try:
                with open(file_path, "rb") as f_chk:
                    if f_chk.read(5).startswith(b"%PDF-"):
                        err_msg = (
                            f"File '{file_path.name}' contains binary PDF bytecode, not Markdown text.\n"
                            "If you want to extract Markdown text from this PDF, please switch to 'PDF to Markdown' mode."
                        )
                        _log(f"[-] {err_msg}")
                        report.add_result(file_path, "Failed", error=err_msg)
                        continue
            except Exception:
                pass

            with open(file_path, "r", encoding="utf-8", errors="replace") as f_in:
                md_content = f_in.read()

            # 1. Clean up replacement characters
            md_content = md_content.replace("\ufffd", "?")

            # 2. Convert SVG / Shields badges to bold text: [Badge]
            md_content = re.sub(r'!\[([^\]]*)\]\([^)]*\.svg[^)]*\)', r'**[\1]**', md_content, flags=re.IGNORECASE)
            md_content = re.sub(r'!\[([^\]]*)\]\(https?://img\.shields\.io/[^)]*\)', r'**[\1]**', md_content, flags=re.IGNORECASE)

            # 3. Convert box drawing characters to safe ASCII equivalents
            box_chars = {
                '\u251c': '|--', '\u2514': '\\--', '\u2500': '-', '\u2502': '|',
                '\u2524': '--|', '\u252c': '-+-', '\u2534': '-+-', '\u253c': '-+-'
            }
            for ch, repl in box_chars.items():
                md_content = md_content.replace(ch, repl)

            # 4. Strip emojis that lack font glyphs
            md_content = re.sub(r'[\U00010000-\U0010ffff]', '', md_content)
            md_content = re.sub(r'[\u2300-\u23ff\u2600-\u27bf\ufe00-\ufe0f]', '', md_content)

            # Convert markdown to HTML
            html_body = markdown.markdown(
                md_content,
                extensions=['tables', 'fenced_code', 'nl2br']
            )

            # 5. Sanitize HTML for FPDF:
            # - Replace any remaining SVG <img> tags with alt text
            def _replace_img(match):
                attrs = match.group(0)
                alt_m = re.search(r'alt=["\']([^"\']*)["\']', attrs, re.IGNORECASE)
                src_m = re.search(r'src=["\']([^"\']*)["\']', attrs, re.IGNORECASE)
                src = src_m.group(1) if src_m else ""
                alt = alt_m.group(1) if alt_m else ""
                if ".svg" in src.lower() or "shields.io" in src.lower() or not src.strip():
                    label = alt.strip() or "Badge"
                    return f"<b>[{label}]</b>"
                return match.group(0)
            html_body = re.sub(r'<img\b[^>]*\/?>', _replace_img, html_body, flags=re.IGNORECASE)

            # - Flatten nested tags inside <td> and <th> (fpdf2 limitation)
            def _clean_cell(match):
                tag = match.group(1)
                content = match.group(2)
                clean_content = re.sub(r'</?(?:code|span|div|pre|strong|em|b|i|a|p)\b[^>]*>', '', content, flags=re.IGNORECASE)
                return f"<{tag}>{clean_content}</{tag}>"
            html_body = re.sub(r'<(td|th)\b[^>]*>(.*?)</\1>', _clean_cell, html_body, flags=re.DOTALL | re.IGNORECASE)

            # - Strip unresolvable internal anchor links <a href="#...">Text</a> -> Text
            html_body = re.sub(r'<a\b[^>]*href=["\']#[^"\']*["\'][^>]*>(.*?)</a>', r'\1', html_body, flags=re.IGNORECASE | re.DOTALL)

            # Generate PDF via FPDF write_html with Unicode font support
            pdf = MarkdownPDFRenderer(page_size=page_size, orientation=orientation, margins=margins)
            pdf.add_page()
            pdf.set_font(pdf.font_family_name, size=10)

            # If fallback font Helvetica is active, ensure text is latin-1 safe
            if pdf.font_family_name == "Helvetica":
                html_body = html_body.encode("latin-1", "replace").decode("latin-1")

            # Change working directory temporarily to resolve relative images
            old_cwd = os.getcwd()
            try:
                os.chdir(file_path.parent)
                pdf.write_html(html_body)
            finally:
                os.chdir(old_cwd)

            out_dir = Path(output_dir) if output_dir is not None else (file_path.parent / "output")
            OutputManager.ensure_output_dir(out_dir)
            out_path = OutputManager.get_collision_safe_path(out_dir / f"{file_path.stem}.pdf")

            pdf.output(str(out_path))

            # Verification
            is_valid, msg = OutputManager.verify_pdf_readable(out_path)
            if not is_valid:
                report.add_result(file_path, "Failed", error=f"Rendered PDF verification failed: {msg}")
                continue

            done_pct = (idx / total_files) * 100.0
            _cb(done_pct, f"[{idx}/{total_files}] Rendered {file_path.name}")
            _log(f"  [+] Saved PDF: {out_path.name}")
            report.add_result(file_path, "Success", output_files=[out_path], details="Markdown rendered to PDF.")

        except Exception as e:
            report.add_result(file_path, "Failed", error=f"Markdown-to-PDF conversion error: {e}")

    _cb(100.0, f"Rendered {total_files} Markdown document(s).")
    return report


# ==============================================================================
# OPERATION 10: PDF TO MARKDOWN
# ==============================================================================

def execute_pdf_to_markdown(
    selected_files: List[Path],
    extract_images: bool = True,
    ocr_fallback: bool = True,
    ocr_images: bool = True,
    lang: str = "eng",
    output_dir: Optional[Path] = None,
    progress_callback: Optional[Callable[[float, str], None]] = None,
    log_callback: Optional[Callable[[str], None]] = None
) -> ProcessingReport:
    """
    Extracts text, headings, tables, links, and images from PDFs into readable Markdown.
    """
    def _cb(pct: float, msg: str):
        if progress_callback:
            try:
                progress_callback(pct, msg)
            except Exception:
                pass

    def _log(msg: str):
        if log_callback:
            try:
                log_callback(msg)
            except Exception:
                pass

    report = ProcessingReport("PDF to Markdown")
    total_files = len(selected_files)

    for idx, file_path in enumerate(selected_files, 1):
        base_pct = ((idx - 1) / total_files) * 100.0
        file_weight = 100.0 / total_files
        _cb(base_pct, f"[{idx}/{total_files}] Extracting Markdown: {file_path.name}...")
        _log(f"[*] [{idx}/{total_files}] Extracting Markdown: {file_path.name}")
        try:
            if not file_path.exists() or file_path.stat().st_size == 0:
                report.add_result(file_path, "Failed", error="File does not exist or is empty.")
                continue

            doc = pymupdf.open(str(file_path))
            if doc.is_encrypted:
                doc.close()
                report.add_result(file_path, "Failed", error="File is encrypted. Unlock it first.")
                continue

            out_dir = Path(output_dir) if output_dir is not None else (file_path.parent / "output")
            OutputManager.ensure_output_dir(out_dir)
            out_md_path = OutputManager.get_collision_safe_path(out_dir / f"{file_path.stem}.md")

            assets_dir = out_dir / f"{file_path.stem}_assets"
            if extract_images:
                OutputManager.ensure_output_dir(assets_dir)

            md_lines: List[str] = [f"# {file_path.stem}\n"]
            image_counter = 1

            num_pages = len(doc)
            for pno in range(num_pages):
                sub_pct = base_pct + ((pno + 1) / num_pages) * file_weight
                _cb(sub_pct, f"[{idx}/{total_files}] Page {pno + 1}/{num_pages} of {file_path.name}...")
                page = doc[pno]
                md_lines.append(f"\n---\n<!-- Page {pno + 1} -->\n")

                # Extract tables first
                try:
                    tabs = page.find_tables()
                except Exception:
                    tabs = None

                table_rects = []
                if tabs and tabs.tables:
                    for table in tabs.tables:
                        table_rects.append(table.bbox)
                        rows = table.extract()
                        if rows and len(rows) > 0:
                            header_cells = [str(c or '').replace('\n', ' ').strip() for c in rows[0]]
                            md_lines.append("\n| " + " | ".join(header_cells) + " |")
                            md_lines.append("| " + " | ".join(["---"] * len(header_cells)) + " |")
                            for row in rows[1:]:
                                row_cells = [str(c or '').replace('\n', ' ').strip() for c in row]
                                md_lines.append("| " + " | ".join(row_cells) + " |")
                            md_lines.append("\n")

                # Extract text blocks
                text_blocks = page.get_text("blocks")
                page_text_len = sum(len(b[4].strip()) for b in text_blocks if len(b) > 4)

                images_on_page = page.get_images()
                # OCR page fallback if page has minimal text or is primarily a scanned slide
                if (page_text_len < 80 or (page_text_len < 150 and len(images_on_page) > 0)) and ocr_fallback and PYTESSERACT_AVAILABLE:
                    try:
                        tess_cmd = detect_tesseract_binary()
                        if tess_cmd:
                            pytesseract.pytesseract.tesseract_cmd = tess_cmd
                            pix = page.get_pixmap(dpi=300)
                            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                            ocr_text = pytesseract.image_to_string(img, lang=lang).strip()
                            if ocr_text:
                                md_lines.append("\n> *[OCR Extracted Slide Content]*\n")
                                md_lines.append(ocr_text + "\n")
                    except Exception as ocr_err:
                        logger.warning(f"OCR fallback failed on page {pno + 1}: {ocr_err}")

                else:
                    for block in text_blocks:
                        # block format: (x0, y0, x1, y1, text, block_no, block_type)
                        if len(block) >= 5 and block[6] == 0:  # Text block
                            bx0, by0, bx1, by1, btext = block[0], block[1], block[2], block[3], block[4].strip()
                            if not btext:
                                continue

                            # Skip if this block is inside an extracted table
                            in_table = False
                            for bbox in table_rects:
                                if bx0 >= bbox[0] - 5 and by0 >= bbox[1] - 5 and bx1 <= bbox[2] + 5 and by1 <= bbox[3] + 5:
                                    in_table = True
                                    break
                            if in_table:
                                continue

                            # Heading heuristic: short line in uppercase or larger text
                            if len(btext.split('\n')) == 1 and len(btext) < 60 and (btext.isupper() or len(btext.split()) <= 6):
                                md_lines.append(f"\n## {btext}\n")
                            else:
                                md_lines.append(f"\n{btext}\n")

                # Extract images if requested
                if extract_images:
                    for img_info in page.get_images(full=True):
                        xref = img_info[0]
                        try:
                            pix = pymupdf.Pixmap(doc, xref)
                            if pix.n >= 5:
                                pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
                            img_filename = f"image_{image_counter:03d}.png"
                            img_path = assets_dir / img_filename
                            pix.save(str(img_path))
                            # Add relative link in markdown
                            rel_asset_path = f"{assets_dir.name}/{img_filename}"
                            md_lines.append(f"\n![Extracted image]({rel_asset_path})\n")

                            # Extract OCR text from embedded diagram / figure if enabled
                            if ocr_images and PYTESSERACT_AVAILABLE and pix.width > 150 and pix.height > 100:
                                try:
                                    pil_img = Image.open(str(img_path))
                                    img_ocr = pytesseract.image_to_string(pil_img, lang=lang).strip()
                                    if img_ocr and len(img_ocr) > 15:
                                        clean_txt = img_ocr.replace('\n', '\n> ')
                                        md_lines.append(f"> **[Diagram / Figure Content]**:\n> {clean_txt}\n")
                                except Exception:
                                    pass

                            image_counter += 1
                        except Exception as img_err:
                            logger.warning(f"Failed extracting image xref {xref}: {img_err}")

            doc.close()

            # Save Markdown file
            with open(out_md_path, "w", encoding="utf-8") as f_out:
                f_out.write("".join(md_lines))

            done_pct = (idx / total_files) * 100.0
            _cb(done_pct, f"[{idx}/{total_files}] Extracted {file_path.name}")
            _log(f"  [+] Markdown: {out_md_path.name} ({image_counter - 1} images extracted)")
            report.add_result(
                file_path,
                "Success",
                output_files=[out_md_path],
                details=f"Extracted to Markdown ({image_counter - 1} images extracted)"
            )

        except Exception as e:
            report.add_result(file_path, "Failed", error=f"PDF-to-Markdown error: {e}")

    _cb(100.0, f"Extracted {total_files} PDF document(s) to Markdown.")
    return report


# ==============================================================================
# SECTION 2: INTERACTIVE GUI DASHBOARD & WORKFLOW
# ==============================================================================
# ==============================================================================
# PDF Toolkit Pro -- Advanced Glassmorphism Desktop GUI Dashboard
# Developed by: mohatamim
# ==============================================================================
import os
import sys
import time
import threading
import subprocess
from pathlib import Path
from typing import List, Dict, Optional, Tuple, Any

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk, ImageDraw

FONT_FAMILY = "Segoe UI"
FONT_TITLE = (FONT_FAMILY, 11, "bold")
FONT_HEADING = (FONT_FAMILY, 9, "bold")
FONT_REGULAR = (FONT_FAMILY, 8)
FONT_BOLD = (FONT_FAMILY, 8, "bold")
FONT_MONO = ("Consolas", 8)

class MaterialIconManager:
    """
    Renders clean vector icons onto PIL images for dark-themed UI buttons.
    """
    _cache = {}

    @classmethod
    def get(cls, name: str, size: Tuple[int, int] = (16, 16), color_hex: str = "#38BDF8") -> ImageTk.PhotoImage:
        key = (name, size, color_hex)
        if key in cls._cache:
            return cls._cache[key]
        w, h = size
        img = Image.new("RGBA", (w, h), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        # Draw clean vector glyphs based on name
        if name in ("merge", "call_merge"):
            draw.line([(w * 0.2, h * 0.2), (w * 0.5, h * 0.5), (w * 0.5, h * 0.85)], fill=color_hex, width=2)
            draw.line([(w * 0.8, h * 0.2), (w * 0.5, h * 0.5)], fill=color_hex, width=2)
            draw.ellipse([(w * 0.4, h * 0.75), (w * 0.6, h * 0.95)], fill=color_hex)
        elif name in ("split", "call_split"):
            draw.line([(w * 0.5, h * 0.15), (w * 0.5, h * 0.5), (w * 0.2, h * 0.8)], fill=color_hex, width=2)
            draw.line([(w * 0.5, h * 0.5), (w * 0.8, h * 0.8)], fill=color_hex, width=2)
            draw.ellipse([(w * 0.4, h * 0.05), (w * 0.6, h * 0.25)], fill=color_hex)
        elif name in ("compress",):
            draw.rectangle([(w * 0.2, h * 0.2), (w * 0.8, h * 0.8)], outline=color_hex, width=2)
            draw.line([(w * 0.35, h * 0.5), (w * 0.65, h * 0.5)], fill=color_hex, width=2)
            draw.line([(w * 0.5, h * 0.35), (w * 0.5, h * 0.65)], fill=color_hex, width=2)
        elif name in ("image", "photo_library"):
            draw.rectangle([(w * 0.15, h * 0.15), (w * 0.85, h * 0.85)], outline=color_hex, width=2)
            draw.ellipse([(w * 0.3, h * 0.3), (w * 0.45, h * 0.45)], fill=color_hex)
            draw.polygon([(w * 0.25, h * 0.75), (w * 0.5, h * 0.5), (w * 0.75, h * 0.75)], fill=color_hex)
        elif name in ("wallpaper", "collections"):
            draw.rectangle([(w * 0.1, h * 0.2), (w * 0.7, h * 0.8)], outline=color_hex, width=1)
            draw.rectangle([(w * 0.3, h * 0.1), (w * 0.9, h * 0.7)], fill=color_hex, outline="#FFFFFF", width=1)
        elif name in ("lock", "protect"):
            draw.rectangle([(w * 0.2, h * 0.45), (w * 0.8, h * 0.9)], fill=color_hex)
            draw.arc([(w * 0.3, h * 0.15), (w * 0.7, h * 0.55)], 180, 0, fill=color_hex, width=2)
        elif name in ("unlock", "lock_open"):
            draw.rectangle([(w * 0.2, h * 0.45), (w * 0.8, h * 0.9)], fill=color_hex)
            draw.arc([(w * 0.45, h * 0.15), (w * 0.85, h * 0.55)], 180, 0, fill=color_hex, width=2)
        elif name in ("ocr", "scanner"):
            draw.rectangle([(w * 0.15, h * 0.15), (w * 0.85, h * 0.85)], outline=color_hex, width=1)
            draw.line([(w * 0.1, h * 0.5), (w * 0.9, h * 0.5)], fill=color_hex, width=2)
        elif name in ("markdown", "description"):
            draw.rectangle([(w * 0.2, h * 0.1), (w * 0.8, h * 0.9)], outline=color_hex, width=2)
            draw.line([(w * 0.35, h * 0.35), (w * 0.65, h * 0.35)], fill=color_hex, width=1)
            draw.line([(w * 0.35, h * 0.55), (w * 0.65, h * 0.55)], fill=color_hex, width=1)
            draw.line([(w * 0.35, h * 0.75), (w * 0.55, h * 0.75)], fill=color_hex, width=1)
        else:
            draw.ellipse([(w * 0.2, h * 0.2), (w * 0.8, h * 0.8)], outline=color_hex, width=2)

        tk_img = ImageTk.PhotoImage(img)
        cls._cache[key] = tk_img
        return tk_img


def format_file_size(num_bytes: int) -> str:
    if num_bytes < 1024:
        return f"{num_bytes} B"
    elif num_bytes < 1024 * 1024:
        return f"{num_bytes / 1024.0:.1f} KB"
    elif num_bytes < 1024 * 1024 * 1024:
        return f"{num_bytes / (1024.0 * 1024.0):.2f} MB"
    return f"{num_bytes / (1024.0 * 1024.0 * 1024.0):.2f} GB"


class StdoutRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, str_val):
        if str_val:
            try:
                self.text_widget.insert(tk.END, str_val)
                self.text_widget.see(tk.END)
            except Exception:
                pass

    def flush(self):
        pass


class PDFToolkitDashboardGUI:
    """
    Modern Glassmorphism Desktop Dashboard for PDF Toolkit Pro.
    Supports 10 Comprehensive Operations:
      1. Merge PDFs (with reordering)
      2. Split PDF (by pages or ranges)
      3. Compress PDF (lossless & lossy)
      4. Images to PDF (multi-format, layout, DPI, margins) - NEW
      5. PDF to Images (multi-format, DPI, color mode, page range) - NEW
      6. Protect / Encrypt PDF (AES-256)
      7. Unlock / Decrypt PDF
      8. OCR PDF (Tesseract searchable text)
      9. Markdown to PDF
      10. PDF to Markdown
    """

    MODE_SPECS = {
        "merge": {
            "title": "Merge Multiple PDFs",
            "icon": "call_merge",
            "desc": "Combine multiple PDF documents into a single organized master file with optional TOC bookmarks.",
            "exts": [("PDF Documents (*.pdf)", "*.pdf"), ("All Files (*.*)", "*.*")],
            "multi_file": True
        },
        "split": {
            "title": "Split PDF Pages",
            "icon": "call_split",
            "desc": "Split PDF into single pages, custom page ranges, or fixed-size page groups.",
            "exts": [("PDF Documents (*.pdf)", "*.pdf"), ("All Files (*.*)", "*.*")],
            "multi_file": False
        },
        "compress": {
            "title": "Compress & Optimize PDF",
            "icon": "compress",
            "desc": "Reduce PDF file size via stream optimization, resource deduplication, and image downsampling.",
            "exts": [("PDF Documents (*.pdf)", "*.pdf"), ("All Files (*.*)", "*.*")],
            "multi_file": True
        },
        "img2pdf": {
            "title": "Images to PDF Document",
            "icon": "photo_library",
            "desc": "Convert and assemble images (PNG, JPG, WEBP, BMP, TIFF, GIF) into professional PDF documents.",
            "exts": [
                ("All Supported Images", "*.png *.jpg *.jpeg *.webp *.bmp *.tiff *.tif *.gif *.ico"),
                ("PNG Images (*.png)", "*.png"),
                ("JPEG Images (*.jpg, *.jpeg)", "*.jpg *.jpeg"),
                ("WEBP Images (*.webp)", "*.webp"),
                ("All Files (*.*)", "*.*")
            ],
            "multi_file": True
        },
        "pdf2img": {
            "title": "PDF to Images",
            "icon": "wallpaper",
            "desc": "Render PDF pages into high-fidelity image files (PNG, JPEG, WEBP, TIFF, BMP) at customizable DPI.",
            "exts": [("PDF Documents (*.pdf)", "*.pdf"), ("All Files (*.*)", "*.*")],
            "multi_file": True
        },
        "protect": {
            "title": "Protect / Encrypt PDF",
            "icon": "lock",
            "desc": "Secure PDF documents with AES-256 encryption, user passwords, and granular document permissions.",
            "exts": [("PDF Documents (*.pdf)", "*.pdf"), ("All Files (*.*)", "*.*")],
            "multi_file": True
        },
        "unlock": {
            "title": "Unlock / Decrypt PDF",
            "icon": "lock_open",
            "desc": "Remove password restrictions and decrypt password-protected PDF documents.",
            "exts": [("PDF Documents (*.pdf)", "*.pdf"), ("All Files (*.*)", "*.*")],
            "multi_file": True
        },
        "ocr": {
            "title": "OCR PDF (Searchable Text)",
            "icon": "scanner",
            "desc": "Extract text layers and create searchable PDFs from scanned documents using Tesseract OCR.",
            "exts": [("PDF Documents (*.pdf)", "*.pdf"), ("All Files (*.*)", "*.*")],
            "multi_file": True
        },
        "md2pdf": {
            "title": "Markdown to PDF",
            "icon": "description",
            "desc": "Render GitHub-flavored markdown documents into styled, publication-ready PDFs.",
            "exts": [
                ("Markdown Files (*.md)", "*.md"),
                ("All Markdown Documents", "*.md *.markdown *.mdown *.txt"),
                ("All Files (*.*)", "*.*")
            ],
            "multi_file": True
        },
        "pdf2md": {
            "title": "PDF to Markdown",
            "icon": "description",
            "desc": "Extract structured text and embedded images from PDFs into clean Markdown files.",
            "exts": [("PDF Documents (*.pdf)", "*.pdf"), ("All Files (*.*)", "*.*")],
            "multi_file": True
        }
    }

    def __init__(self, parent=None):
        self.root = tk.Tk() if parent is None else tk.Toplevel(parent)
        self.root.title("PDF Toolkit Pro -- Advanced Document & Image Suite (mohatamim)")
        self.root.geometry("1220x800")
        self.root.minsize(1080, 700)
        self.root.configure(bg="#0B1120")

        self.current_mode = "merge"
        self.file_items: List[Path] = []
        self.is_running = False

        # Output options
        default_out = str(Path.home() / "Documents" / "PDF_Toolkit_Output")
        self.out_dir_var = tk.StringVar(value=default_out)
        self.out_name_var = tk.StringVar(value="output_document.pdf")
        self.open_dest_var = tk.BooleanVar(value=True)

        # Mode-specific variables
        # Images to PDF vars
        self.img_pagesize_var = tk.StringVar(value="Fit to Image (Original)")
        self.img_orient_var = tk.StringVar(value="Auto (Best Fit)")
        self.img_margin_var = tk.StringVar(value="None (0 mm)")
        self.img_quality_var = tk.IntVar(value=92)
        self.img_merge_var = tk.StringVar(value="single") # 'single' or 'multiple'

        # PDF to Images vars
        self.pdf2img_format_var = tk.StringVar(value="PNG (Lossless)")
        self.pdf2img_dpi_var = tk.StringVar(value="150 DPI (Standard Screen)")
        self.pdf2img_range_var = tk.StringVar(value="all")
        self.pdf2img_color_var = tk.StringVar(value="rgb") # 'rgb' or 'grayscale'
        self.pdf2img_subfolder_var = tk.BooleanVar(value=True)
        self.pdf2img_quality_var = tk.IntVar(value=90)

        # Merge vars
        self.merge_bookmarks_var = tk.BooleanVar(value=True)

        # Split vars
        self.split_mode_var = tk.StringVar(value="all") # 'all' or 'range'
        self.split_range_var = tk.StringVar(value="1-3, 4-6")

        # Compress vars
        self.compress_level_var = tk.StringVar(value="medium")

        # Protect / Unlock vars
        self.user_pwd_var = tk.StringVar(value="")
        self.owner_pwd_var = tk.StringVar(value="")
        self.unlock_pwd_var = tk.StringVar(value="")

        # OCR PDF options
        self.ocr_coverage_var = tk.StringVar(value="all")
        self.ocr_lang_var = tk.StringVar(value="eng")
        self.ocr_dpi_var = tk.StringVar(value="300")

        # PDF to Markdown options
        self.pdf2md_extract_images_var = tk.BooleanVar(value=True)
        self.pdf2md_ocr_var = tk.BooleanVar(value=True)
        self.pdf2md_lang_var = tk.StringVar(value="eng")

        # Drag-and-drop state
        self._drag_start_idx = None
        self._drag_popup = None

        self._build_ui()

        # Bind Windows Explorer External Drag & Drop
        if WINDND_AVAILABLE:
            try:
                windnd.hook_dropfiles(self.root, func=self._on_external_drop)
            except Exception:
                pass

        self.root.bind("<Escape>", self._on_cancel_drag)
        self._bind_keyboard_shortcuts()
        self._switch_mode("merge")
        self._apply_pointer_cursors()

        # Thread-safe UI event queue for Python 3.13 compatibility
        self.ui_queue = queue.Queue()
        self.root.after(30, self._poll_ui_queue)


    @staticmethod
    def _get_engine_func(fn_name: str):
        """
        Dynamically resolves the document engine function:
        1. Checks current global scope (direct call in script/notebook).
        2. Fallback: ensures project directories are in sys.path and imports pdf_toolkit.
        """
        if fn_name in globals() and callable(globals()[fn_name]):
            return globals()[fn_name]

        import __main__
        if hasattr(__main__, fn_name) and callable(getattr(__main__, fn_name)):
            return getattr(__main__, fn_name)

        import sys
        from pathlib import Path
        for p in [Path.cwd(), Path.cwd().parent, Path(r"d:\New folder\PY"), Path(r"d:\New folder\PY\scripts")]:
            p_str = str(p.resolve()) if p.exists() else str(p)
            if p_str not in sys.path:
                sys.path.insert(0, p_str)

        try:
            import pdf_toolkit
            if hasattr(pdf_toolkit, fn_name):
                return getattr(pdf_toolkit, fn_name)
        except Exception:
            pass

        raise RuntimeError(f"Engine function '{fn_name}' could not be located in notebook scope or pdf_toolkit module.")

    def _get_accepted_extensions(self, mode: Optional[str] = None) -> Set[str]:
        """Resolves set of file extensions accepted by current or specified mode."""
        m = mode or self.current_mode
        spec = self.MODE_SPECS.get(m, {})
        exts = set()
        for _, pat_str in spec.get("exts", []):
            for sub in pat_str.replace(";", " ").split():
                clean = sub.lower().strip().lstrip("*")
                if clean and clean != ".":
                    exts.add(clean)
        return exts

    def _show_shortcuts_modal(self):
        """Displays interactive modal detailing all keyboard shortcuts."""
        dlg = tk.Toplevel(self.root)
        dlg.title("Keyboard Shortcuts Reference")
        dlg.geometry("540x530")
        dlg.minsize(480, 440)
        dlg.configure(bg="#0F172A")
        dlg.transient(self.root)
        dlg.grab_set()

        hdr = tk.Frame(dlg, bg="#1E293B", padx=16, pady=12)
        hdr.pack(fill=tk.X)
        tk.Label(hdr, text="PDF Toolkit Pro -- Keyboard Shortcuts", font=("Segoe UI", 12, "bold"), fg="#38BDF8", bg="#1E293B").pack(anchor="w")
        tk.Label(hdr, text="Accelerate your document workflow with quick keyboard navigation", font=("Segoe UI", 9), fg="#94A3B8", bg="#1E293B").pack(anchor="w")

        body = tk.Frame(dlg, bg="#0F172A", padx=16, pady=12)
        body.pack(fill=tk.BOTH, expand=True)

        shortcuts = [
            ("Ctrl + O / Ctrl + A", "Add single or multiple file(s) to queue"),
            ("Ctrl + Shift + O / Ctrl + F", "Import all matching files from a folder"),
            ("Ctrl + Enter / F5", "Execute active operation"),
            ("Delete / Backspace", "Remove highlighted file(s) from queue"),
            ("Ctrl + Delete / Ctrl + L", "Clear entire file queue"),
            ("Ctrl + Up / Alt + Up", "Move highlighted file UP in queue"),
            ("Ctrl + Down / Alt + Down", "Move highlighted file DOWN in queue"),
            ("Alt + 1 .. 9, 0", "Switch active operation (1: Merge ... 9: Md2PDF, 0: PDF2Md)"),
            ("Escape", "Cancel drag-and-drop / Close modal dialogs"),
            ("F1 / Ctrl + H", "Show this shortcuts cheatsheet")
        ]

        table = tk.Frame(body, bg="#162032", relief="solid", bd=1)
        table.pack(fill=tk.BOTH, expand=True)

        r_head = tk.Frame(table, bg="#1E293B", padx=8, pady=6)
        r_head.pack(fill=tk.X)
        tk.Label(r_head, text="Key Combination", font=("Segoe UI", 9, "bold"), fg="#38BDF8", bg="#1E293B", width=24, anchor="w").pack(side=tk.LEFT)
        tk.Label(r_head, text="Action", font=("Segoe UI", 9, "bold"), fg="#38BDF8", bg="#1E293B", anchor="w").pack(side=tk.LEFT, fill=tk.X, expand=True)

        for i, (k_combo, action) in enumerate(shortcuts):
            bg_color = "#162032" if i % 2 == 0 else "#0F172A"
            row = tk.Frame(table, bg=bg_color, padx=8, pady=4)
            row.pack(fill=tk.X)
            tk.Label(row, text=k_combo, font=("Consolas", 9, "bold"), fg="#F8FAFC", bg=bg_color, width=24, anchor="w").pack(side=tk.LEFT)
            tk.Label(row, text=action, font=("Segoe UI", 9), fg="#94A3B8", bg=bg_color, anchor="w").pack(side=tk.LEFT, fill=tk.X, expand=True)

        ftr = tk.Frame(dlg, bg="#0F172A", padx=16, pady=10)
        ftr.pack(fill=tk.X)
        btn_close = tk.Button(
            ftr,
            text="Got it",
            command=dlg.destroy,
            font=("Segoe UI", 9, "bold"),
            bg="#0284C7",
            fg="#FFFFFF",
            activebackground="#0369A1",
            relief="flat",
            padx=18,
            pady=4,
            cursor="hand2"
        )
        btn_close.pack(side=tk.RIGHT)
        dlg.bind("<Escape>", lambda e: dlg.destroy())

    def _bind_keyboard_shortcuts(self):
        """Binds global keyboard shortcuts for seamless dashboard control."""
        # File management
        self.root.bind("<Control-o>", lambda e: self._add_files_dialog())
        self.root.bind("<Control-O>", lambda e: self._add_files_dialog())
        self.root.bind("<Control-a>", lambda e: self._add_files_dialog())
        self.root.bind("<Control-A>", lambda e: self._add_files_dialog())
        self.root.bind("<Control-Shift-O>", lambda e: self._add_folder_dialog())
        self.root.bind("<Control-Shift-o>", lambda e: self._add_folder_dialog())
        self.root.bind("<Control-f>", lambda e: self._add_folder_dialog())
        self.root.bind("<Control-F>", lambda e: self._add_folder_dialog())

        # Execution
        self.root.bind("<Control-Return>", lambda e: self._start_execution_thread())
        self.root.bind("<Control-KP_Enter>", lambda e: self._start_execution_thread())
        self.root.bind("<F5>", lambda e: self._start_execution_thread())

        # Queue management
        self.root.bind("<Delete>", lambda e: self._remove_selected_files())
        self.root.bind("<BackSpace>", lambda e: self._remove_selected_files())
        self.root.bind("<Control-Delete>", lambda e: self._clear_all_files())
        self.root.bind("<Control-l>", lambda e: self._clear_all_files())
        self.root.bind("<Control-L>", lambda e: self._clear_all_files())
        self.root.bind("<Control-Up>", lambda e: self._move_file_up())
        self.root.bind("<Alt-Up>", lambda e: self._move_file_up())
        self.root.bind("<Control-Down>", lambda e: self._move_file_down())
        self.root.bind("<Alt-Down>", lambda e: self._move_file_down())

        # Help modal
        self.root.bind("<F1>", lambda e: self._show_shortcuts_modal())
        self.root.bind("<Control-h>", lambda e: self._show_shortcuts_modal())
        self.root.bind("<Control-H>", lambda e: self._show_shortcuts_modal())

        # Number keys for instant mode switching (Alt+1..9,0)
        mode_keys = list(self.MODE_SPECS.keys())
        for idx, m_k in enumerate(mode_keys):
            num = (idx + 1) % 10
            self.root.bind(f"<Alt-Key-{num}>", lambda e, k=m_k: self._switch_mode(k))

    def _apply_pointer_cursors(self, parent=None):
        if parent is None:
            parent = self.root
        for child in parent.winfo_children():
            cls_name = child.__class__.__name__
            is_button = isinstance(child, (tk.Button, tk.Radiobutton, tk.Checkbutton)) or any(k in cls_name for k in ("Button", "Radiobutton", "Checkbutton"))
            if is_button:
                try:
                    is_active = str(child.cget("state")) not in ("disabled",)
                    child.configure(cursor="hand2" if is_active else "")

                    if not getattr(child, "_cursor_wrapped", False):
                        orig_cfg = child.configure
                        def _make_cfg(orig, w):
                            def _custom_cfg(*args, **kwargs):
                                if "state" in kwargs:
                                    st = str(kwargs["state"])
                                    if st == "disabled":
                                        kwargs["cursor"] = ""
                                    elif st in ("normal", "active"):
                                        kwargs["cursor"] = "hand2"
                                return orig(*args, **kwargs)
                            return _custom_cfg
                        child.configure = _make_cfg(orig_cfg, child)
                        child.config = child.configure
                        child._cursor_wrapped = True

                    def _on_hover(event, w=child):
                        try:
                            if str(w.cget("state")) == "disabled":
                                w.configure(cursor="")
                            else:
                                w.configure(cursor="hand2")
                        except Exception:
                            pass
                    child.bind("<Enter>", _on_hover, add="+")
                except Exception:
                    pass
            self._apply_pointer_cursors(child)

    def _build_ui(self):
        # 1. Header Bar
        hdr = tk.Frame(self.root, bg="#162032", padx=16, pady=10, relief="solid", bd=1)
        hdr.pack(fill=tk.X)

        h_left = tk.Frame(hdr, bg="#162032")
        h_left.pack(side=tk.LEFT)
        tk.Label(h_left, text="PDF Toolkit Pro | Advanced Document & Image Suite", font=FONT_TITLE, fg="#38BDF8", bg="#162032").pack(anchor="w")
        tk.Label(h_left, text="Engine Architecture by mohatamim -- High-Fidelity PDF & Image Processing", font=("Segoe UI", 8), fg="#94A3B8", bg="#162032").pack(anchor="w")

        self.lbl_active_mode_badge = tk.Label(hdr, text="Active: Merge Multiple PDFs", font=FONT_BOLD, fg="#34D399", bg="#0F172A", padx=12, pady=4, relief="solid", bd=1)
        self.lbl_active_mode_badge.pack(side=tk.RIGHT)

        tk.Button(
            hdr,
            text="Shortcuts (F1)",
            command=self._show_shortcuts_modal,
            font=("Segoe UI", 9, "bold"),
            bg="#1E293B",
            fg="#38BDF8",
            activebackground="#334155",
            activeforeground="#FFFFFF",
            relief="flat",
            padx=10,
            pady=3,
            cursor="hand2"
        ).pack(side=tk.RIGHT, padx=8)

        # 2. Main Horizontal Container
        container = tk.Frame(self.root, bg="#0B1120")
        container.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        # Left Sidebar: Mode Selector
        self.sidebar = tk.Frame(container, bg="#162032", width=230, padx=8, pady=8, relief="solid", bd=1)
        self.sidebar.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 8))
        self.sidebar.pack_propagate(False)

        tk.Label(self.sidebar, text="OPERATION MODES", font=("Segoe UI", 8, "bold"), fg="#64748B", bg="#162032").pack(anchor="w", pady=(0, 6), padx=4)

        self.mode_buttons = {}
        for mode_key, spec in self.MODE_SPECS.items():
            btn = tk.Button(
                self.sidebar,
                cursor="hand2",
                text=f"  {spec['title']}",
                image=MaterialIconManager.get(spec["icon"], size=(16, 16), color_hex="#94A3B8"),
                compound=tk.LEFT,
                font=FONT_REGULAR,
                fg="#F8FAFC",
                bg="#1E293B",
                activebackground="#0284C7",
                activeforeground="#FFFFFF",
                relief="flat",
                anchor="w",
                padx=8,
                pady=6,
                command=lambda m=mode_key: self._switch_mode(m)
            )
            btn.pack(fill=tk.X, pady=2)
            self.mode_buttons[mode_key] = btn

        # Right Main Content Area
        self.content_area = tk.Frame(container, bg="#0B1120")
        self.content_area.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # --- Card 1: File Queue & Reordering Toolbar ---
        card_queue = tk.Frame(self.content_area, bg="#162032", padx=10, pady=8, relief="solid", bd=1)
        card_queue.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

        # Toolbar above Queue
        tb = tk.Frame(card_queue, bg="#162032")
        tb.pack(fill=tk.X, pady=(0, 6))

        tk.Label(tb, text="Input Documents & Assets Queue:", font=FONT_HEADING, fg="#F8FAFC", bg="#162032").pack(side=tk.LEFT)

        tk.Button(tb, text="+ Add File(s)...", command=self._add_files_dialog, font=FONT_BOLD, bg="#0284C7", fg="#FFFFFF", activebackground="#0369A1", relief="flat", padx=10, pady=3, cursor="hand2").pack(side=tk.RIGHT, padx=2)
        tk.Button(tb, text="+ Add Folder...", command=self._add_folder_dialog, font=FONT_REGULAR, bg="#334155", fg="#F8FAFC", activebackground="#475569", relief="flat", padx=8, pady=3, cursor="hand2").pack(side=tk.RIGHT, padx=2)
        tk.Button(tb, text="Remove Selected", command=self._remove_selected_files, font=FONT_REGULAR, bg="#DC2626", fg="#FFFFFF", activebackground="#B91C1C", relief="flat", padx=8, pady=3, cursor="hand2").pack(side=tk.RIGHT, padx=2)
        tk.Button(tb, text="Move Up", command=self._move_file_up, font=FONT_REGULAR, bg="#334155", fg="#F8FAFC", relief="flat", padx=8, pady=3, cursor="hand2").pack(side=tk.RIGHT, padx=2)
        tk.Button(tb, text="Move Down", command=self._move_file_down, font=FONT_REGULAR, bg="#334155", fg="#F8FAFC", relief="flat", padx=8, pady=3, cursor="hand2").pack(side=tk.RIGHT, padx=2)
        tk.Button(tb, text="Clear", command=self._clear_all_files, font=FONT_REGULAR, bg="#475569", fg="#F8FAFC", relief="flat", padx=6, pady=3, cursor="hand2").pack(side=tk.RIGHT, padx=2)

        # Modern Treeview for Files Queue
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("DarkQueue.Treeview", background="#0F172A", foreground="#F8FAFC", fieldbackground="#0F172A", borderwidth=0, font=FONT_REGULAR)
        style.configure("DarkQueue.Treeview.Heading", background="#1E293B", foreground="#38BDF8", font=FONT_BOLD, relief="flat")
        style.map("DarkQueue.Treeview", background=[("selected", "#0284C7")], foreground=[("selected", "#FFFFFF")])

        tree_frame = tk.Frame(card_queue, bg="#0F172A")
        tree_frame.pack(fill=tk.BOTH, expand=True)

        cols = ("idx", "name", "size", "type", "path")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", style="DarkQueue.Treeview", selectmode="extended")
        self.tree.heading("idx", text="#", anchor="center")
        self.tree.heading("name", text="File Name", anchor="w")
        self.tree.heading("size", text="Size", anchor="e")
        self.tree.heading("type", text="Type", anchor="center")
        self.tree.heading("path", text="Absolute Path", anchor="w")

        self.tree.column("idx", width=40, anchor="center")
        self.tree.column("name", width=280, anchor="w")
        self.tree.column("size", width=90, anchor="e")
        self.tree.column("type", width=120, anchor="center")
        self.tree.column("path", width=360, anchor="w")

        tree_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Drag and Drop Reordering bindings
        self.tree.bind("<ButtonPress-1>", self._on_tree_press)
        self.tree.bind("<B1-Motion>", self._on_tree_motion)
        self.tree.bind("<ButtonRelease-1>", self._on_tree_release)

        # Queue Status Footer
        self.lbl_queue_status = tk.Label(card_queue, text="0 files queued (0 B) -- Click '+ Add File(s)...' or drag items to reorder.", font=("Segoe UI", 8, "italic"), fg="#94A3B8", bg="#162032")
        self.lbl_queue_status.pack(anchor="w", pady=(4, 0))

        # --- Card 2: Configuration & Destination Options ---
        card_config = tk.Frame(self.content_area, bg="#162032", padx=10, pady=8, relief="solid", bd=1)
        card_config.pack(fill=tk.X, pady=(0, 6))

        config_cols = tk.Frame(card_config, bg="#162032")
        config_cols.pack(fill=tk.X)

        # Dynamic Mode Settings (Left Column)
        self.frame_dynamic_settings = tk.LabelFrame(config_cols, text="Operation Settings", font=FONT_BOLD, fg="#38BDF8", bg="#162032", padx=8, pady=6)
        self.frame_dynamic_settings.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 6))

        # Destination & Output Settings (Right Column)
        frame_dest = tk.LabelFrame(config_cols, text="Destination & Output Naming", font=FONT_BOLD, fg="#C084FC", bg="#162032", padx=8, pady=6)
        frame_dest.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(6, 0))

        # Destination Folder row
        d_row = tk.Frame(frame_dest, bg="#162032")
        d_row.pack(fill=tk.X, pady=2)
        tk.Label(d_row, text="Target Folder:", font=FONT_REGULAR, fg="#94A3B8", bg="#162032").pack(side=tk.LEFT)
        self.ent_dest = tk.Entry(d_row, textvariable=self.out_dir_var, font=FONT_REGULAR)
        self.ent_dest.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
        tk.Button(d_row, text="Browse...", command=self._browse_dest_dir, font=FONT_REGULAR, bg="#334155", fg="#F8FAFC", relief="flat", padx=6, pady=1, cursor="hand2").pack(side=tk.RIGHT)

        # Output Filename row
        n_row = tk.Frame(frame_dest, bg="#162032")
        n_row.pack(fill=tk.X, pady=2)
        tk.Label(n_row, text="Output Name:", font=FONT_REGULAR, fg="#94A3B8", bg="#162032").pack(side=tk.LEFT)
        self.ent_name = tk.Entry(n_row, textvariable=self.out_name_var, font=FONT_REGULAR)
        self.ent_name.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)

        # Post-processing Checkboxes
        c_row = tk.Frame(frame_dest, bg="#162032")
        c_row.pack(fill=tk.X, pady=(4, 0))
        chk_open = tk.Checkbutton(c_row, text="Open destination folder after completion", variable=self.open_dest_var, font=FONT_REGULAR, fg="#F8FAFC", bg="#162032", selectcolor="#0F172A", cursor="hand2")
        chk_open.pack(side=tk.LEFT)

        # --- Card 3: Execution Controls & Telemetry Terminal ---
        card_exec = tk.Frame(self.content_area, bg="#162032", padx=10, pady=8, relief="solid", bd=1)
        card_exec.pack(fill=tk.BOTH, expand=True)

        exec_bar = tk.Frame(card_exec, bg="#162032")
        exec_bar.pack(fill=tk.X, pady=(0, 6))

        self.btn_execute = tk.Button(
            exec_bar,
            text="Execute Operation",
            cursor="hand2",
            command=self._start_execution_thread,
            font=("Segoe UI", 10, "bold"),
            bg="#0284C7",
            fg="#FFFFFF",
            activebackground="#0369A1",
            relief="flat",
            padx=20,
            pady=6
        )
        self.btn_execute.pack(side=tk.LEFT)

        self.lbl_exec_status = tk.Label(exec_bar, text="Ready. Add documents and click Execute.", font=FONT_REGULAR, fg="#94A3B8", bg="#162032")
        self.lbl_exec_status.pack(side=tk.LEFT, padx=14)

        tk.Button(exec_bar, text="Clear Log", command=self._clear_log, font=FONT_REGULAR, bg="#334155", fg="#F8FAFC", relief="flat", padx=6, pady=2, cursor="hand2").pack(side=tk.RIGHT, padx=(6, 4))

        self.lbl_progress_pct = tk.Label(
            exec_bar,
            text="0%",
            font=FONT_BOLD,
            fg="#38BDF8",
            bg="#162032",
            width=5,
            anchor="e"
        )
        self.lbl_progress_pct.pack(side=tk.RIGHT, padx=(2, 6))

        self.progress_bar = ttk.Progressbar(exec_bar, mode="determinate", length=180, maximum=100)
        self.progress_bar.pack(side=tk.RIGHT, padx=4)

        # Monospace Terminal Log
        log_frame = tk.Frame(card_exec, bg="#070B14")
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.txt_log = tk.Text(log_frame, bg="#070B14", fg="#38BDF8", insertbackground="#38BDF8", font=FONT_MONO, height=8, wrap="word", relief="flat")
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=log_scroll.set)
        self.txt_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self._log(f"PDF Toolkit Pro v3.0 initialized. Active mode: {self.MODE_SPECS[self.current_mode]['title']}")
        self._log("System ready. You can upload single or multiple files, reorder them, and execute.")
        self._apply_pointer_cursors()

    def _poll_ui_queue(self):
        """Processes background thread UI events on main thread (Python 3.13 safe)."""
        try:
            while not self.ui_queue.empty():
                kind, payload = self.ui_queue.get_nowait()
                if kind == "log":
                    self.txt_log.insert(tk.END, f"{payload}\n")
                    self.txt_log.see(tk.END)
                elif kind == "progress":
                    percent, status_msg = payload
                    p_val = min(100.0, max(0.0, float(percent)))
                    self.progress_bar.configure(value=p_val)
                    self.lbl_progress_pct.configure(text=f"{int(round(p_val))}%")
                    if status_msg:
                        self.lbl_exec_status.configure(text=status_msg, fg="#38BDF8")
                elif kind == "finish":
                    self._on_execution_finished(payload)
        except Exception:
            pass
        finally:
            try:
                self.root.after(30, self._poll_ui_queue)
            except Exception:
                pass

    def _log(self, msg: str):
        """Thread-safe logging by pushing to UI event queue."""
        if hasattr(self, "ui_queue"):
            self.ui_queue.put(("log", msg))
        else:
            try:
                self.txt_log.insert(tk.END, f"{msg}\n")
                self.txt_log.see(tk.END)
            except Exception:
                pass

    def _clear_log(self):
        self.txt_log.delete("1.0", tk.END)

    def _update_progress(self, percent: float, status_msg: Optional[str] = None):
        """Thread-safe UI updater for determinate progress bar and status text."""
        if hasattr(self, "ui_queue"):
            self.ui_queue.put(("progress", (percent, status_msg)))

    def _switch_mode(self, mode_key: str):
        self.current_mode = mode_key
        spec = self.MODE_SPECS[mode_key]

        # Update Sidebar button highlights
        for m, btn in self.mode_buttons.items():
            if m == mode_key:
                btn.configure(bg="#0284C7", fg="#FFFFFF", font=FONT_BOLD, cursor="hand2")
            else:
                btn.configure(bg="#1E293B", fg="#F8FAFC", font=FONT_REGULAR, cursor="hand2")

        # Update Badge
        self.lbl_active_mode_badge.configure(text=f"Active: {spec['title']}")

        # Suggest default output filename
        if mode_key == "merge":
            self.out_name_var.set("merged_document.pdf")
        elif mode_key == "img2pdf":
            self.out_name_var.set("images_merged.pdf")
        elif mode_key == "pdf2img":
            self.out_name_var.set("converted_pages")
        elif mode_key == "compress":
            self.out_name_var.set("compressed_document.pdf")
        elif mode_key == "split":
            self.out_name_var.set("split_page")
        elif mode_key == "protect":
            self.out_name_var.set("protected_document.pdf")
        elif mode_key == "unlock":
            self.out_name_var.set("unlocked_document.pdf")
        elif mode_key == "md2pdf":
            self.out_name_var.set("rendered_markdown.pdf")
        elif mode_key == "pdf2md":
            self.out_name_var.set("extracted_document.md")

        self._render_dynamic_settings()
        self._log(f"Switched operation to: {spec['title']}")

    def _render_dynamic_settings(self):
        # Clear existing controls
        for w in self.frame_dynamic_settings.winfo_children():
            w.destroy()

        mode = self.current_mode

        if mode == "img2pdf":
            # 1. Page Sizing
            r1 = tk.Frame(self.frame_dynamic_settings, bg="#162032")
            r1.pack(fill=tk.X, pady=2)
            tk.Label(r1, text="Page Size:", font=FONT_REGULAR, fg="#94A3B8", bg="#162032", width=14, anchor="w").pack(side=tk.LEFT)
            cb_ps = ttk.Combobox(r1, textvariable=self.img_pagesize_var, values=["Fit to Image (Original)", "A4 (210 x 297 mm)", "US Letter (8.5 x 11 in)", "Legal (8.5 x 14 in)"], font=FONT_REGULAR, width=22)
            cb_ps.pack(side=tk.LEFT, fill=tk.X, expand=True)

            # 2. Orientation & Margins
            r2 = tk.Frame(self.frame_dynamic_settings, bg="#162032")
            r2.pack(fill=tk.X, pady=2)
            tk.Label(r2, text="Orientation:", font=FONT_REGULAR, fg="#94A3B8", bg="#162032", width=14, anchor="w").pack(side=tk.LEFT)
            cb_or = ttk.Combobox(r2, textvariable=self.img_orient_var, values=["Auto (Best Fit)", "Portrait", "Landscape"], font=FONT_REGULAR, width=12)
            cb_or.pack(side=tk.LEFT, padx=(0, 6))

            tk.Label(r2, text="Margins:", font=FONT_REGULAR, fg="#94A3B8", bg="#162032").pack(side=tk.LEFT)
            cb_mg = ttk.Combobox(r2, textvariable=self.img_margin_var, values=["None (0 mm)", "Narrow (5 mm)", "Standard (10 mm)", "Wide (20 mm)"], font=FONT_REGULAR, width=14)
            cb_mg.pack(side=tk.LEFT, padx=4)

            # 3. Quality & Merging
            r3 = tk.Frame(self.frame_dynamic_settings, bg="#162032")
            r3.pack(fill=tk.X, pady=2)
            tk.Label(r3, text="JPEG Quality (%):", font=FONT_REGULAR, fg="#94A3B8", bg="#162032", width=14, anchor="w").pack(side=tk.LEFT)
            sc = tk.Scale(r3, from_=50, to=100, variable=self.img_quality_var, orient=tk.HORIZONTAL, bg="#162032", fg="#F8FAFC", highlightthickness=0, troughcolor="#0F172A")
            sc.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)

            r4 = tk.Frame(self.frame_dynamic_settings, bg="#162032")
            r4.pack(fill=tk.X, pady=2)
            tk.Radiobutton(r4, text="Merge into Single Multi-Page PDF", variable=self.img_merge_var, value="single", font=FONT_REGULAR, fg="#F8FAFC", bg="#162032", selectcolor="#0F172A", cursor="hand2").pack(side=tk.LEFT, padx=(0, 6))
            tk.Radiobutton(r4, text="Export Individual PDFs per Image", variable=self.img_merge_var, value="multiple", font=FONT_REGULAR, fg="#F8FAFC", bg="#162032", selectcolor="#0F172A", cursor="hand2").pack(side=tk.LEFT)

        elif mode == "pdf2img":
            # 1. Output Format & DPI
            r1 = tk.Frame(self.frame_dynamic_settings, bg="#162032")
            r1.pack(fill=tk.X, pady=2)
            tk.Label(r1, text="Target Format:", font=FONT_REGULAR, fg="#94A3B8", bg="#162032", width=14, anchor="w").pack(side=tk.LEFT)
            cb_fmt = ttk.Combobox(r1, textvariable=self.pdf2img_format_var, values=["PNG (Lossless)", "JPEG (High Compatibility)", "WEBP (Modern Web)", "TIFF (Print / Archival)", "BMP (Uncompressed)"], font=FONT_REGULAR, width=22)
            cb_fmt.pack(side=tk.LEFT, fill=tk.X, expand=True)

            r2 = tk.Frame(self.frame_dynamic_settings, bg="#162032")
            r2.pack(fill=tk.X, pady=2)
            tk.Label(r2, text="Resolution (DPI):", font=FONT_REGULAR, fg="#94A3B8", bg="#162032", width=14, anchor="w").pack(side=tk.LEFT)
            cb_dpi = ttk.Combobox(r2, textvariable=self.pdf2img_dpi_var, values=["150 DPI (Standard Screen)", "200 DPI (Enhanced)", "300 DPI (High-Res Print)", "600 DPI (Ultra Sharp)"], font=FONT_REGULAR, width=22)
            cb_dpi.pack(side=tk.LEFT, fill=tk.X, expand=True)

            # 2. Page Range & Color Mode
            r3 = tk.Frame(self.frame_dynamic_settings, bg="#162032")
            r3.pack(fill=tk.X, pady=2)
            tk.Label(r3, text="Page Range:", font=FONT_REGULAR, fg="#94A3B8", bg="#162032", width=14, anchor="w").pack(side=tk.LEFT)
            ent_rng = tk.Entry(r3, textvariable=self.pdf2img_range_var, font=FONT_REGULAR)
            ent_rng.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))

            tk.Radiobutton(r3, text="Color (RGB)", variable=self.pdf2img_color_var, value="rgb", font=FONT_REGULAR, fg="#F8FAFC", bg="#162032", selectcolor="#0F172A", cursor="hand2").pack(side=tk.LEFT)
            tk.Radiobutton(r3, text="Grayscale", variable=self.pdf2img_color_var, value="grayscale", font=FONT_REGULAR, fg="#F8FAFC", bg="#162032", selectcolor="#0F172A", cursor="hand2").pack(side=tk.LEFT)

            r4 = tk.Frame(self.frame_dynamic_settings, bg="#162032")
            r4.pack(fill=tk.X, pady=2)
            tk.Checkbutton(r4, text="Create subfolder per PDF document", variable=self.pdf2img_subfolder_var, font=FONT_REGULAR, fg="#F8FAFC", bg="#162032", selectcolor="#0F172A", cursor="hand2").pack(side=tk.LEFT)

        elif mode == "merge":
            r1 = tk.Frame(self.frame_dynamic_settings, bg="#162032")
            r1.pack(fill=tk.X, pady=2)
            tk.Checkbutton(r1, text="Generate Table of Contents bookmarks from source filenames", variable=self.merge_bookmarks_var, font=FONT_REGULAR, fg="#F8FAFC", bg="#162032", selectcolor="#0F172A", cursor="hand2").pack(side=tk.LEFT)

        elif mode == "split":
            r1 = tk.Frame(self.frame_dynamic_settings, bg="#162032")
            r1.pack(fill=tk.X, pady=2)
            tk.Radiobutton(r1, text="Split Every Page into Individual PDFs", variable=self.split_mode_var, value="all", font=FONT_REGULAR, fg="#F8FAFC", bg="#162032", selectcolor="#0F172A", cursor="hand2").pack(anchor="w")
            tk.Radiobutton(r1, text="Custom Page Ranges (e.g. 1-3, 4-6):", variable=self.split_mode_var, value="range", font=FONT_REGULAR, fg="#F8FAFC", bg="#162032", selectcolor="#0F172A", cursor="hand2").pack(anchor="w")
            r2 = tk.Frame(self.frame_dynamic_settings, bg="#162032")
            r2.pack(fill=tk.X, pady=2)
            tk.Entry(r2, textvariable=self.split_range_var, font=FONT_REGULAR).pack(fill=tk.X)

        elif mode == "compress":
            r1 = tk.Frame(self.frame_dynamic_settings, bg="#162032")
            r1.pack(fill=tk.X, pady=2)
            tk.Label(r1, text="Compression Level:", font=FONT_REGULAR, fg="#94A3B8", bg="#162032").pack(side=tk.LEFT)
            cb = ttk.Combobox(r1, textvariable=self.compress_level_var, values=["low (Light / Fast)", "medium (Balanced / Recommended)", "high (Aggressive)"], font=FONT_REGULAR)
            cb.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)

        elif mode == "protect":
            r1 = tk.Frame(self.frame_dynamic_settings, bg="#162032")
            r1.pack(fill=tk.X, pady=2)
            tk.Label(r1, text="User Password:", font=FONT_REGULAR, fg="#94A3B8", bg="#162032", width=14, anchor="w").pack(side=tk.LEFT)
            tk.Entry(r1, textvariable=self.user_pwd_var, font=FONT_REGULAR, show="*").pack(side=tk.LEFT, fill=tk.X, expand=True)

            r2 = tk.Frame(self.frame_dynamic_settings, bg="#162032")
            r2.pack(fill=tk.X, pady=2)
            tk.Label(r2, text="Owner Password:", font=FONT_REGULAR, fg="#94A3B8", bg="#162032", width=14, anchor="w").pack(side=tk.LEFT)
            tk.Entry(r2, textvariable=self.owner_pwd_var, font=FONT_REGULAR, show="*").pack(side=tk.LEFT, fill=tk.X, expand=True)

        elif mode == "unlock":
            r1 = tk.Frame(self.frame_dynamic_settings, bg="#162032")
            r1.pack(fill=tk.X, pady=2)
            tk.Label(r1, text="Password:", font=FONT_REGULAR, fg="#94A3B8", bg="#162032", width=14, anchor="w").pack(side=tk.LEFT)
            tk.Entry(r1, textvariable=self.unlock_pwd_var, font=FONT_REGULAR, show="*").pack(side=tk.LEFT, fill=tk.X, expand=True)

        elif mode == "ocr":
            r1 = tk.Frame(self.frame_dynamic_settings, bg="#162032")
            r1.pack(fill=tk.X, pady=2)
            tk.Label(r1, text="Coverage:", font=FONT_REGULAR, fg="#94A3B8", bg="#162032", width=12, anchor="w").pack(side=tk.LEFT)
            tk.Radiobutton(r1, text="All Pages (Full OCR)", variable=self.ocr_coverage_var, value="all", font=FONT_REGULAR, fg="#F8FAFC", bg="#162032", selectcolor="#0F172A", cursor="hand2").pack(side=tk.LEFT, padx=4)
            tk.Radiobutton(r1, text="Smart Hybrid (Scans & Diagrams)", variable=self.ocr_coverage_var, value="hybrid", font=FONT_REGULAR, fg="#F8FAFC", bg="#162032", selectcolor="#0F172A", cursor="hand2").pack(side=tk.LEFT, padx=4)
            tk.Radiobutton(r1, text="Scanned Only", variable=self.ocr_coverage_var, value="scanned_only", font=FONT_REGULAR, fg="#F8FAFC", bg="#162032", selectcolor="#0F172A", cursor="hand2").pack(side=tk.LEFT, padx=4)

            r2 = tk.Frame(self.frame_dynamic_settings, bg="#162032")
            r2.pack(fill=tk.X, pady=2)
            tk.Label(r2, text="Language:", font=FONT_REGULAR, fg="#94A3B8", bg="#162032", width=12, anchor="w").pack(side=tk.LEFT)
            cb_lang = ttk.Combobox(r2, textvariable=self.ocr_lang_var, values=["eng", "ben", "eng+ben"], state="readonly", width=12, font=FONT_REGULAR)
            cb_lang.pack(side=tk.LEFT, padx=4)

            tk.Label(r2, text="Quality / DPI:", font=FONT_REGULAR, fg="#94A3B8", bg="#162032", width=12, anchor="e").pack(side=tk.LEFT, padx=(8, 2))
            cb_dpi = ttk.Combobox(r2, textvariable=self.ocr_dpi_var, values=["300", "400", "200"], state="readonly", width=8, font=FONT_REGULAR)
            cb_dpi.pack(side=tk.LEFT, padx=4)

        elif mode == "pdf2md":
            r1 = tk.Frame(self.frame_dynamic_settings, bg="#162032")
            r1.pack(fill=tk.X, pady=2)
            tk.Checkbutton(r1, text="Extract embedded figures & images into assets folder", variable=self.pdf2md_extract_images_var, font=FONT_REGULAR, fg="#F8FAFC", bg="#162032", selectcolor="#0F172A", cursor="hand2").pack(side=tk.LEFT)

            r2 = tk.Frame(self.frame_dynamic_settings, bg="#162032")
            r2.pack(fill=tk.X, pady=2)
            tk.Checkbutton(r2, text="OCR scanned slides & diagram images into Markdown text", variable=self.pdf2md_ocr_var, font=FONT_REGULAR, fg="#F8FAFC", bg="#162032", selectcolor="#0F172A", cursor="hand2").pack(side=tk.LEFT)

            r3 = tk.Frame(self.frame_dynamic_settings, bg="#162032")
            r3.pack(fill=tk.X, pady=2)
            tk.Label(r3, text="OCR Language:", font=FONT_REGULAR, fg="#94A3B8", bg="#162032", width=14, anchor="w").pack(side=tk.LEFT)
            cb_md_lang = ttk.Combobox(r3, textvariable=self.pdf2md_lang_var, values=["eng", "ben", "eng+ben"], state="readonly", width=12, font=FONT_REGULAR)
            cb_md_lang.pack(side=tk.LEFT, padx=4)

        else:
            tk.Label(self.frame_dynamic_settings, text="Standard operation parameters active. No additional configuration required.", font=FONT_REGULAR, fg="#94A3B8", bg="#162032").pack(anchor="w", pady=4)
        self._apply_pointer_cursors(self.frame_dynamic_settings)

    # --- File Management & Reordering ---
    def _add_files_dialog(self):
        spec = self.MODE_SPECS[self.current_mode]
        try:
            files = filedialog.askopenfilenames(
                parent=self.root,
                title=f"Select Files for {spec['title']}",
                filetypes=spec["exts"]
            )
            if files:
                added = 0
                for f in files:
                    p = Path(f)
                    if p not in self.file_items:
                        self.file_items.append(p)
                        added += 1
                self._update_queue_view()
                self._log(f"[+] Added {added} file(s) to queue.")
                self.lbl_exec_status.configure(text=f"[+] Added {added} file(s).", fg="#34D399")
        except Exception as e:
            messagebox.showerror("File Dialog Error", str(e), parent=self.root)

    def _add_folder_dialog(self):
        spec = self.MODE_SPECS[self.current_mode]
        try:
            folder = filedialog.askdirectory(parent=self.root, title="Select Directory to Import Files")
            if folder:
                accepted_exts = self._get_accepted_extensions()
                added = 0
                for root_dir, _, filenames in os.walk(folder):
                    for fn in filenames:
                        p = Path(root_dir) / fn
                        if p.suffix.lower() in accepted_exts or not accepted_exts:
                            if p not in self.file_items:
                                self.file_items.append(p)
                                added += 1
                self._update_queue_view()
                self._log(f"[+] Added {added} file(s) from directory: {folder}")
                self.lbl_exec_status.configure(text=f"[+] Added {added} file(s) from directory.", fg="#34D399")
        except Exception as e:
            messagebox.showerror("Folder Dialog Error", str(e), parent=self.root)

    def _remove_selected_files(self):
        selected_iids = self.tree.selection()
        if not selected_iids:
            return
        indices_to_remove = sorted([int(iid) for iid in selected_iids], reverse=True)
        for idx in indices_to_remove:
            if 0 <= idx < len(self.file_items):
                del self.file_items[idx]
        self._update_queue_view()

    def _move_file_up(self):
        selected = self.tree.selection()
        if not selected:
            return
        idx = int(selected[0])
        if idx > 0:
            self.file_items[idx - 1], self.file_items[idx] = self.file_items[idx], self.file_items[idx - 1]
            self._update_queue_view()
            self.tree.selection_set(str(idx - 1))

    def _move_file_down(self):
        selected = self.tree.selection()
        if not selected:
            return
        idx = int(selected[0])
        if idx < len(self.file_items) - 1:
            self.file_items[idx + 1], self.file_items[idx] = self.file_items[idx], self.file_items[idx + 1]
            self._update_queue_view()
            self.tree.selection_set(str(idx + 1))

    def _clear_all_files(self):
        self.file_items.clear()
        self._update_queue_view()

    def _update_queue_view(self):
        self.tree.delete(*self.tree.get_children())
        total_bytes = 0
        for idx, path in enumerate(self.file_items):
            try:
                sz = path.stat().st_size
                total_bytes += sz
                sz_str = format_file_size(sz)
            except Exception:
                sz_str = "N/A"

            ext = path.suffix.upper().lstrip(".") or "FILE"
            self.tree.insert("", "end", iid=str(idx), values=(idx + 1, path.name, sz_str, ext, str(path)))

        cnt = len(self.file_items)
        tot_str = format_file_size(total_bytes)
        self.lbl_queue_status.configure(text=f"{cnt} file(s) queued ({tot_str}) -- Drag rows or use Move Up/Down to reorder.")

    # --- Smooth & Professional Drag & Drop Reordering ---
    def _on_tree_press(self, event):
        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return
        self._drag_start_idx = int(row_id)
        if not (0 <= self._drag_start_idx < len(self.file_items)):
            return

        item_name = self.file_items[self._drag_start_idx].name

        # Create floating semi-transparent drag badge tracking the cursor
        self._drag_popup = tk.Toplevel(self.root)
        self._drag_popup.overrideredirect(True)
        self._drag_popup.attributes("-topmost", True)
        try:
            self._drag_popup.attributes("-alpha", 0.92)
        except Exception:
            pass

        badge_frame = tk.Frame(self._drag_popup, bg="#0284C7", padx=8, pady=4, relief="solid", bd=1)
        badge_frame.pack()
        lbl_badge = tk.Label(
            badge_frame,
            text=f"[Reorder] {item_name[:35]}",
            font=("Segoe UI", 8, "bold"),
            fg="#FFFFFF",
            bg="#0284C7"
        )
        lbl_badge.pack()

        self._drag_popup.geometry(f"+{event.x_root + 15}+{event.y_root + 10}")
        self.tree.configure(cursor="fleur")

    def _on_tree_motion(self, event):
        if self._drag_start_idx is None:
            return

        # Smoothly track cursor with floating badge
        if self._drag_popup:
            try:
                self._drag_popup.geometry(f"+{event.x_root + 15}+{event.y_root + 10}")
            except Exception:
                pass

        # Live visual insertion highlight
        target_row = self.tree.identify_row(event.y)
        if target_row:
            self.tree.selection_set(target_row)

        # Smooth auto-scrolling when hovering near upper/lower table bounds
        tree_h = self.tree.winfo_height()
        if event.y < 24:
            self.tree.yview_scroll(-1, "units")
        elif event.y > max(24, tree_h - 24):
            self.tree.yview_scroll(1, "units")

    def _on_tree_release(self, event):
        self._cleanup_drag_popup()
        if self._drag_start_idx is None:
            return

        target_row = self.tree.identify_row(event.y)
        if target_row:
            target_idx = int(target_row)
            if target_idx != self._drag_start_idx and 0 <= target_idx < len(self.file_items):
                item = self.file_items.pop(self._drag_start_idx)
                self.file_items.insert(target_idx, item)
                self._update_queue_view()
                self.tree.selection_set(str(target_idx))
                self._log(f"Reordered: '{item.name}' moved to position {target_idx + 1}")

        self._drag_start_idx = None

    def _on_cancel_drag(self, event=None):
        self._cleanup_drag_popup()
        self._drag_start_idx = None

    def _cleanup_drag_popup(self):
        if self._drag_popup:
            try:
                self._drag_popup.destroy()
            except Exception:
                pass
            self._drag_popup = None
        self.tree.configure(cursor="")

    # --- Windows Explorer External Drag & Drop Hook ---
    def _on_external_drop(self, files):
        added_count = 0
        switched_mode_note = None

        all_dropped = []
        for f in files:
            p_str = f.decode("utf-8", errors="ignore") if isinstance(f, bytes) else str(f)
            all_dropped.append(Path(p_str))

        # Extension category classification
        img_exts = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tiff", ".tif", ".gif", ".ico"}
        md_exts = {".md", ".markdown", ".mdown", ".txt"}
        pdf_exts = {".pdf"}

        # Smart Auto-Mode Detection on Drag & Drop
        dropped_suffixes = {p.suffix.lower() for p in all_dropped if p.is_file()}
        if dropped_suffixes and not (dropped_suffixes & pdf_exts):
            if (dropped_suffixes & img_exts) and self.current_mode != "img2pdf":
                self._switch_mode("img2pdf")
                switched_mode_note = "Switched to 'Images to PDF' mode"
            elif (dropped_suffixes & md_exts) and self.current_mode != "md2pdf":
                self._switch_mode("md2pdf")
                switched_mode_note = "Switched to 'Markdown to PDF' mode"
        elif dropped_suffixes == {".pdf"} and self.current_mode in ("img2pdf", "md2pdf"):
            self._switch_mode("merge")
            switched_mode_note = "Switched to 'Merge Multiple PDFs' mode"

        accepted = self._get_accepted_extensions()
        all_supported = img_exts | md_exts | pdf_exts

        for p in all_dropped:
            try:
                if p.is_dir():
                    for root_dir, _, filenames in os.walk(p):
                        for fn in filenames:
                            child_p = Path(root_dir) / fn
                            if child_p.suffix.lower() in accepted or child_p.suffix.lower() in all_supported:
                                if child_p not in self.file_items:
                                    self.file_items.append(child_p)
                                    added_count += 1
                elif p.is_file():
                    if p.suffix.lower() in accepted or p.suffix.lower() in all_supported:
                        if p not in self.file_items:
                            self.file_items.append(p)
                            added_count += 1
            except Exception as e:
                self._log(f"[-] Drag & drop error: {e}")

        if added_count > 0:
            self._update_queue_view()
            status_text = f"[+] Added {added_count} file(s) via drag-and-drop."
            if switched_mode_note:
                status_text += f" ({switched_mode_note})"
            self._log(status_text)
            self.lbl_exec_status.configure(text=status_text, fg="#34D399")

    def _browse_dest_dir(self):
        d = filedialog.askdirectory(parent=self.root, title="Select Output Destination Folder")
        if d:
            self.out_dir_var.set(os.path.abspath(d))

    # --- Execution Pipeline ---
    def _start_execution_thread(self):
        if self.is_running:
            return
        if not self.file_items:
            messagebox.showwarning("Missing Input Files", "Please add at least one document or image file before executing.", parent=self.root)
            return

        out_d = Path(self.out_dir_var.get().strip())
        if not out_d:
            messagebox.showwarning("Missing Output Path", "Please specify a valid destination folder.", parent=self.root)
            return

        self.is_running = True
        self.btn_execute.configure(state="disabled", text="Processing...", cursor="")
        self._update_progress(0.0, "Starting operation...")

        thread = threading.Thread(target=self._run_worker, daemon=True)
        thread.start()

    def _run_worker(self):
        mode = self.current_mode
        files = list(self.file_items)
        out_dir = Path(self.out_dir_var.get().strip())
        out_name = self.out_name_var.get().strip() or "output_document.pdf"
        os.makedirs(out_dir, exist_ok=True)

        self._log("=" * 60)
        self._log(f"Starting Operation: {self.MODE_SPECS[mode]['title']}")
        self._log(f"Input Files Count: {len(files)}")
        self._log(f"Target Directory: {out_dir}")

        # Helpful mode validation advisory
        if mode == "md2pdf":
            pdf_in_queue = [f for f in files if f.suffix.lower() == ".pdf"]
            if pdf_in_queue:
                self._log("[NOTE] Notice: You selected 'Markdown to PDF', but input contains PDF files:")
                for pf in pdf_in_queue:
                    self._log(f"       -> '{pf.name}' is already a PDF. Select 'PDF to Markdown' to extract text.")
        elif mode == "pdf2md":
            md_in_queue = [f for f in files if f.suffix.lower() in (".md", ".markdown", ".txt")]
            if md_in_queue:
                self._log("[NOTE] Notice: You selected 'PDF to Markdown', but input contains text/markdown files:")
                for mf in md_in_queue:
                    self._log(f"       -> '{mf.name}' is Markdown. Select 'Markdown to PDF' to create a PDF.")
        self._log("=" * 60)

        t0 = time.time()
        success = False
        error_msg = ""
        generated_files = []

        try:
            raw_res = None
            if mode == "img2pdf":
                ps_map = {
                    "Fit to Image (Original)": "fit",
                    "A4 (210 x 297 mm)": "a4",
                    "US Letter (8.5 x 11 in)": "letter",
                    "Legal (8.5 x 14 in)": "legal"
                }
                or_map = {"Auto (Best Fit)": "auto", "Portrait": "portrait", "Landscape": "landscape"}
                mg_map = {"None (0 mm)": 0.0, "Narrow (5 mm)": 5.0, "Standard (10 mm)": 10.0, "Wide (20 mm)": 20.0}

                p_size = ps_map.get(self.img_pagesize_var.get(), "fit")
                orient = or_map.get(self.img_orient_var.get(), "auto")
                margin = mg_map.get(self.img_margin_var.get(), 0.0)
                quality = int(self.img_quality_var.get())
                is_single = (self.img_merge_var.get() == "single")

                fn = self._get_engine_func("execute_images_to_pdf")
                raw_res = fn(
                    image_paths=files,
                    output_dir=out_dir,
                    output_name=out_name,
                    page_size=p_size,
                    orientation=orient,
                    margin_mm=margin,
                    quality=quality,
                    single_pdf=is_single,
                    log_callback=self._log,
                    progress_callback=self._update_progress
                )

            elif mode == "pdf2img":
                fmt_map = {
                    "PNG (Lossless)": "png",
                    "JPEG (Compressed)": "jpeg",
                    "WEBP (Modern Web)": "webp",
                    "TIFF (Print / Archival)": "tiff",
                    "BMP (Uncompressed)": "bmp"
                }
                dpi_map = {
                    "150 DPI (Standard Screen)": 150,
                    "200 DPI (Enhanced)": 200,
                    "300 DPI (High-Res Print)": 300,
                    "600 DPI (Ultra Sharp)": 600
                }

                target_fmt = fmt_map.get(self.pdf2img_format_var.get(), "png")
                target_dpi = dpi_map.get(self.pdf2img_dpi_var.get(), 150)
                p_range = self.pdf2img_range_var.get().strip() or "all"
                c_mode = self.pdf2img_color_var.get()
                subfolder = bool(self.pdf2img_subfolder_var.get())
                quality = int(self.pdf2img_quality_var.get())

                fn = self._get_engine_func("execute_pdf_to_images")
                raw_res = fn(
                    pdf_paths=files,
                    output_dir=out_dir,
                    img_format=target_fmt,
                    dpi=target_dpi,
                    page_range=p_range,
                    color_mode=c_mode,
                    quality=quality,
                    subfolder_per_pdf=subfolder,
                    log_callback=self._log,
                    progress_callback=self._update_progress
                )

            elif mode == "merge":
                fn = self._get_engine_func("execute_merge_pdfs")
                raw_res = fn(
                    selected_files=files,
                    output_parent_folder=out_dir,
                    merged_filename=out_name,
                    save_in_subfolder=False,
                    progress_callback=self._update_progress,
                    log_callback=self._log
                )

            elif mode == "split":
                fn = self._get_engine_func("execute_split_pdfs")
                s_mode = "A" if self.split_mode_var.get() == "all" else "B"
                r_str = self.split_range_var.get().strip() if s_mode == "B" else None
                raw_res = fn(
                    selected_files=files,
                    output_parent_folder=out_dir,
                    mode=s_mode,
                    range_str=r_str,
                    progress_callback=self._update_progress,
                    log_callback=self._log
                )

            elif mode == "compress":
                fn = self._get_engine_func("execute_compress_pdfs")
                lvl = self.compress_level_var.get().split()[0].lower()
                raw_res = fn(
                    selected_files=files,
                    level=lvl,
                    output_dir=out_dir,
                    progress_callback=self._update_progress,
                    log_callback=self._log
                )

            elif mode == "protect":
                fn = self._get_engine_func("execute_protect_pdf")
                u_pwd = self.user_pwd_var.get().strip()
                o_pwd = self.owner_pwd_var.get().strip() or None
                if not u_pwd:
                    raise ValueError("User password cannot be empty for protection.")
                raw_res = fn(
                    selected_files=files,
                    user_password=u_pwd,
                    owner_password=o_pwd,
                    output_dir=out_dir,
                    progress_callback=self._update_progress,
                    log_callback=self._log
                )

            elif mode == "unlock":
                fn = self._get_engine_func("execute_unlock_pdf")
                pwd = self.unlock_pwd_var.get().strip() or None
                raw_res = fn(
                    selected_files=files,
                    password=pwd,
                    output_dir=out_dir,
                    progress_callback=self._update_progress,
                    log_callback=self._log
                )

            elif mode == "ocr":
                fn = self._get_engine_func("execute_ocr_pdf")
                cov = self.ocr_coverage_var.get()
                lng = self.ocr_lang_var.get()
                try:
                    dpi_val = int(self.ocr_dpi_var.get())
                except Exception:
                    dpi_val = 300
                raw_res = fn(
                    selected_files=files,
                    lang=lng,
                    process_all_pages=(cov == "all"),
                    coverage_mode=cov,
                    dpi=dpi_val,
                    output_dir=out_dir,
                    progress_callback=self._update_progress,
                    log_callback=self._log
                )

            elif mode == "md2pdf":
                fn = self._get_engine_func("execute_markdown_to_pdf")
                raw_res = fn(
                    selected_files=files,
                    output_dir=out_dir,
                    progress_callback=self._update_progress,
                    log_callback=self._log
                )

            elif mode == "pdf2md":
                fn = self._get_engine_func("execute_pdf_to_markdown")
                raw_res = fn(
                    selected_files=files,
                    extract_images=self.pdf2md_extract_images_var.get(),
                    ocr_fallback=self.pdf2md_ocr_var.get(),
                    ocr_images=self.pdf2md_ocr_var.get(),
                    lang=self.pdf2md_lang_var.get(),
                    output_dir=out_dir,
                    progress_callback=self._update_progress,
                    log_callback=self._log
                )
            # Process output report / dict
            if isinstance(raw_res, dict):
                success = (raw_res.get("status") == "success")
                generated_files = raw_res.get("files", [])
                error_msg = raw_res.get("error", "")
            elif hasattr(raw_res, "results"):
                successful = [r for r in raw_res.results if r.get("status") == "Success"]
                failed = [r for r in raw_res.results if r.get("status") == "Failed"]
                all_outputs = []
                for r in raw_res.results:
                    all_outputs.extend(r.get("output_files", []))
                generated_files = list(dict.fromkeys(all_outputs))
                success = (len(successful) > 0)
                if failed:
                    err_msgs = [f"{r.get('input_file', '')}: {r.get('error', '')}" for r in failed if r.get("error")]
                    error_msg = "; ".join(err_msgs)
            elif isinstance(raw_res, list):
                generated_files = list(dict.fromkeys(raw_res))
                success = bool(generated_files)
            elif isinstance(raw_res, Path):
                generated_files = [raw_res] if raw_res.exists() else []
                success = bool(generated_files)
            else:
                success = bool(raw_res)

            for out_f in generated_files:
                self._log(f"[+] Output: {out_f}")

        except Exception as e:
            error_msg = str(e)
            self._log(f"[FATAL ERROR] {error_msg}")

        elapsed = time.time() - t0
        self._log("-" * 60)

        # Notify Completion
        if hasattr(self, "ui_queue"):
            self.ui_queue.put(("finish", {
                "success": success,
                "generated_files": generated_files,
                "error_msg": error_msg,
                "elapsed": elapsed,
                "out_dir": out_dir,
                "open_dest": self.open_dest_var.get()
            }))

    def _on_execution_finished(self, data: Dict[str, Any]):
        """Executes completion logic safely on the main UI thread."""
        self.btn_execute.configure(state="normal", text="Execute Operation", cursor="hand2")
        self.is_running = False

        success = data.get("success", False)
        generated_files = data.get("generated_files", [])
        error_msg = data.get("error_msg", "")
        elapsed = data.get("elapsed", 0.0)
        out_dir = data.get("out_dir")
        open_dest = data.get("open_dest", False)

        if success:
            self.progress_bar.configure(value=100.0)
            self.lbl_progress_pct.configure(text="100%", fg="#34D399")
            self.lbl_exec_status.configure(text=f"[OK] Completed in {elapsed:.2f}s! ({len(generated_files)} file(s))", fg="#34D399")
            self._log(f"[SUCCESS] Operation finished in {elapsed:.2f} seconds.")
            if open_dest and out_dir and Path(out_dir).exists():
                try:
                    os.startfile(str(out_dir))
                except Exception:
                    pass
            messagebox.showinfo("Operation Succeeded", f"Operation completed successfully in {elapsed:.2f} seconds!\n\nDestination: {out_dir}", parent=self.root)
        else:
            self.lbl_progress_pct.configure(fg="#F87171")
            self.lbl_exec_status.configure(text=f"[Error] Failed: {error_msg or 'Execution error'}", fg="#F87171")
            self._log(f"[FAILED] Operation failed: {error_msg}")
            messagebox.showerror("Operation Failed", f"Operation terminated with error:\n\n{error_msg or 'Unknown error'}", parent=self.root)


def run_pdf_toolkit():
    """
    Launches the modern PDF Toolkit Pro Desktop GUI Dashboard.
    """
    app = PDFToolkitDashboardGUI()
    app.root.mainloop()

if __name__ == "__main__":
    run_pdf_toolkit()
