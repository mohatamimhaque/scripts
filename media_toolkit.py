"""
==============================================================================
Universal Media Converter Pro -- Audio & Video Processing Studio
Developed by: mohatamim
==============================================================================
Unified enterprise audio and video processing suite combining:
  1. Audio to Audio Converter (21+ audio formats, bitrates, sample rates)
  2. Video to Audio Converter & Extractor (19+ video formats to lossless/lossy audio)
  3. Video to Video Converter & Transcoder (Codecs: H.264, H.265, VP9, MPEG-4; CRF & presets)

Features:
  - Glassmorphism dark-themed Desktop GUI Dashboard
  - Smooth drag-and-drop file reordering with floating badge
  - Windows Explorer external drag & drop support (windnd)
  - Comprehensive queue sorting (Name A-Z/Z-A, Size, Extension, Duration)
  - Hand pointer cursor on all active buttons (hand2) with disabled state safety
  - Real-time FFmpeg progress telemetry terminal with duration calculation
==============================================================================
"""

import os
import sys
import time
import shutil
import logging
import threading
import queue
import subprocess
import re
from pathlib import Path
from typing import List, Dict, Tuple, Optional, Any, Set, Callable

# Ensure local directories and parent paths are in sys.path
for _p in [Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd(), Path.cwd(), Path.cwd().parent, Path(r"d:\New folder\PY"), Path(r"d:\New folder\PY\scripts")]:
    _p_str = str(_p.resolve()) if hasattr(_p, "resolve") and _p.exists() else str(_p)
    if _p_str not in sys.path:
        sys.path.insert(0, _p_str)

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk, ImageDraw

try:
    import windnd
    WINDND_AVAILABLE = True
except ImportError:
    WINDND_AVAILABLE = False

# Setup Logger
logger = logging.getLogger("MediaToolkitPro")
if not logger.handlers:
    _handler = logging.StreamHandler(sys.stdout)
    _handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    logger.addHandler(_handler)
logger.setLevel(logging.INFO)


# ==============================================================================
# SECTION 1: ENVIRONMENT & FFMPEG DETECTION
# ==============================================================================

def detect_ffmpeg() -> Optional[str]:
    """
    Locates the FFmpeg executable across system PATH, local virtual environments,
    imageio_ffmpeg bundled binaries, and standard Windows installation paths.
    """
    exe = shutil.which("ffmpeg")
    if exe:
        return exe

    # Virtual environment paths
    candidate_venvs = [
        Path(r"d:\New folder\PY\scripts\.venv\Scripts\ffmpeg.exe"),
        Path(r"d:\New folder\PY\.venv\Scripts\ffmpeg.exe")
    ]
    for vp in candidate_venvs:
        if vp.is_file():
            os.environ["PATH"] = str(vp.parent) + os.pathsep + os.environ.get("PATH", "")
            return str(vp)

    # imageio_ffmpeg bundled binary
    try:
        import imageio_ffmpeg
        loc = imageio_ffmpeg.get_ffmpeg_exe()
        if loc and os.path.isfile(loc):
            os.environ["PATH"] = str(Path(loc).parent) + os.pathsep + os.environ.get("PATH", "")
            return loc
    except Exception:
        pass

    # Common Windows installations
    common_locations = [
        r"C:\ffmpeg\bin\ffmpeg.exe",
        r"C:\ffmpeg\ffmpeg.exe",
        r"C:\Program Files\ffmpeg\bin\ffmpeg.exe",
        r"C:\ProgramData\chocolatey\bin\ffmpeg.exe",
        os.path.expanduser(r"~\AppData\Local\Microsoft\WinGet\Links\ffmpeg.exe"),
        os.path.expanduser(r"~\AppData\Local\ffmpeg\bin\ffmpeg.exe"),
    ]
    for loc in common_locations:
        if os.path.isfile(loc):
            os.environ["PATH"] = str(Path(loc).parent) + os.pathsep + os.environ.get("PATH", "")
            return loc

    return None

FFMPEG_PATH = detect_ffmpeg()


# ==============================================================================
# SECTION 2: FORMAT SPECIFICATIONS & CODEC MATRIX
# ==============================================================================

AUDIO_FORMATS = [
    ".mp3", ".wav", ".aac", ".m4a", ".flac", ".ogg", ".opus", ".wma",
    ".aiff", ".aif", ".alac", ".amr", ".ac3", ".eac3", ".ape", ".caf", ".au"
]

VIDEO_FORMATS = [
    ".mp4", ".mkv", ".mov", ".avi", ".wmv", ".webm", ".flv", ".mpeg",
    ".mpg", ".m4v", ".3gp", ".ts", ".mts", ".m2ts", ".vob", ".ogv", ".asf", ".divx", ".f4v"
]

VIDEO_CODECS_MAP = {
    ".mp4": ("libx264", "aac"),
    ".mkv": ("libx264", "aac"),
    ".webm": ("libvpx-vp9", "libopus"),
    ".mov": ("libx264", "aac"),
    ".avi": ("mpeg4", "mp3"),
    ".wmv": ("wmv2", "wmav2"),
    ".flv": ("flv1", "mp3"),
    ".f4v": ("libx264", "aac"),
    ".mpeg": ("mpeg2video", "mp2"),
    ".mpg": ("mpeg2video", "mp2"),
    ".vob": ("mpeg2video", "ac3"),
    ".m4v": ("libx264", "aac"),
    ".3gp": ("libx264", "aac"),
    ".ts": ("libx264", "aac"),
    ".mts": ("libx264", "aac"),
    ".m2ts": ("libx264", "aac"),
    ".ogv": ("libtheora", "libvorbis"),
    ".asf": ("wmv2", "wmav2"),
    ".divx": ("mpeg4", "mp3")
}


def format_file_size(num_bytes: int) -> str:
    """Returns human-readable file size string."""
    if num_bytes < 1024:
        return f"{num_bytes} B"
    for unit in ["KB", "MB", "GB", "TB"]:
        num_bytes /= 1024.0
        if num_bytes < 1024.0:
            return f"{num_bytes:.1f} {unit}"
    return f"{num_bytes:.1f} PB"


def format_duration(seconds: Optional[float]) -> str:
    """Formats seconds into HH:MM:SS or MM:SS."""
    if seconds is None or seconds < 0:
        return "N/A"
    hrs = int(seconds // 3600)
    mins = int((seconds % 3600) // 60)
    secs = int(seconds % 60)
    if hrs > 0:
        return f"{hrs:02d}:{mins:02d}:{secs:02d}"
    return f"{mins:02d}:{secs:02d}"


def get_media_duration(file_path: Path, ffmpeg_bin: Optional[str] = None) -> Optional[float]:
    """Extracts duration of audio or video file via FFmpeg."""
    bin_path = ffmpeg_bin or FFMPEG_PATH
    if not bin_path or not file_path.exists():
        return None
    try:
        proc = subprocess.run(
            [bin_path, "-hide_banner", "-i", str(file_path)],
            stderr=subprocess.PIPE,
            stdout=subprocess.DEVNULL,
            text=True,
            encoding="utf-8",
            errors="replace"
        )
        m = re.search(r"Duration:\s*(\d+):(\d+):(\d+(?:\.\d+)?)", proc.stderr)
        if m:
            h, mins, s = float(m.group(1)), float(m.group(2)), float(m.group(3))
            return h * 3600 + mins * 60 + s
    except Exception:
        pass
    return None


# ==============================================================================
# SECTION 3: FFMPEG COMMAND BUILDERS & TRANSCODING ENGINES
# ==============================================================================

def build_audio_ffmpeg_args(
    target_ext: str,
    bitrate: str = "192k",
    sample_rate: str = "original",
    channels: str = "original",
    normalize: bool = False
) -> List[str]:
    """Constructs FFmpeg audio encoding arguments based on format and options."""
    ext = target_ext.lower()
    args: List[str] = []

    if ext == ".mp3":
        args.extend(["-c:a", "libmp3lame", "-b:a", bitrate])
    elif ext == ".wav":
        args.extend(["-c:a", "pcm_s16le"])
    elif ext in (".aac", ".m4a"):
        args.extend(["-c:a", "aac", "-b:a", bitrate])
    elif ext == ".flac":
        args.extend(["-c:a", "flac"])
    elif ext == ".ogg":
        args.extend(["-c:a", "libvorbis", "-b:a", bitrate])
    elif ext == ".opus":
        opus_br = "128k" if bitrate in ("192k", "128k") else ("160k" if bitrate == "256k" else "192k")
        args.extend(["-c:a", "libopus", "-b:a", opus_br])
    elif ext == ".wma":
        args.extend(["-c:a", "wmav2", "-b:a", "160k"])
    elif ext in (".aiff", ".aif"):
        args.extend(["-c:a", "pcm_s16be"])
    elif ext == ".alac":
        args.extend(["-c:a", "alac"])
    elif ext == ".amr":
        args.extend(["-c:a", "amr_nb", "-ar", "8000", "-ac", "1"])
    elif ext == ".ac3":
        args.extend(["-c:a", "ac3", "-b:a", "384k"])
    elif ext == ".eac3":
        args.extend(["-c:a", "eac3", "-b:a", "448k"])
    elif ext == ".ape":
        args.extend(["-c:a", "ape"])
    elif ext == ".caf":
        args.extend(["-c:a", "pcm_s16le"])
    elif ext == ".au":
        args.extend(["-c:a", "pcm_s16be"])
    else:
        args.extend(["-b:a", bitrate])

    if sample_rate != "original" and ext != ".amr":
        args.extend(["-ar", str(sample_rate)])

    if channels == "mono" and ext != ".amr":
        args.extend(["-ac", "1"])
    elif channels == "stereo" and ext != ".amr":
        args.extend(["-ac", "2"])

    if normalize:
        args.extend(["-filter:a", "loudnorm"])

    return args


def execute_media_transcode(
    src_file: Path,
    dst_file: Path,
    cmd_args: List[str],
    ffmpeg_bin: Optional[str] = None,
    log_callback: Optional[Callable[[str], None]] = None,
    progress_callback: Optional[Callable[[float, str], None]] = None
) -> Dict[str, Any]:
    """
    Executes an FFmpeg conversion process with reliable real-time progress parsing.
    """
    bin_path = ffmpeg_bin or detect_ffmpeg() or FFMPEG_PATH
    if not bin_path:
        raise RuntimeError("FFmpeg executable not detected. Please install FFmpeg or imageio-ffmpeg.")

    full_cmd = [bin_path, "-y", "-i", str(src_file)] + cmd_args + ["-progress", "pipe:1", str(dst_file)]
    duration_sec = get_media_duration(src_file, bin_path)
    start_time = time.time()
    last_reported = -1
    log_tail: List[str] = []

    if progress_callback:
        progress_callback(0.0, "Starting transcode...")

    proc = subprocess.Popen(
        full_cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        universal_newlines=True,
        bufsize=1
    )

    for line in proc.stdout:
        line_str = line.strip()
        if not line_str:
            continue
        log_tail.append(line_str)
        if len(log_tail) > 100:
            log_tail.pop(0)

        # Fallback duration discovery from stream info
        if (duration_sec is None or duration_sec <= 0) and "Duration:" in line_str:
            m_dur = re.search(r"Duration:\s*(\d+):(\d+):(\d+(?:\.\d+)?)", line_str)
            if m_dur:
                duration_sec = float(m_dur.group(1))*3600 + float(m_dur.group(2))*60 + float(m_dur.group(3))

        curr_sec = None
        if line_str.startswith("out_time_ms=") or line_str.startswith("out_time_us="):
            val_str = line_str.split("=", 1)[1].strip()
            if val_str and val_str != "N/A":
                try:
                    curr_sec = float(val_str) / 1_000_000.0
                except Exception:
                    pass
        elif line_str.startswith("out_time="):
            val_str = line_str.split("=", 1)[1].strip()
            if val_str and val_str != "N/A":
                m_t = re.match(r"(\d+):(\d+):(\d+(?:\.\d+)?)", val_str)
                if m_t:
                    curr_sec = float(m_t.group(1))*3600 + float(m_t.group(2))*60 + float(m_t.group(3))
        elif "time=" in line_str:
            m_t = re.search(r"time=(\d+):(\d+):(\d+(?:\.\d+)?)", line_str)
            if m_t:
                curr_sec = float(m_t.group(1))*3600 + float(m_t.group(2))*60 + float(m_t.group(3))

        if curr_sec is not None:
            if duration_sec and duration_sec > 0:
                pct = min(99.0, max(0.0, (curr_sec / duration_sec) * 100.0))
                whole = int(pct)
                if whole > last_reported:
                    last_reported = whole
                    msg = f"Progress: {pct:5.1f}% [{curr_sec:.1f}s / {duration_sec:.1f}s]"
                    if whole % 10 == 0 and log_callback:
                        log_callback(f"    {msg}")
                    if progress_callback:
                        progress_callback(pct, msg)
            else:
                msg = f"Transcoded: {curr_sec:.1f}s processed"
                if progress_callback:
                    progress_callback(50.0, msg)
        elif line_str == "progress=end":
            if progress_callback:
                progress_callback(100.0, "Progress: 100.0% [Completed]")

    ret = proc.wait()
    if ret != 0:
        tail_msg = "\n".join(log_tail[-20:])
        raise RuntimeError(f"FFmpeg process exited with code {ret}:\n{tail_msg}")

    elapsed = time.time() - start_time
    out_size = dst_file.stat().st_size if dst_file.exists() else 0
    in_size = src_file.stat().st_size if src_file.exists() else 0

    return {
        "input": str(src_file),
        "output": str(dst_file),
        "status": "Success",
        "elapsed_seconds": round(elapsed, 2),
        "input_size_bytes": in_size,
        "output_size_bytes": out_size
    }


# ==============================================================================
# SECTION 4: MODERN DESKTOP GUI DASHBOARD (TKINTER)
# ==============================================================================

FONT_FAMILY = "Segoe UI"
FONT_TITLE = (FONT_FAMILY, 11, "bold")
FONT_HEADING = (FONT_FAMILY, 9, "bold")
FONT_REGULAR = (FONT_FAMILY, 8)
FONT_BOLD = (FONT_FAMILY, 8, "bold")
FONT_MONO = ("Consolas", 8)


class MaterialIconManager:
    """Renders vector icons onto PIL images for dark-themed UI buttons."""
    _cache: Dict[Tuple[str, Tuple[int, int], str], ImageTk.PhotoImage] = {}

    @classmethod
    def get(cls, name: str, size: Tuple[int, int] = (16, 16), color_hex: str = "#38BDF8") -> ImageTk.PhotoImage:
        key = (name, size, color_hex)
        if key in cls._cache:
            return cls._cache[key]

        w, h = size
        img = Image.new("RGBA", (w, h), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        # Vector drawing logic
        if name == "audio":
            draw.ellipse([2, h - 6, 6, h - 2], fill=color_hex)
            draw.ellipse([w - 7, h - 8, w - 3, h - 4], fill=color_hex)
            draw.line([(5, h - 4), (5, 3), (w - 4, 1), (w - 4, h - 6)], fill=color_hex, width=2)
            draw.line([(5, 4), (w - 4, 2)], fill=color_hex, width=3)
        elif name == "video":
            draw.rectangle([2, 3, w - 6, h - 3], outline=color_hex, width=2)
            draw.polygon([(w - 5, h // 2 - 3), (w - 1, 3), (w - 1, h - 3)], fill=color_hex)
        elif name == "extract":
            draw.rectangle([2, 2, w - 8, h - 2], outline=color_hex, width=2)
            draw.polygon([(w - 6, 4), (w - 1, h // 2), (w - 6, h - 4)], fill=color_hex)
            draw.line([(w // 2 - 2, h // 2), (w - 4, h // 2)], fill=color_hex, width=2)
        elif name == "sort":
            draw.line([(4, 3), (4, h - 3)], fill=color_hex, width=2)
            draw.polygon([(2, 6), (4, 2), (6, 6)], fill=color_hex)
            draw.line([(w - 4, 3), (w - 4, h - 3)], fill=color_hex, width=2)
            draw.polygon([(w - 6, h - 6), (w - 4, h - 2), (w - 2, h - 6)], fill=color_hex)
        else:
            draw.rectangle([2, 2, w - 2, h - 2], outline=color_hex, width=2)

        photo = ImageTk.PhotoImage(img)
        cls._cache[key] = photo
        return photo


class MediaToolkitDashboardGUI:
    """
    State-of-the-art Glassmorphism Desktop GUI for Universal Audio & Video Conversion.
    """

    MODE_SPECS = {
        "audio2audio": {
            "title": "Audio to Audio",
            "desc": "Transcode between 21+ audio formats with bitrate and channel tuning",
            "icon": "audio",
            "exts": [
                ("All Supported Audio Files", " ".join(f"*{e}" for e in AUDIO_FORMATS)),
                ("MP3 Audio (*.mp3)", "*.mp3"),
                ("WAV Audio (*.wav)", "*.wav"),
                ("FLAC Audio (*.flac)", "*.flac"),
                ("AAC / M4A Audio (*.aac *.m4a)", "*.aac *.m4a"),
                ("All Files (*.*)", "*.*")
            ],
            "accepted_inputs": set(AUDIO_FORMATS)
        },
        "video2audio": {
            "title": "Video to Audio",
            "desc": "Extract audio tracks from video into lossless or compressed audio",
            "icon": "extract",
            "exts": [
                ("All Supported Video Files", " ".join(f"*{e}" for e in VIDEO_FORMATS)),
                ("MP4 Video (*.mp4)", "*.mp4"),
                ("MKV Video (*.mkv)", "*.mkv"),
                ("MOV / WebM Video (*.mov *.webm)", "*.mov *.webm"),
                ("All Files (*.*)", "*.*")
            ],
            "accepted_inputs": set(VIDEO_FORMATS)
        },
        "video2video": {
            "title": "Video to Video",
            "desc": "Transcode, compress and re-encode video containers and codecs",
            "icon": "video",
            "exts": [
                ("All Supported Video Files", " ".join(f"*{e}" for e in VIDEO_FORMATS)),
                ("MP4 Video (*.mp4)", "*.mp4"),
                ("MKV Video (*.mkv)", "*.mkv"),
                ("MOV / WebM Video (*.mov *.webm)", "*.mov *.webm"),
                ("All Files (*.*)", "*.*")
            ],
            "accepted_inputs": set(VIDEO_FORMATS)
        }
    }

    def __init__(self, parent=None):
        self.root = tk.Tk() if parent is None else tk.Toplevel(parent)
        self.root.title("Universal Media Converter Pro -- Audio & Video Processing Suite (mohatamim)")
        self.root.geometry("1240x820")
        self.root.minsize(1100, 720)
        self.root.configure(bg="#0B1120")

        self.current_mode = "audio2audio"
        self.file_items: List[Path] = []
        self.is_running = False

        # Output options
        default_out = str(Path.home() / "Documents" / "Media_Toolkit_Output")
        self.out_dir_var = tk.StringVar(value=default_out)
        self.out_name_var = tk.StringVar(value="")
        self.open_dest_var = tk.BooleanVar(value=True)

        # Mode-specific variables
        # Audio to Audio
        self.a2a_format_var = tk.StringVar(value=".mp3")
        self.a2a_bitrate_var = tk.StringVar(value="192k")
        self.a2a_samplerate_var = tk.StringVar(value="original")
        self.a2a_channels_var = tk.StringVar(value="original")
        self.a2a_normalize_var = tk.BooleanVar(value=False)

        # Video to Audio
        self.v2a_format_var = tk.StringVar(value=".mp3")
        self.v2a_bitrate_var = tk.StringVar(value="192k")
        self.v2a_samplerate_var = tk.StringVar(value="original")
        self.v2a_channels_var = tk.StringVar(value="original")

        # Video to Video
        self.v2v_format_var = tk.StringVar(value=".mp4")
        self.v2v_codec_var = tk.StringVar(value="libx264 (H.264 / AVC)")
        self.v2v_quality_var = tk.StringVar(value="Balanced (CRF 23 - Recommended)")
        self.v2v_preset_var = tk.StringVar(value="medium (Standard Speed)")
        self.v2v_resolution_var = tk.StringVar(value="Original (Keep Resolution)")
        self.v2v_audio_var = tk.StringVar(value="Re-encode (AAC 192k)")
        self.v2v_stream_copy_var = tk.BooleanVar(value=False)

        # Drag-and-drop state
        self._drag_start_idx = None
        self._drag_popup = None

        self._build_ui()

        # Windows Explorer drag-and-drop hook
        if WINDND_AVAILABLE:
            try:
                windnd.hook_dropfiles(self.root, func=self._on_external_drop)
            except Exception:
                pass

        self.root.bind("<Escape>", self._on_cancel_drag)
        self._bind_keyboard_shortcuts()
        self._switch_mode("audio2audio")
        self._apply_pointer_cursors()

        # Thread-safe UI event queue for Python 3.13 compatibility
        self.ui_queue = queue.Queue()
        self.root.after(30, self._poll_ui_queue)

    def _show_shortcuts_modal(self):
        """Displays interactive modal detailing all keyboard shortcuts."""
        dlg = tk.Toplevel(self.root)
        dlg.title("Keyboard Shortcuts Reference")
        dlg.geometry("540x500")
        dlg.minsize(480, 420)
        dlg.configure(bg="#0F172A")
        dlg.transient(self.root)
        dlg.grab_set()

        hdr = tk.Frame(dlg, bg="#1E293B", padx=16, pady=12)
        hdr.pack(fill=tk.X)
        tk.Label(hdr, text="Universal Media Converter -- Keyboard Shortcuts", font=("Segoe UI", 12, "bold"), fg="#38BDF8", bg="#1E293B").pack(anchor="w")
        tk.Label(hdr, text="Accelerate your media transcoding with rapid keyboard navigation", font=("Segoe UI", 9), fg="#94A3B8", bg="#1E293B").pack(anchor="w")

        body = tk.Frame(dlg, bg="#0F172A", padx=16, pady=12)
        body.pack(fill=tk.BOTH, expand=True)

        shortcuts = [
            ("Ctrl + O / Ctrl + A", "Add single or multiple media file(s) to queue"),
            ("Ctrl + Shift + O / Ctrl + F", "Import all matching media files from a folder"),
            ("Ctrl + Enter / F5", "Start transcoding / conversion process"),
            ("Delete / Backspace", "Remove highlighted file(s) from queue"),
            ("Ctrl + Delete / Ctrl + L", "Clear entire media queue"),
            ("Ctrl + Up / Alt + Up", "Move highlighted file UP in queue"),
            ("Ctrl + Down / Alt + Down", "Move highlighted file DOWN in queue"),
            ("Alt + 1 / 2 / 3", "Switch mode (1: Audio2Audio, 2: Video2Audio, 3: Video2Video)"),
            ("Escape", "Cancel drag / Close modal dialogs"),
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
        """Binds global keyboard shortcuts for rapid media conversion."""
        self.root.bind("<Control-o>", lambda e: self._add_files_dialog())
        self.root.bind("<Control-O>", lambda e: self._add_files_dialog())
        self.root.bind("<Control-a>", lambda e: self._add_files_dialog())
        self.root.bind("<Control-A>", lambda e: self._add_files_dialog())
        self.root.bind("<Control-Shift-O>", lambda e: self._add_folder_dialog())
        self.root.bind("<Control-Shift-o>", lambda e: self._add_folder_dialog())
        self.root.bind("<Control-f>", lambda e: self._add_folder_dialog())
        self.root.bind("<Control-F>", lambda e: self._add_folder_dialog())

        self.root.bind("<Control-Return>", lambda e: self._start_execution_thread())
        self.root.bind("<Control-KP_Enter>", lambda e: self._start_execution_thread())
        self.root.bind("<F5>", lambda e: self._start_execution_thread())

        self.root.bind("<Delete>", lambda e: self._remove_selected_files())
        self.root.bind("<BackSpace>", lambda e: self._remove_selected_files())
        self.root.bind("<Control-Delete>", lambda e: self._clear_all_files())
        self.root.bind("<Control-l>", lambda e: self._clear_all_files())
        self.root.bind("<Control-L>", lambda e: self._clear_all_files())
        self.root.bind("<Control-Up>", lambda e: self._move_file_up())
        self.root.bind("<Alt-Up>", lambda e: self._move_file_up())
        self.root.bind("<Control-Down>", lambda e: self._move_file_down())
        self.root.bind("<Alt-Down>", lambda e: self._move_file_down())

        self.root.bind("<F1>", lambda e: self._show_shortcuts_modal())
        self.root.bind("<Control-h>", lambda e: self._show_shortcuts_modal())
        self.root.bind("<Control-H>", lambda e: self._show_shortcuts_modal())

        # Alt+1, Alt+2, Alt+3 for mode switching
        self.root.bind("<Alt-Key-1>", lambda e: self._switch_mode("audio2audio"))
        self.root.bind("<Alt-Key-2>", lambda e: self._switch_mode("video2audio"))
        self.root.bind("<Alt-Key-3>", lambda e: self._switch_mode("video2video"))

    def _apply_pointer_cursors(self, parent=None):
        """Recursively binds pointing hand cursor on all active clickable widgets."""
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
        tk.Label(h_left, text="Universal Media Converter Pro | Audio & Video Processing Studio", font=FONT_TITLE, fg="#38BDF8", bg="#162032").pack(anchor="w")
        tk.Label(h_left, text="Engine Architecture by mohatamim -- High-Performance FFmpeg Media Engine", font=("Segoe UI", 8), fg="#94A3B8", bg="#162032").pack(anchor="w")

        self.lbl_active_mode_badge = tk.Label(hdr, text="Active: Audio to Audio", font=FONT_BOLD, fg="#34D399", bg="#0F172A", padx=12, pady=4, relief="solid", bd=1)
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

        tk.Label(self.sidebar, text="CONVERSION MODES", font=("Segoe UI", 8, "bold"), fg="#64748B", bg="#162032").pack(anchor="w", pady=(0, 6), padx=4)

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

        # FFmpeg Status Banner
        ff_box = tk.LabelFrame(self.sidebar, text="FFmpeg Status", font=("Segoe UI", 8, "bold"), fg="#38BDF8", bg="#162032", padx=6, pady=6)
        ff_box.pack(fill=tk.X, side=tk.BOTTOM, pady=4)
        status_txt = "Detected" if FFMPEG_PATH else "Not Detected"
        status_color = "#34D399" if FFMPEG_PATH else "#F87171"
        tk.Label(ff_box, text=f"Engine: {status_txt}", font=FONT_REGULAR, fg=status_color, bg="#162032").pack(anchor="w")

        # Right Main Content Area
        self.content_area = tk.Frame(container, bg="#0B1120")
        self.content_area.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # --- Card 1: File Queue & Toolbar ---
        card_queue = tk.Frame(self.content_area, bg="#162032", padx=10, pady=8, relief="solid", bd=1)
        card_queue.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

        # Toolbar above Queue
        tb = tk.Frame(card_queue, bg="#162032")
        tb.pack(fill=tk.X, pady=(0, 6))

        tk.Label(tb, text="Media Files Queue:", font=FONT_HEADING, fg="#F8FAFC", bg="#162032").pack(side=tk.LEFT)

        tk.Button(tb, text="+ Add File(s)...", command=self._add_files_dialog, font=FONT_BOLD, bg="#0284C7", fg="#FFFFFF", activebackground="#0369A1", relief="flat", padx=10, pady=3, cursor="hand2").pack(side=tk.RIGHT, padx=2)
        tk.Button(tb, text="+ Add Folder...", command=self._add_folder_dialog, font=FONT_REGULAR, bg="#334155", fg="#F8FAFC", activebackground="#475569", relief="flat", padx=8, pady=3, cursor="hand2").pack(side=tk.RIGHT, padx=2)
        tk.Button(tb, text="Remove Selected", command=self._remove_selected_files, font=FONT_REGULAR, bg="#DC2626", fg="#FFFFFF", activebackground="#B91C1C", relief="flat", padx=8, pady=3, cursor="hand2").pack(side=tk.RIGHT, padx=2)
        tk.Button(tb, text="Move Up", command=self._move_file_up, font=FONT_REGULAR, bg="#334155", fg="#F8FAFC", relief="flat", padx=8, pady=3, cursor="hand2").pack(side=tk.RIGHT, padx=2)
        tk.Button(tb, text="Move Down", command=self._move_file_down, font=FONT_REGULAR, bg="#334155", fg="#F8FAFC", relief="flat", padx=8, pady=3, cursor="hand2").pack(side=tk.RIGHT, padx=2)
        
        # Sort Button with Dropdown Menu (tk.Button matching exact toolbar dimensions)
        self.btn_sort = tk.Button(
            tb,
            text="Sort...",
            command=self._show_sort_menu,
            font=FONT_REGULAR,
            bg="#334155",
            fg="#F8FAFC",
            activebackground="#475569",
            activeforeground="#FFFFFF",
            relief="flat",
            padx=8,
            pady=3,
            cursor="hand2"
        )
        self.sort_menu = tk.Menu(
            self.btn_sort,
            tearoff=False,
            bg="#1E293B",
            fg="#F8FAFC",
            activebackground="#0284C7",
            activeforeground="#FFFFFF"
        )
        self.sort_menu.add_command(label="Name (A to Z)", command=lambda: self._sort_queue("name_asc"))
        self.sort_menu.add_command(label="Name (Z to A)", command=lambda: self._sort_queue("name_desc"))
        self.sort_menu.add_separator()
        self.sort_menu.add_command(label="Size (Small to Large)", command=lambda: self._sort_queue("size_asc"))
        self.sort_menu.add_command(label="Size (Large to Small)", command=lambda: self._sort_queue("size_desc"))
        self.sort_menu.add_separator()
        self.sort_menu.add_command(label="Type / Format", command=lambda: self._sort_queue("type_asc"))
        self.btn_sort.pack(side=tk.RIGHT, padx=2)

        tk.Button(tb, text="Clear", command=self._clear_all_files, font=FONT_REGULAR, bg="#475569", fg="#F8FAFC", activebackground="#334155", activeforeground="#FFFFFF", relief="flat", padx=8, pady=3, cursor="hand2").pack(side=tk.RIGHT, padx=2)

        # Dark Style for Treeview Queue
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "MediaQueue.Treeview",
            background="#0F172A",
            foreground="#F8FAFC",
            fieldbackground="#0F172A",
            rowheight=24,
            font=FONT_REGULAR,
            borderwidth=0
        )
        style.configure(
            "MediaQueue.Treeview.Heading",
            background="#1E293B",
            foreground="#38BDF8",
            font=FONT_BOLD,
            borderwidth=1,
            relief="flat"
        )
        style.map("MediaQueue.Treeview", background=[("selected", "#0284C7")], foreground=[("selected", "#FFFFFF")])

        tree_frame = tk.Frame(card_queue, bg="#0F172A")
        tree_frame.pack(fill=tk.BOTH, expand=True)

        cols = ("idx", "name", "size", "type", "duration", "path")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", style="MediaQueue.Treeview", selectmode="extended")
        self.tree.heading("idx", text="#", anchor="center")
        self.tree.heading("name", text="File Name", anchor="w")
        self.tree.heading("size", text="Size", anchor="e")
        self.tree.heading("type", text="Type", anchor="center")
        self.tree.heading("duration", text="Duration", anchor="center")
        self.tree.heading("path", text="Absolute Path", anchor="w")

        self.tree.column("idx", width=40, anchor="center")
        self.tree.column("name", width=280, anchor="w")
        self.tree.column("size", width=80, anchor="e")
        self.tree.column("type", width=70, anchor="center")
        self.tree.column("duration", width=80, anchor="center")
        self.tree.column("path", width=380, anchor="w")

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
        self.frame_dynamic_settings = tk.LabelFrame(config_cols, text="Conversion Settings", font=FONT_BOLD, fg="#38BDF8", bg="#162032", padx=8, pady=6)
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
        tk.Label(n_row, text="Custom Name:", font=FONT_REGULAR, fg="#94A3B8", bg="#162032").pack(side=tk.LEFT)
        self.ent_name = tk.Entry(n_row, textvariable=self.out_name_var, font=FONT_REGULAR)
        self.ent_name.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
        tk.Label(n_row, text="(Optional for 1 file)", font=("Segoe UI", 7, "italic"), fg="#64748B", bg="#162032").pack(side=tk.RIGHT)

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
            text="Execute Conversion",
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

        self.lbl_exec_status = tk.Label(exec_bar, text="Ready. Add media files and click Execute Conversion.", font=FONT_REGULAR, fg="#94A3B8", bg="#162032")
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

        self.progress_bar = ttk.Progressbar(exec_bar, mode="determinate", length=200, maximum=100)
        self.progress_bar.pack(side=tk.RIGHT, padx=4)

        # Monospace Terminal Log
        log_frame = tk.Frame(card_exec, bg="#070B14")
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.txt_log = tk.Text(log_frame, bg="#070B14", fg="#38BDF8", insertbackground="#38BDF8", font=FONT_MONO, height=8, wrap="word", relief="flat")
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=log_scroll.set)
        self.txt_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self._log(f"Media Converter Pro v3.0 initialized. Active mode: {self.MODE_SPECS[self.current_mode]['title']}")
        if FFMPEG_PATH:
            self._log(f"FFmpeg binary detected: {FFMPEG_PATH}")
        else:
            self._log("[WARNING] FFmpeg executable was not detected. Please install FFmpeg or run: pip install imageio-ffmpeg")
        self._log("System ready. Drag and drop media files, reorder or sort, and convert.")
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
                    p, msg = payload
                    self.progress_bar.configure(value=p)
                    self.lbl_progress_pct.configure(text=f"{int(round(p))}%")
                    if msg:
                        self.lbl_exec_status.configure(text=msg, fg="#38BDF8")
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

        self._render_dynamic_settings()
        self._log(f"Switched operation to: {spec['title']}")

    def _render_dynamic_settings(self):
        for w in self.frame_dynamic_settings.winfo_children():
            w.destroy()

        mode = self.current_mode

        if mode == "audio2audio":
            # Target Format & Bitrate
            r1 = tk.Frame(self.frame_dynamic_settings, bg="#162032")
            r1.pack(fill=tk.X, pady=2)
            tk.Label(r1, text="Target Format:", font=FONT_REGULAR, fg="#94A3B8", bg="#162032", width=14, anchor="w").pack(side=tk.LEFT)
            cb_fmt = ttk.Combobox(r1, textvariable=self.a2a_format_var, values=AUDIO_FORMATS, font=FONT_REGULAR, width=12)
            cb_fmt.pack(side=tk.LEFT, padx=(0, 8))

            tk.Label(r1, text="Bitrate:", font=FONT_REGULAR, fg="#94A3B8", bg="#162032").pack(side=tk.LEFT)
            cb_br = ttk.Combobox(r1, textvariable=self.a2a_bitrate_var, values=["320k (Maximum)", "256k (High)", "192k (Standard)", "128k (Medium)", "64k (Low)"], font=FONT_REGULAR, width=16)
            cb_br.pack(side=tk.LEFT, padx=4)

            # Sample Rate & Channels
            r2 = tk.Frame(self.frame_dynamic_settings, bg="#162032")
            r2.pack(fill=tk.X, pady=2)
            tk.Label(r2, text="Sample Rate:", font=FONT_REGULAR, fg="#94A3B8", bg="#162032", width=14, anchor="w").pack(side=tk.LEFT)
            cb_sr = ttk.Combobox(r2, textvariable=self.a2a_samplerate_var, values=["original", "48000", "44100", "22050"], font=FONT_REGULAR, width=12)
            cb_sr.pack(side=tk.LEFT, padx=(0, 8))

            tk.Label(r2, text="Channels:", font=FONT_REGULAR, fg="#94A3B8", bg="#162032").pack(side=tk.LEFT)
            cb_ch = ttk.Combobox(r2, textvariable=self.a2a_channels_var, values=["original", "stereo", "mono"], font=FONT_REGULAR, width=16)
            cb_ch.pack(side=tk.LEFT, padx=4)

            # Audio Normalization
            r3 = tk.Frame(self.frame_dynamic_settings, bg="#162032")
            r3.pack(fill=tk.X, pady=2)
            chk_norm = tk.Checkbutton(r3, text="Apply Loudness Normalization (-filter:a loudnorm)", variable=self.a2a_normalize_var, font=FONT_REGULAR, fg="#F8FAFC", bg="#162032", selectcolor="#0F172A", cursor="hand2")
            chk_norm.pack(side=tk.LEFT)

        elif mode == "video2audio":
            # Target Audio Format & Bitrate
            r1 = tk.Frame(self.frame_dynamic_settings, bg="#162032")
            r1.pack(fill=tk.X, pady=2)
            tk.Label(r1, text="Audio Format:", font=FONT_REGULAR, fg="#94A3B8", bg="#162032", width=14, anchor="w").pack(side=tk.LEFT)
            cb_fmt = ttk.Combobox(r1, textvariable=self.v2a_format_var, values=AUDIO_FORMATS, font=FONT_REGULAR, width=12)
            cb_fmt.pack(side=tk.LEFT, padx=(0, 8))

            tk.Label(r1, text="Audio Bitrate:", font=FONT_REGULAR, fg="#94A3B8", bg="#162032").pack(side=tk.LEFT)
            cb_br = ttk.Combobox(r1, textvariable=self.v2a_bitrate_var, values=["320k (Maximum)", "256k (High)", "192k (Standard)", "128k (Medium)", "64k (Low)"], font=FONT_REGULAR, width=16)
            cb_br.pack(side=tk.LEFT, padx=4)

            # Sample Rate & Channels
            r2 = tk.Frame(self.frame_dynamic_settings, bg="#162032")
            r2.pack(fill=tk.X, pady=2)
            tk.Label(r2, text="Sample Rate:", font=FONT_REGULAR, fg="#94A3B8", bg="#162032", width=14, anchor="w").pack(side=tk.LEFT)
            cb_sr = ttk.Combobox(r2, textvariable=self.v2a_samplerate_var, values=["original", "48000", "44100", "22050"], font=FONT_REGULAR, width=12)
            cb_sr.pack(side=tk.LEFT, padx=(0, 8))

            tk.Label(r2, text="Channels:", font=FONT_REGULAR, fg="#94A3B8", bg="#162032").pack(side=tk.LEFT)
            cb_ch = ttk.Combobox(r2, textvariable=self.v2a_channels_var, values=["original", "stereo", "mono"], font=FONT_REGULAR, width=16)
            cb_ch.pack(side=tk.LEFT, padx=4)

        elif mode == "video2video":
            # Target Format & Codec
            r1 = tk.Frame(self.frame_dynamic_settings, bg="#162032")
            r1.pack(fill=tk.X, pady=2)
            tk.Label(r1, text="Target Format:", font=FONT_REGULAR, fg="#94A3B8", bg="#162032", width=14, anchor="w").pack(side=tk.LEFT)
            cb_fmt = ttk.Combobox(r1, textvariable=self.v2v_format_var, values=VIDEO_FORMATS, font=FONT_REGULAR, width=12)
            cb_fmt.pack(side=tk.LEFT, padx=(0, 8))

            tk.Label(r1, text="Quality (CRF):", font=FONT_REGULAR, fg="#94A3B8", bg="#162032").pack(side=tk.LEFT)
            cb_crf = ttk.Combobox(r1, textvariable=self.v2v_quality_var, values=["Near Lossless (CRF 18)", "Balanced (CRF 23 - Recommended)", "Compact / Fast (CRF 28)"], font=FONT_REGULAR, width=22)
            cb_crf.pack(side=tk.LEFT, padx=4)

            # Preset Speed & Scaling
            r2 = tk.Frame(self.frame_dynamic_settings, bg="#162032")
            r2.pack(fill=tk.X, pady=2)
            tk.Label(r2, text="Preset Speed:", font=FONT_REGULAR, fg="#94A3B8", bg="#162032", width=14, anchor="w").pack(side=tk.LEFT)
            cb_pre = ttk.Combobox(r2, textvariable=self.v2v_preset_var, values=["slow (Better Compression)", "medium (Standard Speed)", "fast (Quick Encoding)", "veryfast (Real-time)"], font=FONT_REGULAR, width=22)
            cb_pre.pack(side=tk.LEFT, padx=(0, 8))

            tk.Label(r2, text="Resolution:", font=FONT_REGULAR, fg="#94A3B8", bg="#162032").pack(side=tk.LEFT)
            cb_res = ttk.Combobox(r2, textvariable=self.v2v_resolution_var, values=["Original (Keep Resolution)", "1080p (1920x1080)", "720p (1280x720)", "480p (854x480)"], font=FONT_REGULAR, width=20)
            cb_res.pack(side=tk.LEFT, padx=4)

            # Stream Copy Option
            r3 = tk.Frame(self.frame_dynamic_settings, bg="#162032")
            r3.pack(fill=tk.X, pady=2)
            chk_copy = tk.Checkbutton(r3, text="Stream Copy (-c copy: Instant & Lossless, skips re-encoding if container matches)", variable=self.v2v_stream_copy_var, font=FONT_REGULAR, fg="#F8FAFC", bg="#162032", selectcolor="#0F172A", cursor="hand2")
            chk_copy.pack(side=tk.LEFT)

        self._apply_pointer_cursors(self.frame_dynamic_settings)

    # --- File Management, Drag-and-Drop, Sorting ---
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
                self._log(f"[+] Added {added} media file(s) to queue.")
                self.lbl_exec_status.configure(text=f"[+] Added {added} file(s).", fg="#34D399")
        except Exception as e:
            messagebox.showerror("File Dialog Error", str(e), parent=self.root)

    def _add_folder_dialog(self):
        spec = self.MODE_SPECS[self.current_mode]
        try:
            folder = filedialog.askdirectory(parent=self.root, title="Select Directory to Import Media Files")
            if folder:
                accepted_exts = spec["accepted_inputs"]
                added = 0
                for root_dir, _, filenames in os.walk(folder):
                    for fn in filenames:
                        p = Path(root_dir) / fn
                        if p.suffix.lower() in accepted_exts:
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

    def _show_sort_menu(self):
        try:
            x = self.btn_sort.winfo_rootx()
            y = self.btn_sort.winfo_rooty() + self.btn_sort.winfo_height()
            self.sort_menu.tk_popup(x, y)
        except Exception:
            pass
        finally:
            self.sort_menu.grab_release()

    def _sort_queue(self, sort_key: str):
        if not self.file_items:
            return
        if sort_key == "name_asc":
            self.file_items.sort(key=lambda p: p.name.lower())
            self._log("[QUEUE] Sorted queue by Name (A to Z).")
        elif sort_key == "name_desc":
            self.file_items.sort(key=lambda p: p.name.lower(), reverse=True)
            self._log("[QUEUE] Sorted queue by Name (Z to A).")
        elif sort_key == "size_asc":
            self.file_items.sort(key=lambda p: p.stat().st_size if p.exists() else 0)
            self._log("[QUEUE] Sorted queue by Size (Small to Large).")
        elif sort_key == "size_desc":
            self.file_items.sort(key=lambda p: p.stat().st_size if p.exists() else 0, reverse=True)
            self._log("[QUEUE] Sorted queue by Size (Large to Small).")
        elif sort_key == "type_asc":
            self.file_items.sort(key=lambda p: (p.suffix.lower(), p.name.lower()))
            self._log("[QUEUE] Sorted queue by File Extension / Type.")
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
            self.tree.insert("", "end", iid=str(idx), values=(idx + 1, path.name, sz_str, ext, "Analyzing...", str(path)))

        cnt = len(self.file_items)
        tot_str = format_file_size(total_bytes)
        self.lbl_queue_status.configure(text=f"{cnt} file(s) queued ({tot_str}) -- Drag rows or use Sort/Move buttons to reorder.")

        # Non-blocking duration discovery thread for queue items
        if cnt > 0 and FFMPEG_PATH:
            threading.Thread(target=self._async_fetch_durations, daemon=True).start()

    def _async_fetch_durations(self):
        items_snapshot = list(enumerate(self.file_items))
        for idx, path in items_snapshot:
            if not path.exists():
                continue
            dur = get_media_duration(path, FFMPEG_PATH)
            dur_str = format_duration(dur)
            self.root.after(0, lambda i=idx, d=dur_str: self._update_row_duration(i, d))

    def _update_row_duration(self, idx: int, duration_str: str):
        iid = str(idx)
        if self.tree.exists(iid):
            cur_vals = list(self.tree.item(iid, "values"))
            if len(cur_vals) >= 5:
                cur_vals[4] = duration_str
                self.tree.item(iid, values=cur_vals)

    # --- Drag & Drop Reordering within Treeview ---
    def _on_tree_press(self, event):
        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return
        self._drag_start_idx = int(row_id)
        if not (0 <= self._drag_start_idx < len(self.file_items)):
            return

        item_name = self.file_items[self._drag_start_idx].name
        if self._drag_popup:
            try:
                self._drag_popup.destroy()
            except Exception:
                pass

        self._drag_popup = tk.Toplevel(self.root)
        self._drag_popup.overrideredirect(True)
        self._drag_popup.attributes("-alpha", 0.85)
        self._drag_popup.configure(bg="#0284C7")

        lbl = tk.Label(
            self._drag_popup,
            text=f"  Dragging: {item_name}  ",
            font=FONT_BOLD,
            fg="#FFFFFF",
            bg="#0284C7",
            padx=8,
            pady=4,
            relief="solid",
            bd=1
        )
        lbl.pack()
        self._drag_popup.geometry(f"+{event.x_root + 15}+{event.y_root + 15}")
        self.tree.configure(cursor="fleur")

    def _on_tree_motion(self, event):
        if self._drag_start_idx is None:
            return
        if self._drag_popup:
            self._drag_popup.geometry(f"+{event.x_root + 15}+{event.y_root + 15}")
        target_row = self.tree.identify_row(event.y)
        if target_row:
            for item in self.tree.get_children():
                tags = list(self.tree.item(item, "tags"))
                if "drag_target" in tags:
                    tags.remove("drag_target")
                    self.tree.item(item, tags=tags)
            tags = list(self.tree.item(target_row, "tags"))
            if "drag_target" not in tags:
                tags.append("drag_target")
            self.tree.item(target_row, tags=tags)
            self.tree.tag_configure("drag_target", background="#0369A1", foreground="#FFFFFF")

    def _on_tree_release(self, event):
        if self._drag_start_idx is None:
            return
        target_row = self.tree.identify_row(event.y)
        if self._drag_popup:
            self._drag_popup.destroy()
            self._drag_popup = None

        if target_row:
            dest_idx = int(target_row)
            if dest_idx != self._drag_start_idx and 0 <= dest_idx < len(self.file_items):
                item = self.file_items.pop(self._drag_start_idx)
                self.file_items.insert(dest_idx, item)
                self._update_queue_view()
                self.tree.selection_set(str(dest_idx))
                self._log(f"[REORDER] Moved item '{item.name}' from position #{self._drag_start_idx + 1} to #{dest_idx + 1}")

        self._drag_start_idx = None
        self.tree.configure(cursor="")

    def _on_cancel_drag(self, event):
        if self._drag_popup:
            self._drag_popup.destroy()
            self._drag_popup = None
        self._drag_start_idx = None
        self.tree.configure(cursor="")

    def _on_external_drop(self, files):
        added_count = 0
        switched_mode_note = None

        all_dropped = []
        for f in files:
            p_str = f.decode("utf-8", errors="ignore") if isinstance(f, bytes) else str(f)
            all_dropped.append(Path(p_str))

        audio_exts = set(AUDIO_FORMATS)
        video_exts = set(VIDEO_FORMATS)
        all_media_exts = audio_exts | video_exts

        dropped_suffixes = {p.suffix.lower() for p in all_dropped if p.is_file()}
        if dropped_suffixes:
            if (dropped_suffixes & video_exts) and not (dropped_suffixes & audio_exts):
                if self.current_mode == "audio2audio":
                    self._switch_mode("video2video")
                    switched_mode_note = "Auto-switched to 'Video to Video' mode"
            elif (dropped_suffixes & audio_exts) and not (dropped_suffixes & video_exts):
                if self.current_mode in ("video2video", "video2audio"):
                    self._switch_mode("audio2audio")
                    switched_mode_note = "Auto-switched to 'Audio to Audio' mode"

        accepted_exts = self.MODE_SPECS[self.current_mode]["accepted_inputs"]

        for p in all_dropped:
            try:
                if p.is_dir():
                    for root_dir, _, filenames in os.walk(p):
                        for fn in filenames:
                            child_p = Path(root_dir) / fn
                            if child_p.suffix.lower() in accepted_exts or child_p.suffix.lower() in all_media_exts:
                                if child_p not in self.file_items:
                                    self.file_items.append(child_p)
                                    added_count += 1
                elif p.is_file():
                    if p.suffix.lower() in accepted_exts or p.suffix.lower() in all_media_exts:
                        if p not in self.file_items:
                            self.file_items.append(p)
                            added_count += 1
            except Exception as e:
                self._log(f"[-] Drag & drop import error: {e}")

        if added_count > 0:
            self._update_queue_view()
            status_text = f"[+] Added {added_count} media file(s) via drag-and-drop."
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
            messagebox.showwarning("Missing Media Files", "Please add at least one audio or video file before executing.", parent=self.root)
            return

        out_d = Path(self.out_dir_var.get().strip())
        if not out_d:
            messagebox.showwarning("Missing Destination Path", "Please specify a valid destination folder.", parent=self.root)
            return

        ffmpeg_bin = FFMPEG_PATH or detect_ffmpeg()
        if not ffmpeg_bin:
            messagebox.showerror("FFmpeg Missing", "FFmpeg executable is required for conversion. Please install FFmpeg or run: pip install imageio-ffmpeg", parent=self.root)
            return

        # Snapshot all parameters on the main UI thread to guarantee thread-safe background execution
        params = {
            "mode": self.current_mode,
            "files": list(self.file_items),
            "out_dir": out_d,
            "custom_name": self.out_name_var.get().strip(),
            "open_dest": self.open_dest_var.get(),
            "ffmpeg_bin": ffmpeg_bin,
            "a2a_format": self.a2a_format_var.get().lower().strip(),
            "a2a_bitrate": self.a2a_bitrate_var.get().split()[0].lower(),
            "a2a_samplerate": self.a2a_samplerate_var.get().strip(),
            "a2a_channels": self.a2a_channels_var.get().strip(),
            "a2a_normalize": bool(self.a2a_normalize_var.get()),
            "v2a_format": self.v2a_format_var.get().lower().strip(),
            "v2a_bitrate": self.v2a_bitrate_var.get().split()[0].lower(),
            "v2a_samplerate": self.v2a_samplerate_var.get().strip(),
            "v2a_channels": self.v2a_channels_var.get().strip(),
            "v2v_format": self.v2v_format_var.get().lower().strip(),
            "v2v_stream_copy": bool(self.v2v_stream_copy_var.get()),
            "v2v_quality": self.v2v_quality_var.get(),
            "v2v_preset": self.v2v_preset_var.get().split()[0].lower(),
            "v2v_resolution": self.v2v_resolution_var.get(),
        }

        self.is_running = True
        self.btn_execute.configure(state="disabled", text="Converting...", cursor="")
        self.progress_bar.configure(value=0)
        self.lbl_progress_pct.configure(text="0%", fg="#38BDF8")
        self.lbl_exec_status.configure(text=f"Starting conversion of {len(params['files'])} file(s)...", fg="#38BDF8")

        thread = threading.Thread(target=self._run_worker, args=(params,), daemon=True)
        thread.start()

    def _run_worker(self, params: Dict[str, Any]):
        mode = params["mode"]
        files = params["files"]
        out_dir = params["out_dir"]
        custom_name = params["custom_name"]
        ffmpeg_bin = params["ffmpeg_bin"]
        os.makedirs(out_dir, exist_ok=True)

        self._log("=" * 60)
        self._log(f"Starting Operation: {self.MODE_SPECS[mode]['title']}")
        self._log(f"Input Files Count: {len(files)}")
        self._log(f"Target Directory: {out_dir}")
        self._log("=" * 60)

        t0 = time.time()
        success_count = 0
        error_count = 0
        generated_files = []
        total_files = len(files)

        try:
            for idx, src_file in enumerate(files, 1):
                if not src_file.exists():
                    self._log(f"[-] [{idx}/{total_files}] File missing: {src_file.name}")
                    error_count += 1
                    continue

                try:
                    # 1. Determine output filename and command arguments
                    if mode == "audio2audio":
                        tgt_ext = params["a2a_format"]
                        if not tgt_ext.startswith("."): tgt_ext = f".{tgt_ext}"
                        br = params["a2a_bitrate"]
                        sr = params["a2a_samplerate"]
                        ch = params["a2a_channels"]
                        norm = params["a2a_normalize"]

                        args = build_audio_ffmpeg_args(tgt_ext, bitrate=br, sample_rate=sr, channels=ch, normalize=norm)

                    elif mode == "video2audio":
                        tgt_ext = params["v2a_format"]
                        if not tgt_ext.startswith("."): tgt_ext = f".{tgt_ext}"
                        br = params["v2a_bitrate"]
                        sr = params["v2a_samplerate"]
                        ch = params["v2a_channels"]

                        args = ["-vn"] + build_audio_ffmpeg_args(tgt_ext, bitrate=br, sample_rate=sr, channels=ch)

                    elif mode == "video2video":
                        tgt_ext = params["v2v_format"]
                        if not tgt_ext.startswith("."): tgt_ext = f".{tgt_ext}"
                        stream_copy = params["v2v_stream_copy"]

                        if stream_copy:
                            args = ["-c", "copy"]
                        else:
                            v_codec, a_codec = VIDEO_CODECS_MAP.get(tgt_ext, ("libx264", "aac"))
                            args = ["-c:v", v_codec]

                            # CRF extraction
                            crf_raw = params["v2v_quality"]
                            m_crf = re.search(r"CRF\s*(\d+)", crf_raw)
                            crf_val = m_crf.group(1) if m_crf else "23"

                            # Preset extraction
                            pre_raw = params["v2v_preset"]

                            if v_codec in ("libx264", "libx265"):
                                args.extend(["-crf", crf_val, "-preset", pre_raw])
                                if tgt_ext == ".mp4":
                                    args.extend(["-pix_fmt", "yuv420p"])
                            elif v_codec == "libvpx-vp9":
                                args.extend(["-crf", crf_val, "-b:v", "0"])
                            elif v_codec in ("mpeg4", "mpeg2video"):
                                args.extend(["-q:v", "3"])

                            # Resolution Scaling
                            res_choice = params["v2v_resolution"]
                            if "1080p" in res_choice:
                                args.extend(["-vf", "scale=1920:1080:force_original_aspect_ratio=decrease,pad=1920:1080:(ow-iw)/2:(oh-ih)/2"])
                            elif "720p" in res_choice:
                                args.extend(["-vf", "scale=1280:720:force_original_aspect_ratio=decrease,pad=1280:720:(ow-iw)/2:(oh-ih)/2"])
                            elif "480p" in res_choice:
                                args.extend(["-vf", "scale=854:480:force_original_aspect_ratio=decrease,pad=854:480:(ow-iw)/2:(oh-ih)/2"])

                            # Audio stream with optional fallback if source lacks audio
                            args.extend(["-c:a", a_codec])
                            if a_codec in ("aac", "libmp3lame", "mp3", "libvorbis"):
                                args.extend(["-b:a", "192k"])
                            elif a_codec == "libopus":
                                args.extend(["-b:a", "128k"])

                    # Output Filename Determination
                    if len(files) == 1 and custom_name:
                        out_fname = custom_name if custom_name.lower().endswith(tgt_ext) else f"{custom_name}{tgt_ext}"
                    else:
                        out_fname = f"{src_file.stem}{tgt_ext}"

                    dst_path = out_dir / out_fname
                    # Collision handling
                    counter = 1
                    while dst_path.exists():
                        dst_path = out_dir / f"{src_file.stem}_{counter}{tgt_ext}"
                        counter += 1

                    self._update_progress_bar(idx, total_files, 0.0, f"[{idx}/{total_files}] Converting {src_file.name}...")
                    self._log(f"[{idx}/{total_files}] Processing: {src_file.name} -> {dst_path.name}")

                    rec = execute_media_transcode(
                        src_file=src_file,
                        dst_file=dst_path,
                        cmd_args=args,
                        ffmpeg_bin=ffmpeg_bin,
                        log_callback=self._log,
                        progress_callback=lambda p, msg: self._update_progress_bar(idx, total_files, p, msg)
                    )

                    in_sz_str = format_file_size(rec["input_size_bytes"])
                    out_sz_str = format_file_size(rec["output_size_bytes"])
                    self._update_progress_bar(idx, total_files, 100.0, f"[{idx}/{total_files}] Converted {src_file.name}")
                    self._log(f"[+] Output: {dst_path} ({out_sz_str}, took {rec['elapsed_seconds']}s)")
                    generated_files.append(dst_path)
                    success_count += 1

                except Exception as ex:
                    self._log(f"[-] Conversion error on {src_file.name}: {ex}")
                    error_count += 1

        except Exception as global_ex:
            self._log(f"[FATAL ERROR] {global_ex}")

        finally:
            elapsed_total = time.time() - t0
            self._log("-" * 60)
            self._log(f"Batch completed in {elapsed_total:.2f}s | Success: {success_count} | Failed: {error_count}")

            if hasattr(self, "ui_queue"):
                self.ui_queue.put(("finish", {
                    "success_count": success_count,
                    "error_count": error_count,
                    "elapsed_total": elapsed_total,
                    "out_dir": out_dir,
                    "open_dest": params.get("open_dest", False)
                }))

    def _on_execution_finished(self, data: Dict[str, Any]):
        """Executes completion logic safely on the main UI thread."""
        self.btn_execute.configure(state="normal", text="Execute Conversion", cursor="hand2")
        self.is_running = False

        success_count = data.get("success_count", 0)
        error_count = data.get("error_count", 0)
        elapsed_total = data.get("elapsed_total", 0.0)
        out_dir = data.get("out_dir")
        open_dest = data.get("open_dest", False)

        if success_count > 0:
            self.progress_bar.configure(value=100.0)
            self.lbl_progress_pct.configure(text="100%", fg="#34D399")
            self.lbl_exec_status.configure(text=f"[OK] Completed {success_count} file(s) in {elapsed_total:.2f}s!", fg="#34D399")
            if open_dest and out_dir and Path(out_dir).exists():
                try:
                    os.startfile(str(out_dir))
                except Exception:
                    pass
            messagebox.showinfo("Conversion Succeeded", f"Converted {success_count} file(s) successfully in {elapsed_total:.2f}s!\n\nDestination:\n{out_dir}", parent=self.root)
        else:
            self.lbl_progress_pct.configure(fg="#F87171")
            self.lbl_exec_status.configure(text=f"[Error] Failed: {error_count} error(s).", fg="#F87171")
            messagebox.showerror("Conversion Failed", "Conversion completed with errors. See terminal log for details.", parent=self.root)

    def _update_progress_bar(self, cur_idx: int, total: int, item_pct: float, status_msg: Optional[str] = None):
        base_pct = ((cur_idx - 1) / total) * 100.0
        cur_total = base_pct + (item_pct / total)
        p = min(100.0, max(0.0, cur_total))
        msg = status_msg or f"[{cur_idx}/{total}] Converting... ({int(round(p))}%)"
        if hasattr(self, "ui_queue"):
            self.ui_queue.put(("progress", (p, msg)))


def run_media_toolkit():
    """Launches the modern Media Converter Pro Desktop GUI Dashboard."""
    app = MediaToolkitDashboardGUI()
    app.root.mainloop()


if __name__ == "__main__":
    run_media_toolkit()
