#!/usr/bin/env python3
"""
Universal Video to Audio Converter & Extractor
Extracts and converts audio tracks from video into any audio format:
- Single video file
- Multiple selected video files
- Full folder of videos (batch)

Supported Video Inputs (19 Formats):
.mp4, .mkv, .mov, .avi, .wmv, .webm, .flv, .mpeg, .mpg, .m4v, .3gp, .ts, .mts, .m2ts, .vob, .ogv, .asf, .divx, .f4v

Supported Audio Outputs (21 Formats):
.mp3, .wav, .aac, .m4a, .flac, .ogg, .opus, .wma, .aiff, .aif, .alac, .amr, .mid, .midi, .ac3, .eac3, .ape, .dsf, .dff, .caf, .au
"""

import os
import re
import sys
import time
import json
import shutil
import argparse
import subprocess
from pathlib import Path

VIDEO_FORMATS = [
    ".mp4", ".mkv", ".mov", ".avi", ".wmv", ".webm", ".flv", ".mpeg",
    ".mpg", ".m4v", ".3gp", ".ts", ".mts", ".m2ts", ".vob", ".ogv",
    ".asf", ".divx", ".f4v"
]

AUDIO_FORMATS = [
    ".mp3", ".wav", ".aac", ".m4a", ".flac", ".ogg", ".opus", ".wma",
    ".aiff", ".aif", ".alac", ".amr", ".mid", ".midi", ".ac3", ".eac3",
    ".ape", ".dsf", ".dff", ".caf", ".au"
]

def detect_ffmpeg():
    # 1. System PATH
    exe = shutil.which("ffmpeg")
    if exe:
        return exe
        
    # 2. Local venv
    venv_exe = Path(__file__).resolve().parent / ".venv" / "Scripts" / "ffmpeg.exe"
    if venv_exe.is_file():
        os.environ["PATH"] = str(venv_exe.parent) + os.pathsep + os.environ.get("PATH", "")
        return str(venv_exe)

    # 3. imageio_ffmpeg bundled binary
    try:
        import imageio_ffmpeg
        loc = imageio_ffmpeg.get_ffmpeg_exe()
        if loc and os.path.isfile(loc):
            os.environ["PATH"] = str(Path(loc).parent) + os.pathsep + os.environ.get("PATH", "")
            return loc
    except Exception:
        pass

    # 4. Common locations
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

def get_duration(file_path, ffmpeg_bin):
    """Extracts duration using ffmpeg metadata banner without requiring ffprobe."""
    if not ffmpeg_bin:
        return None
    try:
        proc = subprocess.run(
            [ffmpeg_bin, "-hide_banner", "-i", str(file_path)],
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

def build_audio_ffmpeg_args(target_ext, bitrate="192k"):
    ext = target_ext.lower()
    args = ["-vn"]
    
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
    elif ext in (".dsf", ".dff"):
        args.extend(["-c:a", "copy"])
    else:
        args.extend(["-b:a", bitrate])
        
    return args

def pick_media_inputs(supported_formats, media_type="video"):
    selected_files = []
    source_dir = None
    is_folder = False

    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        print("\nChoose input selection mode:")
        print("  [1] Select MULTIPLE files (Hold Ctrl/Shift to choose multiple files)")
        print("  [2] Select a SINGLE file")
        print("  [3] Select a FULL FOLDER (Process all matching files in directory)")
        print("  [4] Type or paste path(s) manually")
        choice = input("Enter choice [1/2/3/4] (default: 1): ").strip()
        
        ext_filter = [(f"{media_type.title()} Files", "*" + ";*".join(supported_formats)), ("All Files", "*.*")]
        
        if choice == "2":
            single_path = filedialog.askopenfilename(title=f"Select Single {media_type.title()} File", filetypes=ext_filter)
            if single_path:
                p = Path(single_path)
                selected_files = [p]
                source_dir = p.parent
        elif choice == "3":
            folder_path = filedialog.askdirectory(title=f"Select Folder Containing {media_type.title()} Files")
            if folder_path:
                source_dir = Path(folder_path)
                is_folder = True
        elif choice == "4":
            pass
        else:
            multi_paths = filedialog.askopenfilenames(title=f"Select One or Multiple {media_type.title()} Files", filetypes=ext_filter)
            if multi_paths:
                selected_files = [Path(p) for p in multi_paths]
                source_dir = selected_files[0].parent
        root.destroy()
    except Exception:
        pass
        
    if not selected_files and not is_folder:
        raw = input(f"Enter {media_type} path(s) or folder (separate multiple files with comma): ").strip().strip('"').strip("'")
        if raw:
            if "," in raw:
                parts = [p.strip().strip('"').strip("'") for p in raw.split(",") if p.strip()]
            elif ";" in raw:
                parts = [p.strip().strip('"').strip("'") for p in raw.split(";") if p.strip()]
            else:
                parts = [raw]

            if len(parts) == 1 and Path(parts[0]).is_dir():
                is_folder = True
                source_dir = Path(parts[0])
            else:
                selected_files = [Path(p) for p in parts]
                source_dir = selected_files[0].parent

    return selected_files, is_folder, source_dir

def select_input_filter():
    print("\n" + "=" * 55)
    print("SELECT VIDEO INPUT FILTER")
    print("=" * 55)
    print("  [0] ALL Supported Video Formats (*.*)")
    for idx, ext in enumerate(VIDEO_FORMATS, 1):
        print(f"  [{idx:2d}] {ext}")
    print("=" * 55)

    in_choice = input("Enter input format choice [0 for all, or 1-19, or type extension] (default: 0): ").strip()
    if not in_choice or in_choice == "0":
        return "all"
    if in_choice.isdigit() and 1 <= int(in_choice) <= len(VIDEO_FORMATS):
        return VIDEO_FORMATS[int(in_choice) - 1]
    if in_choice.startswith(".") and in_choice.lower() in VIDEO_FORMATS:
        return in_choice.lower()
    if ("." + in_choice.lower()) in VIDEO_FORMATS:
        return "." + in_choice.lower()
    print(f"Unrecognized choice '{in_choice}'. Defaulting to 'all'.")
    return "all"

def select_target_output():
    print("\n" + "=" * 55)
    print("SELECT TARGET AUDIO OUTPUT FORMAT")
    print("=" * 55)
    for idx, ext in enumerate(AUDIO_FORMATS, 1):
        print(f"  [{idx:2d}] {ext}")
    print("=" * 55)

    out_choice = input("Enter target audio format [1-21, or type extension e.g. .mp3] (default: 1 [.mp3]): ").strip()
    if not out_choice or out_choice == "1":
        return ".mp3"
    if out_choice.isdigit() and 1 <= int(out_choice) <= len(AUDIO_FORMATS):
        return AUDIO_FORMATS[int(out_choice) - 1]
    if out_choice.startswith(".") and out_choice.lower() in AUDIO_FORMATS:
        return out_choice.lower()
    if ("." + out_choice.lower()) in AUDIO_FORMATS:
        return "." + out_choice.lower()
    print(f"Unrecognized choice '{out_choice}'. Defaulting to '.mp3'.")
    return ".mp3"

def convert_video_to_audio(src_file, dst_file, target_ext, ffmpeg_bin="ffmpeg", bitrate="192k"):
    cmd = [ffmpeg_bin, "-y", "-i", str(src_file)]
    audio_flags = build_audio_ffmpeg_args(target_ext, bitrate)
    cmd.extend(audio_flags)
    cmd.extend(["-nostats", "-progress", "pipe:1", str(dst_file)])
    
    duration_sec = get_duration(src_file, ffmpeg_bin)
    start_time = time.time()
    last_reported = -1
    log_tail = []

    proc = subprocess.Popen(
        cmd,
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

        if duration_sec and line_str.startswith("out_time_ms="):
            try:
                ms_val = float(line_str.split("=", 1)[1])
                curr_sec = ms_val / 1_000_000.0
                pct = min(100.0, (curr_sec / duration_sec) * 100.0)
                whole = int(pct)
                if whole >= last_reported + 10:
                    last_reported = whole
                    print(f"    Progress: {pct:5.1f}% [{curr_sec:.1f}s / {duration_sec:.1f}s]")
            except Exception:
                pass
        elif line_str == "progress=end":
            print("    Progress: 100.0% [Completed]")

    ret = proc.wait()
    if ret != 0:
        tail_msg = "\n".join(log_tail[-20:])
        raise RuntimeError(f"FFmpeg extraction failed (code {ret}):\n{tail_msg}")

    elapsed = time.time() - start_time
    return {
        "input": str(src_file),
        "output": str(dst_file),
        "elapsed_seconds": round(elapsed, 2),
        "input_size_bytes": src_file.stat().st_size,
        "output_size_bytes": dst_file.stat().st_size if dst_file.exists() else 0,
        "status": "Success"
    }

def main():
    parser = argparse.ArgumentParser(description="Universal Video to Audio Converter & Extractor (Single file, Multiple files, or Full folder)")
    parser.add_argument("--input", "-i", nargs="+", help="Input video file(s) or folder path", default=None)
    parser.add_argument("--filter", "-f", help="Video input filter (e.g. .mp4 or all)", default=None)
    parser.add_argument("--output-format", "-o", help="Target audio output format (e.g. .mp3)", default=None)
    parser.add_argument("--bitrate", "-b", default="192k", help="Audio bitrate (e.g. 320k, 192k, 128k)")
    args = parser.parse_args()

    print("=" * 60)
    print("   UNIVERSAL VIDEO TO AUDIO EXTRACTOR (19 In -> 21 Out)")
    print("   [Supports: 1 File, Multiple Files, or Full Folder]")
    print("=" * 60)

    ffmpeg_bin = detect_ffmpeg()
    if not ffmpeg_bin:
        print("\nERROR: FFmpeg executable could not be found.")
        print("Run: pip install imageio-ffmpeg")
        sys.exit(1)
    print(f"FFmpeg detected: {ffmpeg_bin}")

    if args.input:
        if len(args.input) == 1 and Path(args.input[0]).is_dir():
            source_dir = Path(args.input[0])
            is_folder = True
            selected_files = []
        else:
            selected_files = [Path(p) for p in args.input]
            source_dir = selected_files[0].parent
            is_folder = False
    else:
        selected_files, is_folder, source_dir = pick_media_inputs(VIDEO_FORMATS, media_type="video")

    if not is_folder and not selected_files:
        print("Error: No valid input files or directory specified.")
        sys.exit(1)

    input_filter = args.filter if args.filter else (select_input_filter() if (is_folder or len(selected_files) > 1) else "all")
    target_output = args.output_format if args.output_format else select_target_output()
    if not target_output.startswith("."):
        target_output = "." + target_output

    if is_folder:
        if input_filter == "all":
            target_exts = set(VIDEO_FORMATS)
        else:
            target_exts = {input_filter}
        files = [f for f in source_dir.iterdir() if f.is_file() and f.suffix.lower() in target_exts]
        if not files:
            files = [f for f in source_dir.rglob("*") if f.is_file() and f.suffix.lower() in target_exts]
    else:
        if input_filter != "all":
            files = [f for f in selected_files if f.is_file() and f.suffix.lower() == input_filter]
        else:
            files = [f for f in selected_files if f.is_file() and f.suffix.lower() in VIDEO_FORMATS]

    if not files:
        print(f"No video files matching filter '{input_filter}' found for audio extraction.")
        sys.exit(1)

    out_dir = source_dir / f"extracted_audio_{target_output.lstrip('.')}"
    out_dir.mkdir(parents=True, exist_ok=True)
    print(f"\nProcessing {len(files)} file(s). Bitrate: {args.bitrate}. Destination: {out_dir}")

    conversion_records = []
    for idx, f in enumerate(files, 1):
        dst = out_dir / f"{f.stem}{target_output}"
        print(f"\n[{idx}/{len(files)}] Extracting: {f.name} -> {dst.name}")
        try:
            rec = convert_video_to_audio(f, dst, target_output, ffmpeg_bin=ffmpeg_bin, bitrate=args.bitrate)
            conversion_records.append(rec)
            print(f"  OK: Finished in {rec['elapsed_seconds']}s")
        except Exception as e:
            print(f"  FAILED: {e}")
            conversion_records.append({"input": str(f), "output": str(dst), "status": "Failed", "error": str(e)})

    # Summary
    print("\n" + "=" * 70)
    print(f"{'Source Video':<28} | {'Video Size':<10} | {'Audio Size':<10} | {'Savings':<7} | {'Time'}")
    print("-" * 70)
    for rec in conversion_records:
        if rec.get("status") == "Success":
            in_mb = rec["input_size_bytes"] / (1024 * 1024)
            out_mb = rec["output_size_bytes"] / (1024 * 1024)
            savings = ((in_mb - out_mb) / in_mb * 100.0) if in_mb > 0 else 0.0
            fname = Path(rec["input"]).name[:27]
            print(f"{fname:<28} | {in_mb:7.2f} MB | {out_mb:7.2f} MB | {savings:5.1f}% | {rec['elapsed_seconds']}s")
        else:
            fname = Path(rec["input"]).name[:27]
            print(f"{fname:<28} | FAILED")

    report_path = out_dir / "audio_extraction_summary.json"
    with open(report_path, "w", encoding="utf-8") as f:
        json.dump({"type": "video_to_audio", "records": conversion_records}, f, indent=2)
    print(f"\nReport saved to: {report_path}")

if __name__ == "__main__":
    main()
