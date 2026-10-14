# Python Automation & Modern Desktop GUI Suite Pro

[![Python](https://img.shields.io/badge/Python-3.10%20%7C%203.11%20%7C%203.12%20%7C%203.13-3776AB.svg?style=flat-square&logo=python&logoColor=white)](https://www.python.org/)
[![GUI Engine](https://img.shields.io/badge/GUI-Tkinter%20%2F%20ttk%20Modern%20Dark-0F172A.svg?style=flat-square)](https://docs.python.org/3/library/tkinter.html)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%20%2F%2011%20x64-0078D6.svg?style=flat-square&logo=windows&logoColor=white)](https://microsoft.com)
[![License](https://img.shields.io/badge/License-MIT-green.svg?style=flat-square)](LICENSE)
[![Status](https://img.shields.io/badge/Build-Production%20Ready%20%E2%9C%93-10B981.svg?style=flat-square)]()
[![Code Aesthetics](https://img.shields.io/badge/Design-Zero%20Emojis%20%7C%20Cyber%20Dark-0284C7.svg?style=flat-square)]()

A comprehensive, enterprise-grade suite of modern Python desktop GUI dashboards, media conversion workstations, certificate mail-merge engines, and document processing systems. Designed from the ground up with elegant dark glassmorphism aesthetics, responsive multi-threaded workers, and zero external cloud dependencies.

Developed and maintained by **mohatamim**.

---

## Table of Contents

1. [Suite Overview & Architecture](#suite-overview--architecture)
2. [Interactive Desktop GUI Toolkits (GUI/)](#interactive-desktop-gui-toolkits)
   - [YouTube Playlist & Stream Downloader Pro](#1-youtube-playlist--stream-downloader-pro)
   - [Media Toolkit Pro](#2-media-toolkit-pro)
   - [AutoCrat Pro Enterprise Document Automation](#3-autocrat-pro-enterprise-document-automation)
   - [All-in-One PDF & Image Processing Toolkit Pro](#4-all-in-one-pdf--image-processing-toolkit-pro)
3. [Standalone CLI Converters & Engines](#standalone-cli-converters--engines)
4. [Adobe Illustrator Bulk Generator](#adobe-illustrator-bulk-generator)
5. [Repository Structure](#repository-structure)
6. [System Requirements & Installation](#system-requirements--installation)
7. [Quick-Start User Guide](#quick-start-user-guide)
8. [Automated Verification & Test Suite](#automated-verification--test-suite)
9. [Publishing & Uploading to GitHub](#publishing--uploading-to-github)
10. [License](#license)

---

## Suite Overview & Architecture

```
+-----------------------------------------------------------------------------------+
|               Python Automation & Desktop GUI Suite Pro Architecture              |
+-----------------------------------------------------------------------------------+
|                                                                                   |
|  [ GUI/YouTube_Playlist_Downloader.ipynb ]                                        |
|  * Multi-Threaded Engine (1 to 8 concurrent download streams)                     |
|  * Real-Time Task Monitor Modal with individual progress bars & ETA               |
|  * Embedded Direct CDN Link Exporter with single / select-all clipboard copying   |
|  * Custom Cyber Spectrum Audio/Video Player with animated spectrum visualizer     |
|                                                                                   |
|  [ GUI/media_toolkit.ipynb ] & [ media_toolkit.py ]                               |
|  * Video Transcoder (MP4, MKV, AVI, WEBM, MOV, FLV)                               |
|  * Audio Extractor & Transcoder (MP3, WAV, AAC, FLAC, OGG, M4A)                   |
|  * Integrated Windows Media Player OCX with dynamic frequency spectrum analyzer   |
|  * High-density dark Treeview with alternating row striping & context menus       |
|                                                                                   |
|  [ GUI/autocrat_pro.ipynb ] & [ autocrat_pro.py ]                                 |
|  * 100% Local PowerPoint COM Vector PDF Engine (preserves custom fonts)           |
|  * Dynamic QR Code generation & dynamic photo injection (local path or web URL)   |
|  * Interactive Spreadsheet Table Viewer modal & 1-Click placeholder auto-mapper   |
|  * 5 Generation Modes (Individual PDFs, Consolidated Master, Batch Folders, etc.)|
|                                                                                   |
|  [ GUI/pdf_toolkit.ipynb ] & [ pdf_toolkit.py ]                                   |
|  * 10 Document Operations: Merge, Split, Compress, Img2PDF, PDF2Img, OCR, etc.    |
|  * Interactive Document Inspector modal (PDF text stream reader, Image thumbnail) |
|  * 27/27 Passing Automated Acceptance Criteria Tests                              |
+-----------------------------------------------------------------------------------+
```

---

## Interactive Desktop GUI Toolkits

All primary graphical applications reside in the `GUI/` directory and can be launched directly within Jupyter Notebooks, JupyterLab, or VS Code.

### 1. YouTube Playlist & Stream Downloader Pro
**File:** [`GUI/YouTube_Playlist_Downloader.ipynb`](GUI/YouTube_Playlist_Downloader.ipynb)

A media acquisition interface featuring:
- **Real-Time Task Monitor Modal:** Click the primary progress bar to open an interactive telemetry modal showing individual, real-time download progress bars, throughput speeds, and time-to-completion (ETA) for each concurrent thread.
- **Custom Cyber Spectrum Player:** Preview downloaded media directly inside a custom modal player with animated cyber audio visualizers, seek controls, and volume adjustments.
- **Concurrent Thread Selector:** Choose how many videos download simultaneously (1 to 8 threads).
- **In-App Direct CDN Link Exporter:** Extract raw CDN stream URLs directly into a dashboard treeview with single or batch clipboard export options.
- **Byte-Accurate File Telemetry:** Exact byte formatting in B, KB, MB, and GB.
- **Batch Audio Extraction:** Convert video files to MP3 audio files with ID3 tag handling.

### 2. Media Toolkit Pro
**Files:** [`GUI/media_toolkit.ipynb`](GUI/media_toolkit.ipynb) | [`media_toolkit.py`](media_toolkit.py)

A conversion workstation for video and audio production:
- **Video & Audio Transcoding:** Convert between video containers (MP4, MKV, AVI, WEBM, MOV, FLV) and audio formats (MP3, WAV, AAC, FLAC, OGG, M4A).
- **Embedded Media Player Preview:** Audition queued media inside a built-in player modal with live audio spectrum waveform animation.
- **Context Menu Operations:** Right-click context menus for quick actions: preview media, open containing folder, move up/down, remove, and clear.
- **Standardized Dark UI:** Row height 28 with alternating dark highlights (`#0F172A` / `#141E33`), uniform padding, and pointer cursor feedback.

### 3. AutoCrat Pro Enterprise Document Automation
**Files:** [`GUI/autocrat_pro.ipynb`](GUI/autocrat_pro.ipynb) | [`autocrat_pro.py`](autocrat_pro.py)

An enterprise certificate and document generation system merging PowerPoint (`.pptx`) templates with Excel (`.xlsx`, `.xls`) or CSV datasets:
- **Native Vector PDF Engine:** Utilizes Microsoft PowerPoint COM automation (`win32com` / `comtypes`) for vector rendering and preserved typography.
- **5 Operating Modes:**
  1. *Individual PDFs:* Generates personalized certificates per spreadsheet row.
  2. *Consolidated Master Document:* Merges all recipient slides into a single multi-page master PDF/PPTX.
  3. *Folder Batching:* Processes multiple event directories and spreadsheets simultaneously.
  4. *PDF Consolidation Engine:* Merges standalone PDFs using `pypdf.PdfWriter`.
  5. *Pre-Flight Template Validator:* Audits template tags against dataset columns before execution.
- **Interactive Data Table Viewer:** View loaded spreadsheet rows, columns, and verify data before generating documents.
- **1-Click Auto-Mapper:** Intelligently pairs presentation slide `{{placeholders}}` with spreadsheet columns.
- **Dynamic Assets:** Generates dynamic QR codes (`{{qr}}`) and injects remote or local photos (`{{image}}` / `{{photo}}`).

### 4. All-in-One PDF & Image Processing Toolkit Pro
**Files:** [`GUI/pdf_toolkit.ipynb`](GUI/pdf_toolkit.ipynb) | [`pdf_toolkit.py`](pdf_toolkit.py)

A multi-tool document processing suite:
- **10 Core Operations:**
  1. *Merge Multiple PDFs:* Combines PDFs with Table of Contents bookmark generation.
  2. *Split PDF Pages:* Extracts single pages, custom ranges (e.g. `1-3, 5, 8-10`), or fixed chunks.
  3. *Compress & Optimize PDF:* Stream optimization and downsampling with Low, Medium, and High presets.
  4. *Images to PDF:* Assembles PNG, JPG, WEBP, TIFF, BMP, and GIF images into PDF documents.
  5. *PDF to Images:* Renders PDF pages to images (PNG, JPEG, WEBP, TIFF, BMP) at customizable DPI (150-600 DPI).
  6. *Protect / Encrypt PDF:* Secures documents with AES-256 encryption, passwords, and permissions.
  7. *Unlock / Decrypt PDF:* Removes restrictions from authorized password-protected documents.
  8. *OCR PDF:* Extracts text layers from scanned documents using Tesseract OCR (Bengali and English).
  9. *Markdown to PDF:* Renders GitHub-flavored Markdown into styled PDFs.
  10. *PDF to Markdown:* Extracts text layers and embedded figures into Markdown files.
- **Interactive Document Inspector Modal:** Double-click or press `Ctrl+P` to view metadata, page count, live text stream extraction across pages, or scaled image thumbnails.
- **Verified Stability:** Backed by 27 automated unit and integration tests (`test_pdf_toolkit.py`).

---

## Standalone CLI Converters & Engines

For headless or automated batch processing without a GUI, standalone scripts are provided:

| Script / Notebook | Type | Description |
| :--- | :---: | :--- |
| **`video_to_video_converter.py`** | CLI / Engine | Converts between video formats with configurable resolution and bitrate. |
| **`video_to_audio_converter.py`** | CLI / Engine | Extracts audio tracks from video files in batch mode. |
| **`audio_to_audio_converter.py`** | CLI / Engine | Transcodes audio files between sample rates, bitrates, and formats. |
| **`pdf_toolkit.py`** | Core Engine | Headless command-line interface for all 10 PDF and image operations. |
| **`autocrat_pro.py`** | Core Engine | Headless document generation engine for scheduled mail-merge jobs. |

---

## Adobe Illustrator Bulk Generator

- **`CSV_XLSX_to_Illustrator_JSON.ipynb`**: Transforms spreadsheet columns and recipient lists into formatted JSON data matrices.
- **`Illustrator_Bulk_Artwork_Generator.jsx`**: Adobe Illustrator ExtendScript that reads JSON datasets and automates bulk artwork generation, vector text replacement, and export directly inside Adobe Illustrator.

---

## Repository Structure

```
.
|-- GUI/
|   |-- YouTube_Playlist_Downloader.ipynb   # Multi-threaded YouTube acquisition dashboard
|   |-- media_toolkit.ipynb                 # Media conversion & preview dashboard
|   |-- autocrat_pro.ipynb                  # Document mail-merge & auto-mapper dashboard
|   `-- pdf_toolkit.ipynb                   # 10-operation PDF & image suite dashboard
|-- .assets/                                # UI icons and visual assets
|-- pdf_toolkit.py                          # Core PDF engine module
|-- test_pdf_toolkit.py                     # 27 automated tests for PDF engine
|-- autocrat_pro.py                         # Core AutoCrat document automation engine
|-- media_toolkit.py                        # Standalone media processing engine
|-- video_to_video_converter.py             # CLI video-to-video converter
|-- video_to_audio_converter.py             # CLI video-to-audio converter
|-- audio_to_audio_converter.py             # CLI audio-to-audio converter
|-- Illustrator_Bulk_Artwork_Generator.jsx  # Adobe Illustrator bulk artwork script
|-- CSV_XLSX_to_Illustrator_JSON.ipynb      # Dataset-to-Illustrator JSON pipeline
|-- sample_certificates.csv                 # Sample dataset for certificate testing
|-- sample_conference_badges.csv            # Sample dataset for badge generation
|-- sample_conference_badges.xlsx           # Excel format sample dataset
|-- requirements.txt                        # Complete Python dependencies
|-- .gitignore                              # Git exclusion rules
|-- LICENSE                                 # MIT License
`-- README.md                               # Repository documentation
```

---

## System Requirements & Installation

### Requirements
- **Operating System:** Windows 10 or Windows 11 (64-bit recommended)
- **Python Version:** Python 3.10, 3.11, 3.12, or 3.13
- **Microsoft Office:** PowerPoint 2016, 2019, 2021, or Microsoft 365 (Required for AutoCrat Pro COM rendering)
- **External Binaries (Optional):**
  - `ffmpeg` in system PATH (Recommended for advanced audio/video transcoding)
  - `tesseract-ocr` (Required for OCR operations in PDF Toolkit)

### Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/mohatamim/YOUR_REPOSITORY_NAME.git
   cd YOUR_REPOSITORY_NAME
   ```

2. **Create and activate a virtual environment (Recommended):**
   ```bash
   python -m venv .venv
   .venv\Scripts\activate
   ```

3. **Install dependencies:**
   ```bash
   pip install --upgrade pip
   pip install -r requirements.txt
   ```

---

## Quick-Start User Guide

### Launching the Graphical Dashboards
Open any notebook in `GUI/` (such as `GUI/pdf_toolkit.ipynb` or `GUI/YouTube_Playlist_Downloader.ipynb`) in VS Code, JupyterLab, or classic Jupyter Notebook, select your Python kernel, and run all cells.

### Running Acceptance Tests
Verify the PDF processing engine and its 27 test cases:
```bash
python test_pdf_toolkit.py
```
Expected output:
```
=================================================================
 SUMMARY: 27/27 Tests Passed.
=================================================================
```

---

## Publishing & Uploading to GitHub

This repository is already initialized and prepared with clean commits:

```bash
# 1. Create a new empty repository on your GitHub account (e.g. 'python-automation-gui-suite')

# 2. Add your GitHub remote repository URL
git remote add origin https://github.com/mohatamim/YOUR_REPOSITORY_NAME.git

# 3. Push to the main branch
git push -u origin main
```

---

## License

Distributed under the **MIT License**. Developed and maintained by **mohatamim**. See [LICENSE](LICENSE) for full details.
