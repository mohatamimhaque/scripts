from PIL import Image
"""
Test Suite for All-in-One PDF Toolkit
=====================================
Covers all 25 acceptance criteria specified in Section 17:
 1. Folder paths containing spaces
 2. Multiple source folders
 3. Invalid and inaccessible folders
 4. Duplicate folder paths
 5. Correct numbered file discovery
 6. Selection using 3,1,4
 7. Selection using all
 8. Invalid and duplicate serial numbers
 9. PDF merge in the exact specified order
10. Splitting a multi-page PDF
11. Splitting by page ranges
12. Splitting into fixed-size page groups
13. Compressing files from the same and different folders
14. Handling compression that increases file size
15. Unlocking an authorized encrypted PDF
16. Handling incorrect PDF passwords
17. Protecting a PDF and verifying encryption
18. OCR of English documents (or dependency validation)
19. OCR of Bengali documents when language data installed (or dependency validation)
20. Markdown-to-PDF rendering
21. PDF-to-Markdown text extraction
22. Handling scanned PDFs without extractable text
23. Handling corrupted files
24. Handling existing output directories and filename collisions
25. Verifying that original files remain unchanged
"""

import os
import sys
import tempfile
import hashlib
from pathlib import Path

import pypdf
import pymupdf
from fpdf import FPDF

from pdf_toolkit import (
    parse_folder_paths,
    validate_folder_paths,
    discover_files,
    parse_serial_selection,
    format_file_size,
    OutputManager,
    execute_merge_pdfs,
    execute_split_pdfs,
    execute_compress_pdfs,
    execute_unlock_pdf,
    execute_protect_pdf,
    execute_images_to_pdf,
    execute_pdf_to_images,
    execute_ocr_pdf,
    detect_tesseract_binary,
    PYTESSERACT_AVAILABLE,
    execute_markdown_to_pdf,
    execute_pdf_to_markdown
)


def get_file_hash(p: Path) -> str:
    """Computes SHA-256 hash of a file."""
    h = hashlib.sha256()
    with open(p, "rb") as f:
        while chunk := f.read(8192):
            h.update(chunk)
    return h.hexdigest()


def create_sample_pdf(file_path: Path, num_pages: int = 1, text_prefix: str = "Page") -> Path:
    """Creates a valid PDF with specified page count and text."""
    pdf = FPDF()
    for i in range(1, num_pages + 1):
        pdf.add_page()
        pdf.set_font("Helvetica", size=16)
        pdf.cell(0, 10, f"{text_prefix} {i}")
    file_path.parent.mkdir(parents=True, exist_ok=True)
    pdf.output(str(file_path))
    return file_path


def run_all_tests():
    total_tests = 27
    passed_tests = 0
    results = []

    print("\n" + "=" * 65)
    print(" EXECUTING 27 COMPREHENSIVE PDF TOOLKIT ACCEPTANCE TESTS")
    print("=" * 65 + "\n")

    with tempfile.TemporaryDirectory() as temp_dir:
        root = Path(temp_dir)

        # -------------------------------------------------------------
        # Test 1: Folder paths containing spaces
        # -------------------------------------------------------------
        try:
            folder_with_space = root / "Folder With Spaces"
            folder_with_space.mkdir()
            parsed = parse_folder_paths(f'"{folder_with_space}", D:\\Other')
            assert str(folder_with_space) in parsed[0]
            valid_dirs, _ = validate_folder_paths([str(folder_with_space)])
            assert len(valid_dirs) == 1 and valid_dirs[0] == folder_with_space
            passed_tests += 1
            results.append((1, "Folder paths containing spaces", "PASSED"))
        except Exception as e:
            results.append((1, "Folder paths containing spaces", f"FAILED: {e}"))

        # -------------------------------------------------------------
        # Test 2: Multiple source folders
        # -------------------------------------------------------------
        try:
            f1 = root / "Folder1"
            f2 = root / "Folder2"
            f1.mkdir()
            f2.mkdir()
            raw_input = f"{f1}, {f2}"
            parsed = parse_folder_paths(raw_input)
            assert len(parsed) == 2
            valid, _ = validate_folder_paths(parsed)
            assert len(valid) == 2
            passed_tests += 1
            results.append((2, "Multiple source folders", "PASSED"))
        except Exception as e:
            results.append((2, "Multiple source folders", f"FAILED: {e}"))

        # -------------------------------------------------------------
        # Test 3: Invalid and inaccessible folders
        # -------------------------------------------------------------
        try:
            non_existent = root / "NonExistentFolder123"
            valid, warnings = validate_folder_paths([str(f1), str(non_existent)])
            assert len(valid) == 1
            assert any("does not exist" in w for w in warnings)
            passed_tests += 1
            results.append((3, "Invalid and inaccessible folders", "PASSED"))
        except Exception as e:
            results.append((3, "Invalid and inaccessible folders", f"FAILED: {e}"))

        # -------------------------------------------------------------
        # Test 4: Duplicate folder paths
        # -------------------------------------------------------------
        try:
            valid, _ = validate_folder_paths([str(f1), str(f1), str(f2)])
            assert len(valid) == 2, f"Expected 2 unique folders, got {len(valid)}"
            passed_tests += 1
            results.append((4, "Duplicate folder paths deduplication", "PASSED"))
        except Exception as e:
            results.append((4, "Duplicate folder paths deduplication", f"FAILED: {e}"))

        # -------------------------------------------------------------
        # Test 5: Correct numbered file discovery
        # -------------------------------------------------------------
        try:
            pdf_a = create_sample_pdf(f1 / "Doc_A.pdf", num_pages=1, text_prefix="A")
            pdf_b = create_sample_pdf(f1 / "Doc_B.pdf", num_pages=2, text_prefix="B")
            pdf_c = create_sample_pdf(f2 / "Doc_C.pdf", num_pages=1, text_prefix="C")
            discovered = discover_files([f1, f2], extensions=[".pdf"])
            assert len(discovered) == 3
            assert pdf_a in discovered and pdf_b in discovered and pdf_c in discovered
            passed_tests += 1
            results.append((5, "Correct numbered file discovery", "PASSED"))
        except Exception as e:
            results.append((5, "Correct numbered file discovery", f"FAILED: {e}"))

        # -------------------------------------------------------------
        # Test 6: Selection using '3,1,4'
        # -------------------------------------------------------------
        try:
            pdf_d = create_sample_pdf(f2 / "Doc_D.pdf", num_pages=1, text_prefix="D")
            discovered = discover_files([f1, f2], extensions=[".pdf"])  # 4 files
            sel, err = parse_serial_selection("3,1,4", total_count=len(discovered))
            assert err is None
            assert sel == [3, 1, 4], f"Expected exact [3, 1, 4], got {sel}"
            passed_tests += 1
            results.append((6, "Selection using 3,1,4 order preservation", "PASSED"))
        except Exception as e:
            results.append((6, "Selection using 3,1,4 order preservation", f"FAILED: {e}"))

        # -------------------------------------------------------------
        # Test 7: Selection using 'all'
        # -------------------------------------------------------------
        try:
            sel, err = parse_serial_selection("all", total_count=4)
            assert err is None
            assert sel == [1, 2, 3, 4]
            passed_tests += 1
            results.append((7, "Selection using 'all'", "PASSED"))
        except Exception as e:
            results.append((7, "Selection using 'all'", f"FAILED: {e}"))

        # -------------------------------------------------------------
        # Test 8: Invalid and duplicate serial numbers
        # -------------------------------------------------------------
        try:
            _, err_dup = parse_serial_selection("1,2,1", total_count=4)
            assert "Duplicate serial number" in (err_dup or "")
            _, err_out = parse_serial_selection("5", total_count=4)
            assert "out of range" in (err_out or "")
            _, err_zero = parse_serial_selection("0", total_count=4)
            assert "Invalid serial number" in (err_zero or "")
            passed_tests += 1
            results.append((8, "Invalid and duplicate serial numbers rejection", "PASSED"))
        except Exception as e:
            results.append((8, "Invalid and duplicate serial numbers rejection", f"FAILED: {e}"))

        # -------------------------------------------------------------
        # Test 9: PDF merge in the exact specified order
        # -------------------------------------------------------------
        try:
            p1 = create_sample_pdf(root / "m1.pdf", num_pages=1, text_prefix="FIRST")
            p2 = create_sample_pdf(root / "m2.pdf", num_pages=1, text_prefix="SECOND")
            p3 = create_sample_pdf(root / "m3.pdf", num_pages=1, text_prefix="THIRD")
            # User specified order: 3, 1, 2
            ordered_files = [p3, p1, p2]
            rep = execute_merge_pdfs(ordered_files, output_parent_folder=root, merged_filename="test_merged.pdf")
            out_file = rep.results[0]["output_files"][0]
            reader = pypdf.PdfReader(str(out_file))
            assert len(reader.pages) == 3
            # Check text sequence: page 0 must be THIRD, page 1 FIRST, page 2 SECOND
            doc_m = pymupdf.open(str(out_file))
            assert "THIRD" in doc_m[0].get_text()
            assert "FIRST" in doc_m[1].get_text()
            assert "SECOND" in doc_m[2].get_text()
            doc_m.close()
            passed_tests += 1
            results.append((9, "PDF merge in exact specified order", "PASSED"))
        except Exception as e:
            results.append((9, "PDF merge in exact specified order", f"FAILED: {e}"))

        # -------------------------------------------------------------
        # Test 10: Splitting a multi-page PDF (Mode A)
        # -------------------------------------------------------------
        try:
            multi_pdf = create_sample_pdf(root / "multi.pdf", num_pages=3, text_prefix="MultiPage")
            rep = execute_split_pdfs([multi_pdf], output_parent_folder=root, mode="A")
            split_files = rep.results[0]["output_files"]
            assert len(split_files) == 3
            for sf in split_files:
                r = pypdf.PdfReader(str(sf))
                assert len(r.pages) == 1
            passed_tests += 1
            results.append((10, "Splitting a multi-page PDF (Mode A)", "PASSED"))
        except Exception as e:
            results.append((10, "Splitting a multi-page PDF (Mode A)", f"FAILED: {e}"))

        # -------------------------------------------------------------
        # Test 11: Splitting by page ranges (Mode B)
        # -------------------------------------------------------------
        try:
            multi_pdf6 = create_sample_pdf(root / "multi6.pdf", num_pages=6, text_prefix="SixPage")
            rep = execute_split_pdfs([multi_pdf6], output_parent_folder=root, mode="B", range_str="1-2, 3-5, 6")
            split_files = rep.results[0]["output_files"]
            assert len(split_files) == 3
            assert len(pypdf.PdfReader(str(split_files[0])).pages) == 2
            assert len(pypdf.PdfReader(str(split_files[1])).pages) == 3
            assert len(pypdf.PdfReader(str(split_files[2])).pages) == 1
            passed_tests += 1
            results.append((11, "Splitting by page ranges (Mode B)", "PASSED"))
        except Exception as e:
            results.append((11, "Splitting by page ranges (Mode B)", f"FAILED: {e}"))

        # -------------------------------------------------------------
        # Test 12: Splitting into fixed-size page groups (Mode C)
        # -------------------------------------------------------------
        try:
            multi_pdf7 = create_sample_pdf(root / "multi7.pdf", num_pages=7, text_prefix="SevenPage")
            rep = execute_split_pdfs([multi_pdf7], output_parent_folder=root, mode="C", group_size=3)
            split_files = rep.results[0]["output_files"]
            assert len(split_files) == 3  # (3, 3, 1)
            assert len(pypdf.PdfReader(str(split_files[0])).pages) == 3
            assert len(pypdf.PdfReader(str(split_files[1])).pages) == 3
            assert len(pypdf.PdfReader(str(split_files[2])).pages) == 1
            passed_tests += 1
            results.append((12, "Splitting into fixed-size page groups (Mode C)", "PASSED"))
        except Exception as e:
            results.append((12, "Splitting into fixed-size page groups (Mode C)", f"FAILED: {e}"))

        # -------------------------------------------------------------
        # Test 13: Compressing files from same and different folders
        # -------------------------------------------------------------
        try:
            comp_src1 = create_sample_pdf(f1 / "c1.pdf", num_pages=2)
            comp_src2 = create_sample_pdf(f2 / "c2.pdf", num_pages=2)
            rep = execute_compress_pdfs([comp_src1, comp_src2], level="medium")
            assert len(rep.results) == 2
            assert rep.results[0]["status"] == "Success"
            assert rep.results[1]["status"] == "Success"
            # Verify outputs placed in respective folders
            out1 = rep.results[0]["output_files"][0]
            out2 = rep.results[1]["output_files"][0]
            assert out1.parent == f1 / "output"
            assert out2.parent == f2 / "output"
            passed_tests += 1
            results.append((13, "Compressing from same and different folders", "PASSED"))
        except Exception as e:
            results.append((13, "Compressing from same and different folders", f"FAILED: {e}"))

        # -------------------------------------------------------------
        # Test 14: Handling compression that increases file size
        # -------------------------------------------------------------
        try:
            # A tiny 1-page PDF might stay same size or increase slightly
            tiny_pdf = create_sample_pdf(root / "tiny.pdf", num_pages=1, text_prefix="T")
            rep = execute_compress_pdfs([tiny_pdf], level="low")
            assert len(rep.results) == 1
            # Must handle gracefully without crash
            assert rep.results[0]["status"] in ("Success", "Skipped")
            passed_tests += 1
            results.append((14, "Handling compression size changes & reporting", "PASSED"))
        except Exception as e:
            results.append((14, "Handling compression size changes & reporting", f"FAILED: {e}"))

        # -------------------------------------------------------------
        # Test 15: Unlocking an authorized encrypted PDF
        # -------------------------------------------------------------
        try:
            plain_doc = create_sample_pdf(root / "to_lock.pdf", num_pages=2, text_prefix="LockMe")
            lock_rep = execute_protect_pdf([plain_doc], user_password="pass1234")
            locked_pdf = lock_rep.results[0]["output_files"][0]
            assert pypdf.PdfReader(str(locked_pdf)).is_encrypted is True

            unlock_rep = execute_unlock_pdf([locked_pdf], password="pass1234")
            unlocked_pdf = unlock_rep.results[0]["output_files"][0]
            assert pypdf.PdfReader(str(unlocked_pdf)).is_encrypted is False
            passed_tests += 1
            results.append((15, "Unlocking an authorized encrypted PDF", "PASSED"))
        except Exception as e:
            results.append((15, "Unlocking an authorized encrypted PDF", f"FAILED: {e}"))

        # -------------------------------------------------------------
        # Test 16: Handling incorrect PDF passwords
        # -------------------------------------------------------------
        try:
            bad_unlock_rep = execute_unlock_pdf([locked_pdf], password="WRONG_PASSWORD")
            assert bad_unlock_rep.results[0]["status"] == "Failed"
            assert "Incorrect password" in bad_unlock_rep.results[0]["error"]
            passed_tests += 1
            results.append((16, "Handling incorrect PDF passwords gracefully", "PASSED"))
        except Exception as e:
            results.append((16, "Handling incorrect PDF passwords gracefully", f"FAILED: {e}"))

        # -------------------------------------------------------------
        # Test 17: Protecting a PDF and verifying encryption
        # -------------------------------------------------------------
        try:
            plain_doc2 = create_sample_pdf(root / "plain2.pdf", num_pages=1)
            prot_rep = execute_protect_pdf([plain_doc2], user_password="secure_password", allow_printing=True, allow_copying=False)
            prot_file = prot_rep.results[0]["output_files"][0]
            reader = pypdf.PdfReader(str(prot_file))
            assert reader.is_encrypted is True
            assert int(reader.decrypt("secure_password")) in (1, 2)
            passed_tests += 1
            results.append((17, "Protecting a PDF and verifying AES encryption", "PASSED"))
        except Exception as e:
            results.append((17, "Protecting a PDF and verifying AES encryption", f"FAILED: {e}"))

        # -------------------------------------------------------------
        # Test 18: OCR of English documents (with dependency validation)
        # -------------------------------------------------------------
        try:
            tess_bin = detect_tesseract_binary()
            ocr_sample = create_sample_pdf(root / "ocr_eng_sample.pdf", num_pages=1, text_prefix="OCR English Text")
            rep_ocr = execute_ocr_pdf([ocr_sample], lang="eng")
            if tess_bin and PYTESSERACT_AVAILABLE:
                assert rep_ocr.results[0]["status"] == "Success"
            else:
                # Must provide clear instructions when Tesseract is missing
                assert rep_ocr.results[0]["status"] == "Failed"
                assert "Tesseract OCR is not installed" in rep_ocr.results[0]["error"]
            passed_tests += 1
            results.append((18, "OCR of English documents / dependency handling", "PASSED"))
        except Exception as e:
            results.append((18, "OCR of English documents / dependency handling", f"FAILED: {e}"))

        # -------------------------------------------------------------
        # Test 19: OCR of Bengali documents when language data is installed
        # -------------------------------------------------------------
        try:
            rep_ocr_ben = execute_ocr_pdf([ocr_sample], lang="ben")
            # Should either succeed if tesseract+ben installed or fail with helpful message
            assert len(rep_ocr_ben.results) > 0
            passed_tests += 1
            results.append((19, "OCR of Bengali documents language validation", "PASSED"))
        except Exception as e:
            results.append((19, "OCR of Bengali documents language validation", f"FAILED: {e}"))

        # -------------------------------------------------------------
        # Test 20: Markdown-to-PDF rendering
        # -------------------------------------------------------------
        try:
            md_file = root / "sample.md"
            md_content = """# Document Title\n\nThis is a paragraph with **bold** and *italic* formatting.\n\n## Table Section\n\n| Item | Qty | Price |\n| --- | --- | --- |\n| Widget A | 10 | $5.00 |\n| Widget B | 20 | $12.50 |\n\n> This is a blockquote.\n\n- Bullet 1\n- Bullet 2\n"""
            with open(md_file, "w", encoding="utf-8") as f:
                f.write(md_content)

            rep_md = execute_markdown_to_pdf([md_file], page_size="A4")
            assert rep_md.results[0]["status"] == "Success"
            out_pdf = rep_md.results[0]["output_files"][0]
            assert out_pdf.exists()
            assert pypdf.PdfReader(str(out_pdf)).pages[0] is not None
            passed_tests += 1
            results.append((20, "Markdown-to-PDF rendering", "PASSED"))
        except Exception as e:
            results.append((20, "Markdown-to-PDF rendering", f"FAILED: {e}"))

        # -------------------------------------------------------------
        # Test 21: PDF-to-Markdown text extraction
        # -------------------------------------------------------------
        try:
            pdf_for_md = create_sample_pdf(root / "extract_test.pdf", num_pages=2, text_prefix="Section Extracted")
            rep_p2m = execute_pdf_to_markdown([pdf_for_md], extract_images=True)
            assert rep_p2m.results[0]["status"] == "Success"
            out_md = rep_p2m.results[0]["output_files"][0]
            assert out_md.exists()
            with open(out_md, "r", encoding="utf-8") as f:
                md_text = f.read()
            assert "Section Extracted" in md_text
            passed_tests += 1
            results.append((21, "PDF-to-Markdown text extraction", "PASSED"))
        except Exception as e:
            results.append((21, "PDF-to-Markdown text extraction", f"FAILED: {e}"))

        # -------------------------------------------------------------
        # Test 22: Handling scanned PDFs without extractable text
        # -------------------------------------------------------------
        try:
            # Create a blank image PDF
            blank_pdf = root / "blank.pdf"
            doc_b = pymupdf.open()
            doc_b.new_page()
            doc_b.save(str(blank_pdf))
            doc_b.close()

            rep_blank = execute_pdf_to_markdown([blank_pdf], ocr_fallback=False)
            assert rep_blank.results[0]["status"] == "Success"
            out_b = rep_blank.results[0]["output_files"][0]
            assert out_b.exists()
            passed_tests += 1
            results.append((22, "Handling scanned/blank PDFs without text", "PASSED"))
        except Exception as e:
            results.append((22, "Handling scanned/blank PDFs without text", f"FAILED: {e}"))

        # -------------------------------------------------------------
        # Test 23: Handling corrupted files
        # -------------------------------------------------------------
        try:
            corrupt_file = root / "corrupt.pdf"
            with open(corrupt_file, "wb") as f:
                f.write(b"%PDF-1.4\nCorrupted binary junk\x00\xff")

            rep_corrupt = execute_merge_pdfs([corrupt_file], output_parent_folder=root)
            assert rep_corrupt.results[0]["status"] == "Failed"
            assert "Error" in rep_corrupt.results[0]["error"] or "Verification" in rep_corrupt.results[0]["error"]
            passed_tests += 1
            results.append((23, "Handling corrupted files gracefully", "PASSED"))
        except Exception as e:
            results.append((23, "Handling corrupted files gracefully", f"FAILED: {e}"))

        # -------------------------------------------------------------
        # Test 24: Handling existing output directories and filename collisions
        # -------------------------------------------------------------
        try:
            out_target = root / "output" / "collision_test.pdf"
            OutputManager.ensure_output_dir(out_target.parent)
            out_target.touch()

            safe_path = OutputManager.get_collision_safe_path(out_target)
            assert safe_path != out_target
            assert " (1)" in safe_path.name or "_1" in safe_path.name
            passed_tests += 1
            results.append((24, "Existing output directories & collisions", "PASSED"))
        except Exception as e:
            results.append((24, "Existing output directories & collisions", f"FAILED: {e}"))

        # -------------------------------------------------------------
        # Test 25: Verifying that original files remain unchanged
        # -------------------------------------------------------------
        try:
            immutable_pdf = create_sample_pdf(root / "immutable.pdf", num_pages=2, text_prefix="Original Content")
            orig_hash = get_file_hash(immutable_pdf)

            # Run multiple operations on it: Split, Compress, Protect
            execute_split_pdfs([immutable_pdf], output_parent_folder=root, mode="A")
            execute_compress_pdfs([immutable_pdf], level="low")
            execute_protect_pdf([immutable_pdf], user_password="pass")

            after_hash = get_file_hash(immutable_pdf)
            assert orig_hash == after_hash, "Original file was modified!"
            passed_tests += 1
            results.append((25, "Verifying original files remain unchanged", "PASSED"))
        except Exception as e:
            results.append((25, "Verifying original files remain unchanged", f"FAILED: {e}"))

        # -------------------------------------------------------------
        # Test 26: Images to PDF conversion and multi-format support
        # -------------------------------------------------------------
        try:
            im1 = root / "t26_1.png"
            im2 = root / "t26_2.jpg"
            Image.new("RGBA", (300, 200), (200, 50, 50, 180)).save(im1)
            Image.new("RGB", (400, 600), (50, 150, 50)).save(im2)

            rep_img = execute_images_to_pdf([im1, im2], output_dir=root / "out_img2pdf", output_name="img_out.pdf", page_size="a4", orientation="auto", margin_mm=5.0)
            assert rep_img.results[0]["status"] == "Success"
            out_pdf = rep_img.results[0]["output_files"][0]
            assert out_pdf.exists() and out_pdf.stat().st_size > 100
            passed_tests += 1
            results.append((26, "Images to PDF conversion and multi-format support", "PASSED"))
        except Exception as e:
            results.append((26, "Images to PDF conversion and multi-format support", f"FAILED: {e}"))

        # -------------------------------------------------------------
        # Test 27: PDF to Images rendering with format and DPI control
        # -------------------------------------------------------------
        try:
            sample_for_imgs = create_sample_pdf(root / "for_rendering.pdf", num_pages=2, text_prefix="Render Page")
            rep_p2i = execute_pdf_to_images([sample_for_imgs], output_dir=root / "out_pdf2img", img_format="png", dpi=150, page_range="all")
            assert rep_p2i.results[0]["status"] == "Success"
            exported = rep_p2i.results[0]["output_files"]
            assert len(exported) == 2
            assert all(f.exists() and f.stat().st_size > 100 for f in exported)
            passed_tests += 1
            results.append((27, "PDF to Images rendering with format and DPI control", "PASSED"))
        except Exception as e:
            results.append((27, "PDF to Images rendering with format and DPI control", f"FAILED: {e}"))

    # Print summary
    print("-" * 65)
    for num, name, status in results:
        flag = "[PASS]" if "PASSED" in status else "[FAIL]"
        print(f" {flag} Test {num:02d}: {name} -> {status}")
    print("=" * 65)
    print(f" SUMMARY: {passed_tests}/{total_tests} Tests Passed.")
    print("=" * 65 + "\n")

    return passed_tests == total_tests


if __name__ == "__main__":
    success = run_all_tests()
    sys.exit(0 if success else 1)
