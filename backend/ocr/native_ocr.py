"""
Khanna Travels & Holidays — Native Windows Media OCR Wrapper
Executes hardware-accelerated Windows.Media.Ocr on images and rendered PDF pages.
Requires no external third-party C++ binaries and works out of the box on Windows.
"""

import os
import platform
import subprocess
import shutil
from typing import Optional


def run_windows_ocr(image_path: str) -> str:
    """
    Executes Windows.Media.Ocr.OcrEngine via PowerShell script on an image file.
    Returns extracted raw text. Used on Windows localhost.
    """
    if not os.path.exists(image_path):
        return ""
        
    script_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "run_ocr.ps1"))
    if not os.path.exists(script_path):
        return ""

    try:
        proc = subprocess.run(
            ["powershell", "-ExecutionPolicy", "Bypass", "-File", script_path, os.path.abspath(image_path)],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=25
        )
        if proc.returncode == 0 and proc.stdout:
            return proc.stdout.strip()
    except Exception as e:
        print(f"Native Windows OCR error: {e}")
        
    return ""


def run_linux_ocr(image_path: str, psm: int = 3) -> str:
    """
    Executes native Tesseract OCR on Linux (Render container, Docker, Ubuntu/Debian).
    psm: 3 = Fully automatic page segmentation (suitable for full passport page)
    psm: 6 = Uniform single block of text (suitable for MRZ crop)
    Returns extracted raw text.
    """
    if not os.path.exists(image_path):
        return ""

    tesseract_bin = shutil.which("tesseract") or "/usr/bin/tesseract"
    if not os.path.exists(tesseract_bin) and not shutil.which("tesseract"):
        print(f"[OCR] Tesseract binary not found at {tesseract_bin}")
        return ""

    try:
        proc = subprocess.run(
            [tesseract_bin, os.path.abspath(image_path), "stdout", "-l", "eng", "--psm", str(psm)],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=30
        )
        if proc.returncode == 0 and proc.stdout:
            return proc.stdout.strip()
        elif proc.stderr:
            print(f"[OCR] Tesseract stderr notice: {proc.stderr.strip()[:200]}")
    except Exception as e:
        print(f"Native Linux Tesseract OCR error: {e}")

    return ""


def run_system_ocr(image_path: str, psm: int = 3) -> str:
    """
    Automatically selects the appropriate OCR engine based on the host OS:
    - Windows localhost: Windows Media OCR (hardware-accelerated, native Windows UWP)
    - Linux / Render: Tesseract OCR (standard open-source C++ engine)
    With automatic cross-platform fallback.
    """
    current_os = platform.system().lower()

    if "windows" in current_os:
        # Windows localhost: use existing Windows Media OCR
        text = run_windows_ocr(image_path)
        if text.strip():
            return text
        # Fallback to Tesseract if Windows Media OCR produced no text and Tesseract is installed
        if shutil.which("tesseract"):
            return run_linux_ocr(image_path, psm=psm)
        return text

    # Linux (Render container, Docker, Ubuntu/Debian)
    return run_linux_ocr(image_path, psm=psm)

