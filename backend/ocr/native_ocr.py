"""
Khanna Travels & Holidays — Native Windows Media OCR Wrapper
Executes hardware-accelerated Windows.Media.Ocr on images and rendered PDF pages.
Requires no external third-party C++ binaries and works out of the box on Windows.
"""

import os
import subprocess
from typing import Optional


def run_windows_ocr(image_path: str) -> str:
    """
    Executes Windows.Media.Ocr.OcrEngine via PowerShell script on an image file.
    Returns extracted raw text.
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
