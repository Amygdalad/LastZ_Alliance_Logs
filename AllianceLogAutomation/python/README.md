# Alliance Log Automation (Python)

Python rewrite of the PowerShell automation used to drive MEmu, dismiss ads/popups, and navigate to the Alliance screen to capture logs.

## Quick start
1. Install Python 3.11+ and Tesseract OCR (make sure `tesseract.exe` is on PATH or pass `--tesseract-path`).
2. `pip install -r requirements.txt`
3. Run:
   ```
   python -m alliance_automation.main --indexes 59 57 173 --memu-root "Z:\Program Files\Microvirt\MEmu" --instance-root "Z:\Program Files\Microvirt\MEmu\MemuHyperv VMs"
   ```

## Notable behavior
- After each close/X tap, waits 5 seconds and re-OCRs via Tesseract to see if the Alliance button/text is present; if not, loops to find and tap another close.
- Close button detection uses OCR for `X`, `x`, `Close`, `Skip`, plus cropped OCR passes on common ad regions (top-right 30%, bottom-center 15% height).
- Alliance detection uses Tesseract text search (`Alliance`) and optional template matching if a template is provided.

## Files
- `alliance_automation/adb.py` — wrappers for `adb`/`memuc`.
- `alliance_automation/image_utils.py` — screenshot capture and basic image helpers.
- `alliance_automation/ocr.py` — Tesseract-based OCR helpers.
- `alliance_automation/detect.py` — close-button and alliance detection.
- `alliance_automation/automation.py` — orchestration loop to start MEmu, launch app, dismiss ads, and capture logs.
- `main.py` — CLI entry point.

## Notes
- This rewrite aims to be cross-platform but still assumes MEmu/adb availability on Windows paths by default.
- Template matching uses OpenCV’s `matchTemplate` when a template image is available. Provide `--alliance-template` to improve detection.
