from __future__ import annotations

import time
from pathlib import Path
from typing import Optional

from . import adb
from .config import Settings
from .detect import Detection, alliance_visible, find_close_button
from .image_utils import capture_screenshot
from .ocr import ocr_text


class AllianceAutomation:
    def __init__(self, settings: Settings):
        self.settings = settings
        self.settings.ensure_dirs()
        self.screenshot_root = self.settings.report_root / "screenshots"
        self.ocr_root = self.settings.report_root / "ocr"

    def _screenshot_path(self, index: int, suffix: str) -> Path:
        stamp = time.strftime("%Y%m%d_%H%M%S")
        return self.screenshot_root / f"memu-{index}-{suffix}-{stamp}.png"

    def _report_path(self, index: int) -> Path:
        stamp = time.strftime("%Y%m%d_%H%M%S")
        return self.ocr_root / f"memu-{index}-{stamp}.txt"

    def _dismiss_ads(self, serial: str, index: int) -> Optional[Path]:
        start = time.time()
        shot_path = self._screenshot_path(index, "dismiss-start")
        capture_screenshot(serial, self.settings.memu_root, shot_path)

        while time.time() - start < self.settings.dismiss_timeout_seconds:
            # Alliance check first
            alli_hit = alliance_visible(
                shot_path,
                self.settings.tesseract_path,
                use_tesseract=self.settings.use_tesseract,
                alliance_template=self.settings.alliance_template if self.settings.use_alliance_template else None,
                template_threshold=self.settings.match_threshold,
            )
            if alli_hit.found:
                return shot_path

            # Close button search
            close_hit = find_close_button(
                shot_path, self.settings.tesseract_path, use_tesseract=self.settings.use_tesseract
            )
            if close_hit.found:
                adb.tap(serial, self.settings.memu_root, close_hit.x, close_hit.y, pause_ms=self.settings.tap_pause_ms)
                time.sleep(self.settings.tap_after_close_wait_seconds)
                # Re-screenshot and re-check for alliance
                shot_path = self._screenshot_path(index, "dismiss-after-close")
                capture_screenshot(serial, self.settings.memu_root, shot_path)
                alli_hit = alliance_visible(
                    shot_path,
                    self.settings.tesseract_path,
                    use_tesseract=True,
                    alliance_template=self.settings.alliance_template if self.settings.use_alliance_template else None,
                    template_threshold=self.settings.match_threshold,
                )
                if alli_hit.found:
                    return shot_path
                # Not found; continue loop to look for more close buttons
                continue

            time.sleep(self.settings.wait_retry_seconds)
            shot_path = self._screenshot_path(index, "dismiss-retry")
            capture_screenshot(serial, self.settings.memu_root, shot_path)

        return shot_path

    def process_index(self, index: int) -> None:
        serial = adb.start_instance(
            index, self.settings.memu_root, self.settings.initial_start_wait_seconds, self.settings.post_start_wait_seconds
        )
        try:
            adb.unlock_device(serial, self.settings.memu_root)
            if self.settings.target_package:
                adb.launch_app(serial, self.settings.memu_root, self.settings.target_package)
            time.sleep(self.settings.after_launch_delay_seconds)

            last_shot = self._dismiss_ads(serial, index)
            if not last_shot:
                last_shot = self._screenshot_path(index, "final")
                capture_screenshot(serial, self.settings.memu_root, last_shot)

            report_path = self._report_path(index)
            text = ""
            try:
                text = ocr_text(last_shot, self.settings.tesseract_path)
            except Exception as exc:  # noqa: BLE001
                text = f"(OCR failed: {exc})"

            lines = [
                f"Timestamp: {time.strftime('%Y-%m-%d %H:%M:%S')}",
                f"Index: {index}",
                f"ADB: {serial}",
                f"Screenshot: {last_shot}",
                "",
                "---- OCR ----",
                text,
            ]
            report_path.parent.mkdir(parents=True, exist_ok=True)
            report_path.write_text("\n".join(lines), encoding="utf-8")
        finally:
            if not self.settings.leave_running:
                adb.stop_instance(index, self.settings.memu_root)

    def run(self) -> None:
        for idx in self.settings.indexes:
            self.process_index(idx)
