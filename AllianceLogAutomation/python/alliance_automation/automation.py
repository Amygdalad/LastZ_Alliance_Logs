from __future__ import annotations

import atexit
import csv
import signal
import threading
import time
from pathlib import Path
from typing import Optional

from . import adb
from .config import Settings
from .detect import Detection, alliance_visible, find_close_button
from .image_utils import capture_screenshot
from .ocr import ocr_text


class _CleanupRegistry:
    """
    Tracks running MEmu instances and tears them down on signals/exit so the
    console window does not stay open with a 'core process stopped' dialog.
    """

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._active = []  # (index, serial, proc, memu_root, should_stop)
        self._hooks_installed = False

    def register(self, index: int, serial: str, proc, memu_root: Path, should_stop: bool = True) -> None:
        with self._lock:
            self._active.append((index, serial, proc, memu_root, should_stop))
        self._install_hooks()

    def unregister(self, index: int, serial: Optional[str]) -> None:
        if serial is None:
            return
        with self._lock:
            self._active = [row for row in self._active if not (row[0] == index and row[1] == serial)]

    def cleanup_all(self, reason: str | None = None) -> None:
        with self._lock:
            pending = list(self._active)
            self._active.clear()

        for index, serial, proc, memu_root, should_stop in reversed(pending):
            if not should_stop:
                continue
            try:
                adb.stop_instance(index, memu_root, serial=serial, proc=proc)
            except Exception as exc:  # noqa: BLE001
                print(f"[cleanup] Failed to stop instance {index} ({serial}) during {reason or 'exit'}: {exc}")

    def _install_hooks(self) -> None:
        if self._hooks_installed:
            return

        atexit.register(self.cleanup_all)

        for sig_name in ("SIGINT", "SIGTERM", "SIGBREAK"):
            sig = getattr(signal, sig_name, None)
            if sig is None:
                continue
            try:
                previous = signal.getsignal(sig)

                def handler(signum, frame, sig=sig, previous=previous):
                    print(f"Received {getattr(sig, 'name', signum)}. Cleaning up MEmu instances...")
                    self.cleanup_all(reason=f"signal {signum}")
                    if previous in (signal.SIG_IGN, None):
                        return
                    if previous in (signal.SIG_DFL, signal.default_int_handler):
                        raise KeyboardInterrupt
                    try:
                        previous(signum, frame)
                    except Exception:
                        pass

                signal.signal(sig, handler)
            except Exception:
                continue

        self._hooks_installed = True


_cleanup_registry = _CleanupRegistry()


class AllianceAutomation:
    def __init__(self, settings: Settings):
        self.settings = settings
        self.settings.ensure_dirs()
        self.screenshot_root = self.settings.report_root / "screenshots"
        self.ocr_root = self.settings.report_root / "ocr"
        self._cleanup = _cleanup_registry

    def _screenshot_path(self, index: int, suffix: str) -> Path:
        stamp = time.strftime("%Y%m%d_%H%M%S")
        return self.screenshot_root / f"memu-{index}-{suffix}-{stamp}.png"

    def _report_path(self, index: int) -> Path:
        stamp = time.strftime("%Y%m%d_%H%M%S")
        return self.ocr_root / f"memu-{index}-{stamp}.txt"

    def _csv_path(self, index: int) -> Path:
        stamp = time.strftime("%Y%m%d_%H%M%S")
        return self.ocr_root / f"memu-{index}-{stamp}.csv"

    def _check_alliance_visible(self, shot_path: Path) -> bool:
        template_path = self.settings.alliance_template if self.settings.use_alliance_template else None
        if template_path and not template_path.exists():
            print(f"Warning: Template not found at {template_path}")
            
        alli_hit = alliance_visible(
            shot_path,
            self.settings.tesseract_path,
            use_tesseract=self.settings.use_tesseract,
            alliance_template=template_path,
            template_threshold=self.settings.match_threshold,
        )
        if alli_hit.found:
            print(f"Alliance button found via {alli_hit.note} at {alli_hit.x},{alli_hit.y} (Score: {alli_hit.score})")
        return alli_hit.found

    def _dismiss_ads_and_check_alliance(self, serial: str, index: int) -> bool:
        """
        Clears pop-ups, checking if alliance button is visible between each click.
        Max pop-up clear is defined in settings (default 2).
        Returns True if Alliance button is visible at the end, False otherwise.
        """
        clears = 0
        
        # Loop with retries to handle loading time and popups
        max_retries = 20
        for i in range(max_retries):
            shot_path = self._screenshot_path(index, f"check-start-{i}")
            capture_screenshot(serial, self.settings.memu_root, shot_path)
            
            if self._check_alliance_visible(shot_path):
                return True
            
            # Check for close button
            close_hit = find_close_button(
                shot_path, self.settings.tesseract_path, use_tesseract=self.settings.use_tesseract
            )
            
            if close_hit.found and clears < self.settings.max_popup_clears:
                print(f"Found close button at {close_hit.x}, {close_hit.y}. Tapping...")
                adb.tap(serial, self.settings.memu_root, close_hit.x, close_hit.y, pause_ms=self.settings.tap_pause_ms)
                clears += 1
                time.sleep(self.settings.tap_after_close_wait_seconds)
                continue

            if close_hit.found:
                print("Max popup clears reached, but close button still found. Treating as stuck.")
            else:
                print(f"No close button found. Waiting... ({i+1}/{max_retries})")
            
            # Fallback strategies
            if i > 2:
                if i % 4 == 0:
                    print("Stuck? Trying Android Back button...")
                    adb.run_adb(serial, self.settings.memu_root, ["shell", "input", "keyevent", "4"])
                elif i % 4 == 2:
                    print("Stuck? Tapping top-center to dismiss...")
                    # Tap top center (assuming 540 width, tap at 270, 100)
                    adb.tap(serial, self.settings.memu_root, 270, 100, pause_ms=self.settings.tap_pause_ms)
            
            time.sleep(self.settings.wait_retry_seconds)

        return False

    def _navigate_to_logs(self, serial: str) -> None:
        print("Navigating to Alliance Logs...")
        # 1. Click Alliance
        adb.tap(serial, self.settings.memu_root, self.settings.alliance_tap_x, self.settings.alliance_tap_y, pause_ms=self.settings.tap_pause_ms)
        time.sleep(self.settings.alliance_to_logs_delay_seconds)
        
        # 2. Click Logs
        adb.tap(serial, self.settings.memu_root, self.settings.alliance_logs_tap_x, self.settings.alliance_logs_tap_y, pause_ms=self.settings.tap_pause_ms)
        time.sleep(self.settings.alliance_to_logs_delay_seconds) # Wait for logs to load

    def _parse_to_csv(self, text: str, csv_path: Path) -> None:
        """
        Parses the OCR text and writes to CSV.
        Assumes a simple format for now, or just dumps lines.
        """
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        
        with open(csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(["Line"]) # Header
            for line in lines:
                writer.writerow([line])

    def process_index(self, index: int) -> None:
        serial: Optional[str] = None
        proc = None
        serial, proc = adb.start_instance(
            index,
            self.settings.memu_root,
            self.settings.instance_root,
            self.settings.initial_start_wait_seconds,
            self.settings.post_start_wait_seconds,
        )
        self._cleanup.register(index, serial, proc, self.settings.memu_root, should_stop=not self.settings.leave_running)
        try:
            adb.unlock_device(serial, self.settings.memu_root)
            
            target = self.settings.target_package
            if not target:
                # Auto-discover
                pkgs = adb.get_installed_packages(serial, self.settings.memu_root)
                # Prefer com.readygo.barrel.gp
                if "com.readygo.barrel.gp" in pkgs:
                    target = "com.readygo.barrel.gp"
                else:
                    # Fallback to anything with 'z' or 'Z'
                    for p in pkgs:
                        if "z" in p.lower():
                            target = p
                            break
            
            if target:
                print(f"Launching {target}...")
                activity = None
                if target == "com.readygo.barrel.gp":
                    activity = "com.im30.aps.debug.UnityPlayerActivityCustom"
                adb.launch_app(serial, self.settings.memu_root, target, activity)
            else:
                print("No target package found to launch.")

            time.sleep(self.settings.after_launch_delay_seconds)

            # 3. Clear pop-ups & Check Alliance
            if self._dismiss_ads_and_check_alliance(serial, index):
                print("Alliance button found.")
                
                # 4. Click Alliance Logs
                self._navigate_to_logs(serial)
                
                # 5. Screenshot Alliance Logs
                logs_shot = self._screenshot_path(index, "logs")
                capture_screenshot(serial, self.settings.memu_root, logs_shot)
                
                # 6. OCR
                text = ""
                try:
                    text = ocr_text(logs_shot, self.settings.tesseract_path)
                except Exception as exc:  # noqa: BLE001
                    text = f"(OCR failed: {exc})"
                
                # Report Text
                report_path = self._report_path(index)
                lines = [
                    f"Timestamp: {time.strftime('%Y-%m-%d %H:%M:%S')}",
                    f"Index: {index}",
                    f"ADB: {serial}",
                    f"Screenshot: {logs_shot}",
                    "",
                    "---- OCR ----",
                    text,
                ]
                report_path.parent.mkdir(parents=True, exist_ok=True)
                report_path.write_text("\n".join(lines), encoding="utf-8")
                
                # 7. Parse to CSV
                csv_path = self._csv_path(index)
                self._parse_to_csv(text, csv_path)
                print(f"Report saved to {report_path}")
                print(f"CSV saved to {csv_path}")

            else:
                print("Could not find Alliance button after dismissing ads.")
                # Take a final screenshot for debugging
                fail_shot = self._screenshot_path(index, "fail")
                capture_screenshot(serial, self.settings.memu_root, fail_shot)

        finally:
            if not self.settings.leave_running and serial:
                try:
                    adb.stop_instance(index, self.settings.memu_root, serial=serial, proc=proc)
                finally:
                    self._cleanup.unregister(index, serial)
            else:
                self._cleanup.unregister(index, serial)

    def run(self) -> None:
        for idx in self.settings.indexes:
            self.process_index(idx)
