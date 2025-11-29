from __future__ import annotations

import argparse
from pathlib import Path

from .automation import AllianceAutomation
from .config import Settings


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Alliance Log Automation (Python)")
    p.add_argument("--indexes", nargs="+", type=int, default=[59, 57, 173], help="MEmu instance indexes")
    p.add_argument("--memu-root", type=Path, default=Path(r"Z:\Program Files\Microvirt\MEmu"), help="Path to MEmu root containing adb.exe/memuc.exe")
    p.add_argument("--instance-root", type=Path, default=Path(r"Z:\Program Files\Microvirt\MEmu\MemuHyperv VMs"), help="Path to instance root (for reference)")
    p.add_argument("--report-root", type=Path, default=Path("reports"), help="Root output directory for screenshots/OCR")
    p.add_argument("--target-package", type=str, default=None, help="Package name to launch (monkey).")
    p.add_argument("--after-launch-delay", type=int, default=8, help="Delay after launching app before first screenshot.")
    p.add_argument("--device-boot-timeout", type=int, default=180, help="Device boot timeout seconds.")
    p.add_argument("--initial-start-wait", type=int, default=45, help="Initial wait after memu start.")
    p.add_argument("--post-start-wait", type=int, default=10, help="Extra wait before actions.")
    p.add_argument("--dismiss-timeout", type=int, default=60, help="Seconds to loop dismissing ads/close buttons.")
    p.add_argument("--tap-after-close-wait", type=int, default=5, help="Seconds to wait after tapping a close before re-checking alliance.")
    p.add_argument("--wait-retry-seconds", type=int, default=2, help="Delay between no-close retries.")
    p.add_argument("--tesseract-path", type=Path, default=Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"))
    p.add_argument("--alliance-template", type=Path, default=None, help="Template image for alliance button/icon (optional).")
    p.add_argument("--no-template", action="store_true", help="Disable template matching for alliance.")
    p.add_argument("--leave-running", action="store_true", help="Leave MEmu instances running after capture.")
    return p


def main() -> None:
    args = build_parser().parse_args()
    settings = Settings(
        indexes=args.indexes,
        instance_root=args.instance_root,
        memu_root=args.memu_root,
        report_root=args.report_root,
        target_package=args.target_package,
        after_launch_delay_seconds=args.after_launch_delay,
        device_boot_timeout_seconds=args.device_boot_timeout,
        initial_start_wait_seconds=args.initial_start_wait,
        post_start_wait_seconds=args.post_start_wait,
        dismiss_timeout_seconds=args.dismiss_timeout,
        tap_after_close_wait_seconds=args.tap_after_close_wait,
        wait_retry_seconds=args.wait_retry_seconds,
        tesseract_path=args.tesseract_path,
        alliance_template=args.alliance_template,
        use_alliance_template=not args.no_template,
        leave_running=args.leave_running,
    )
    automation = AllianceAutomation(settings)
    automation.run()


if __name__ == "__main__":
    main()
