from __future__ import annotations

import pathlib
from dataclasses import dataclass, field
from typing import List, Optional


@dataclass
class Settings:
    indexes: List[int] = field(default_factory=lambda: [59, 57, 173])
    instance_root: pathlib.Path = pathlib.Path(r"Z:\Program Files\Microvirt\MEmu\MemuHyperv VMs")
    memu_root: pathlib.Path = pathlib.Path(r"Z:\Program Files\Microvirt\MEmu")
    report_root: pathlib.Path = pathlib.Path("reports")
    target_package: Optional[str] = None
    after_launch_delay_seconds: int = 8
    device_boot_timeout_seconds: int = 180
    initial_start_wait_seconds: int = 45
    post_start_wait_seconds: int = 10
    tap_interval_seconds: int = 5
    max_tap_wait_seconds: int = 180
    tap_pause_ms: int = 700
    search_region_fraction: float = 0.6
    color_tolerance: int = 80
    match_threshold: float = 0.78
    template_sample_step: int = 2
    alliance_template: Optional[pathlib.Path] = None
    tesseract_path: pathlib.Path = pathlib.Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe")
    use_tesseract: bool = True
    use_alliance_template: bool = True
    dismiss_timeout_seconds: int = 60
    tap_after_close_wait_seconds: int = 5
    wait_retry_seconds: int = 2
    leave_running: bool = False

    def ensure_dirs(self) -> None:
        self.report_root.mkdir(parents=True, exist_ok=True)
        (self.report_root / "screenshots").mkdir(parents=True, exist_ok=True)
        (self.report_root / "ocr").mkdir(parents=True, exist_ok=True)
