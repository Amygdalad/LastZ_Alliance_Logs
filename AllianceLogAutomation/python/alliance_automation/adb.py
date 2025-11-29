from __future__ import annotations

import subprocess
import time
from pathlib import Path
from typing import Iterable, List, Optional


class AdbError(RuntimeError):
    pass


def _run(cmd: List[str], timeout: Optional[int] = None) -> str:
    try:
        res = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout, check=False)
    except FileNotFoundError as exc:
        raise AdbError(f"Command not found: {cmd[0]}") from exc
    if res.returncode != 0:
        raise AdbError(f"Command failed ({res.returncode}): {' '.join(cmd)}\nSTDOUT:{res.stdout}\nSTDERR:{res.stderr}")
    return res.stdout


def adb_path(memu_root: Path) -> Path:
    return memu_root / "adb.exe"


def memuc_path(memu_root: Path) -> Path:
    return memu_root / "memuc.exe"


def run_adb(serial: str, memu_root: Path, args: Iterable[str], timeout: Optional[int] = None) -> str:
    cmd = [str(adb_path(memu_root)), "-s", serial, *map(str, args)]
    return _run(cmd, timeout=timeout)


def run_memuc(memu_root: Path, args: Iterable[str], timeout: Optional[int] = None) -> str:
    cmd = [str(memuc_path(memu_root)), *map(str, args)]
    return _run(cmd, timeout=timeout)


def wait_for_device(serial: str, memu_root: Path, timeout: int) -> None:
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            out = run_adb(serial, memu_root, ["get-state"], timeout=10)
            if "device" in out:
                return
        except AdbError:
            time.sleep(2)
            continue
        time.sleep(2)
    raise AdbError(f"Timed out waiting for device {serial}")


def start_instance(index: int, memu_root: Path, initial_wait: int, post_wait: int) -> str:
    out = run_memuc(memu_root, ["start", str(index)])
    serial = None
    for line in out.splitlines():
        if "adb:" in line and "Android Emulator" in line:
            parts = line.split()
            serial = parts[-1]
    # Fallback: construct expected serial pattern for MEmu
    if not serial:
        serial = f"127.0.0.1:{21500 + index}"
    time.sleep(initial_wait)
    wait_for_device(serial, memu_root, post_wait)
    return serial


def stop_instance(index: int, memu_root: Path) -> None:
    try:
        run_memuc(memu_root, ["stop", str(index)], timeout=30)
    except AdbError:
        # Best-effort stop; swallow errors to avoid aborting the run.
        pass


def unlock_device(serial: str, memu_root: Path) -> None:
    # Wake and unlock with simple swipe and menu key; best-effort.
    try:
        run_adb(serial, memu_root, ["shell", "input", "keyevent", "26"])
        time.sleep(0.5)
        run_adb(serial, memu_root, ["shell", "input", "swipe", "300", "1000", "300", "300"])
        time.sleep(0.5)
        run_adb(serial, memu_root, ["shell", "input", "keyevent", "82"])
    except AdbError:
        pass


def launch_app(serial: str, memu_root: Path, package: Optional[str]) -> None:
    if not package:
        return
    try:
        run_adb(
            serial,
            memu_root,
            ["shell", "monkey", "-p", package, "-c", "android.intent.category.LAUNCHER", "1"],
            timeout=15,
        )
    except AdbError:
        pass


def tap(serial: str, memu_root: Path, x: int, y: int, pause_ms: int = 0) -> None:
    try:
        run_adb(serial, memu_root, ["shell", "input", "tap", str(x), str(y)], timeout=5)
    finally:
        if pause_ms > 0:
            time.sleep(pause_ms / 1000)
