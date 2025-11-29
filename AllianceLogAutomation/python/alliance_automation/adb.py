from __future__ import annotations

import subprocess
import time
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Iterable, List, Optional, Tuple


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


def memu_console_path(memu_root: Path) -> Path:
    return memu_root / "MEmuConsole.exe"


def run_adb(serial: str, memu_root: Path, args: Iterable[str], timeout: Optional[int] = None) -> str:
    cmd = [str(adb_path(memu_root)), "-s", serial, *map(str, args)]
    return _run(cmd, timeout=timeout)


def run_memuc(memu_root: Path, args: Iterable[str], timeout: Optional[int] = None) -> str:
    cmd = [str(memuc_path(memu_root)), *map(str, args)]
    return _run(cmd, timeout=timeout)


def get_adb_port(index: int, instance_root: Path) -> int:
    # e.g. MEmu_173/MEmu_173.memu
    name = f"MEmu_{index}" if index > 0 else "MEmu"
    memu_file = instance_root / name / f"{name}.memu"
    if not memu_file.exists():
        # Try flat structure
        memu_file = instance_root / f"{name}.memu"
    
    if not memu_file.exists():
        raise FileNotFoundError(f"Could not find .memu file for index {index} at {memu_file}")

    try:
        tree = ET.parse(memu_file)
        root = tree.getroot()
        # Namespace handling might be needed, but ElementTree often handles it if we ignore it or use wildcards
        # The PS script uses: //m:Forwarding[@name='ADB']
        # Let's try to find Forwarding with name='ADB'
        # The namespace is http://www.memuhyperv.org/
        ns = {'m': 'http://www.memuhyperv.org/'}
        for forwarding in root.findall(".//m:Forwarding", ns):
            if forwarding.get("name") == "ADB":
                return int(forwarding.get("hostport"))
    except Exception as e:
        print(f"Warning: Failed to parse .memu file: {e}")
    
    # Fallback
    return 21503 + index * 10


def wait_for_device(serial: str, memu_root: Path, timeout: int) -> None:
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            out = run_adb(serial, memu_root, ["get-state"], timeout=10)
            if "device" in out:
                return
            print(f"Device state: {out.strip()}")
        except AdbError as e:
            # print(f"Waiting for device {serial}... ({e})")
            # Try to connect if not connected
            try:
                run_adb(serial, memu_root, ["connect", serial], timeout=5)
            except AdbError:
                pass
            time.sleep(2)
            continue
        time.sleep(2)
    raise AdbError(f"Timed out waiting for device {serial}")


from typing import Iterable, List, Optional, Tuple

# ... (imports)

def start_instance(index: int, memu_root: Path, instance_root: Path, initial_wait: int, post_wait: int) -> Tuple[str, subprocess.Popen]:
    # Use MEmuConsole.exe to start
    name = f"MEmu_{index}" if index > 0 else "MEmu"
    console_exe = memu_console_path(memu_root)
    
    print(f"Starting {name} via {console_exe}...")
    proc: Optional[subprocess.Popen] = None
    serial: Optional[str] = None
    try:
        proc = subprocess.Popen([str(console_exe), name], cwd=str(memu_root))
        
        # Get port
        port = get_adb_port(index, instance_root)
        serial = f"127.0.0.1:{port}"
        print(f"Expecting serial: {serial}")
        
        time.sleep(initial_wait)
        wait_for_device(serial, memu_root, post_wait)
        return serial, proc
    except BaseException:
        # Ensure we do not leave the console open if startup fails or is interrupted.
        try:
            stop_instance(index, memu_root, serial=serial, proc=proc)
        except Exception as cleanup_exc:  # noqa: BLE001
            print(f"Failed to clean up {name} after start error: {cleanup_exc}")
        raise


def stop_instance(index: int, memu_root: Path, serial: Optional[str] = None, proc: Optional[subprocess.Popen] = None) -> None:
    print(f"Stopping instance {index}...")
    
    # 1. Kill the console window first (if we own it)
    # This prevents the "Core process ended" popup if the VM dies later.
    if proc:
        if proc.poll() is None:
            print(f"Terminating console process {proc.pid}...")
            proc.terminate()
            try:
                proc.wait(timeout=5)
            except subprocess.TimeoutExpired:
                print("Console did not terminate. Killing...")
                proc.kill()
                proc.wait()
            print("Console process killed.")
        else:
            print("Console process already exited.")

    # 2. Try memuc stop (to stop the engine gracefully if still running)
    try:
        run_memuc(memu_root, ["stop", str(index)], timeout=30)
    except AdbError:
        print("memuc stop failed, continuing...")

    if not serial:
        return

    if not serial:
        return

    # 2. Wait and check if offline
    time.sleep(2)
    if _is_device_online(serial, memu_root):
        # 3. Force kill via ADB
        print(f"Instance {index} still online. Attempting 'adb emu kill'...")
        try:
            run_adb(serial, memu_root, ["emu", "kill"], timeout=10)
        except AdbError:
            pass
        time.sleep(5)
        if _is_device_online(serial, memu_root):
            print(f"Instance {index} still running. Force killing processes...")
        else:
            print("Instance stopped via adb emu kill. Cleaning up console...")
    else:
        print("Instance stopped successfully. Cleaning up console...")

    # 4. Force kill Windows processes (Cleanup) even if offline to close console/UI
    _kill_memu_processes(index, memu_root)


def _is_device_online(serial: str, memu_root: Path) -> bool:
    try:
        cmd = [str(adb_path(memu_root)), "devices"]
        out = _run(cmd, timeout=10)
        # Check if serial is listed and not offline/unauthorized
        # Output format: "serial\tdevice"
        for line in out.splitlines():
            parts = line.split()
            if len(parts) >= 2 and parts[0] == serial and parts[1] == "device":
                return True
        return False
    except AdbError:
        return False


def _kill_memu_processes(index: int, memu_root: Path) -> None:
    """
    Kills MEmuHeadless.exe, MEmuHyper.exe, and MEmuConsole.exe for the specific instance using PowerShell.
    """
    name = f"MEmu_{index}" if index > 0 else "MEmu"
    
    # Use PowerShell to find and kill processes with matching command line
    # We use a regex to ensure we match the instance name as an argument, 
    # not as part of the file path (e.g. ...\MEmu\...)
    # Regex explanation:
    # (?<![\\/])  : Negative lookbehind to ensure not preceded by \ or / (path separators)
    # {name}      : The instance name (e.g. MEmu or MEmu_173)
    # (?=\s|$|")  : Lookahead to ensure followed by space, end of string, or quote
    
    ps_command = (
        f"Get-CimInstance Win32_Process -Filter \"name='MEmu.exe' OR name='MEmuHeadless.exe' OR name='MEmuConsole.exe' OR name='MEmuHyper.exe'\" | "
        f"Where-Object {{ $_.CommandLine -match '(?<![\\\\/]){name}(?=\\s|$|\")' }} | "
        f"ForEach-Object {{ Stop-Process -Id $_.ProcessId -Force }}"
    )
    
    try:
        subprocess.run(["powershell", "-Command", ps_command], capture_output=True, check=False)
    except Exception as e:
        print(f"Failed to kill processes via PowerShell: {e}")

    # 5. Final cleanup with memuc stop to update service state
    # If we force-killed the process, MEmu service might still think it's running.
    # Running stop again usually clears this state.
    try:
        run_memuc(memu_root, ["stop", str(index)], timeout=10)
    except AdbError:
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


def launch_app(serial: str, memu_root: Path, package: str, activity: Optional[str] = None) -> None:
    if not package:
        return
    try:
        if activity:
            # Use am start
            component = f"{package}/{activity}"
            run_adb(serial, memu_root, ["shell", "am", "start", "-n", component], timeout=15)
        else:
            run_adb(
                serial,
                memu_root,
                ["shell", "monkey", "-p", package, "-c", "android.intent.category.LAUNCHER", "1"],
                timeout=15,
            )
    except AdbError:
        pass


def get_installed_packages(serial: str, memu_root: Path) -> List[str]:
    try:
        out = run_adb(serial, memu_root, ["shell", "pm", "list", "packages", "-3"], timeout=10)
        packages = []
        for line in out.splitlines():
            if line.startswith("package:"):
                packages.append(line.split(":", 1)[1].strip())
        return packages
    except AdbError:
        return []


def tap(serial: str, memu_root: Path, x: int, y: int, pause_ms: int = 0) -> None:
    try:
        run_adb(serial, memu_root, ["shell", "input", "tap", str(x), str(y)], timeout=5)
    finally:
        if pause_ms > 0:
            time.sleep(pause_ms / 1000)
