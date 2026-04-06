"""
Executes PowerShell scripts via 'pwsh' and streams output line by line.
Also writes a timestamped log file.
"""

import subprocess
import threading
import tempfile
import os
import sys
from datetime import datetime
from pathlib import Path


def _log_dir():
    """Return the logs directory, creating it if needed."""
    if sys.platform == "win32":
        base = Path(os.environ.get("APPDATA", Path.home())) / "ETD_Configurator" / "logs"
    else:
        base = Path.home() / "Library" / "Logs" / "ETD_Configurator"
    base.mkdir(parents=True, exist_ok=True)
    return base


def _pwsh_executable():
    """Return the full path to the PowerShell executable."""
    import shutil

    # Try PATH first (works when launched from terminal)
    found = shutil.which("pwsh")
    if found:
        return found

    # Common install locations on macOS (Homebrew Intel/ARM, direct installer)
    mac_paths = [
        "/opt/homebrew/bin/pwsh",           # Homebrew Apple Silicon
        "/usr/local/bin/pwsh",              # Homebrew Intel
        "/usr/local/microsoft/powershell/7/pwsh",  # Microsoft direct installer
    ]
    for path in mac_paths:
        if os.path.isfile(path):
            return path

    # Fallback — will raise FileNotFoundError with a clear message
    return "pwsh"


def run_script(ps_script, output_callback, done_callback):
    """
    Run a PowerShell script asynchronously.

    Args:
        ps_script      : str  - Full PowerShell script content
        output_callback: callable(str) - Called for each output line (thread-safe via after())
        done_callback  : callable(bool, str) - Called when done with (success, log_path)

    Returns:
        stop: callable() - Call to terminate the running process.
    """
    # Write script to a temporary file
    tmp = tempfile.NamedTemporaryFile(
        mode="w", suffix=".ps1", delete=False, encoding="utf-8"
    )
    tmp.write(ps_script)
    tmp.close()
    temp_path = tmp.name

    # Prepare log file
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = _log_dir() / f"etd_run_{timestamp}.log"

    _proc = [None]
    _stopped = [False]

    def stop():
        _stopped[0] = True
        if _proc[0] is not None:
            try:
                _proc[0].kill()
            except Exception:
                pass

    def _run():
        pwsh = _pwsh_executable()
        log_lines = []

        def record(line):
            log_lines.append(line)
            output_callback(line)

        record(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] ETD Configurator - Run started")
        record(f"Script: {temp_path}")
        record("=" * 60)

        success = False
        try:
            _proc[0] = subprocess.Popen(
                [pwsh, "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", temp_path],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
            )

            for line in _proc[0].stdout:
                record(line.rstrip())

            _proc[0].wait()
            success = _proc[0].returncode == 0 and not _stopped[0]

        except FileNotFoundError:
            record("ERROR: PowerShell (pwsh) not found.")
            record("Please install PowerShell 7: https://aka.ms/powershell")
        except Exception as e:
            if not _stopped[0]:
                record(f"ERROR: {e}")
        finally:
            try:
                os.unlink(temp_path)
            except OSError:
                pass

        if _stopped[0]:
            record("  Process stopped by user.")

        record("=" * 60)
        record(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Run finished - {'SUCCESS' if success else 'FAILED'}")

        # Write log file
        with open(log_path, "w", encoding="utf-8") as f:
            f.write("\n".join(log_lines))

        done_callback(success, str(log_path))

    thread = threading.Thread(target=_run, daemon=True)
    thread.start()
    return stop
