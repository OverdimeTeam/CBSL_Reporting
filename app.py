import openpyxl
from pathlib import Path
import os
import shutil
import subprocess
import sys
import threading
from datetime import datetime, timedelta
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Dict, Tuple, List

from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_file
from threading import Lock
from collections import deque
# from waitress import serve

app = Flask(__name__)

# Suppress Waitress task queue warnings (harmless, just indicates polling during long tasks)
import logging
logging.getLogger('waitress.queue').setLevel(logging.ERROR)


def kill_excel_instances():
    """Kill any running Excel instances to prevent COM conflicts."""
    try:
        print("Checking for running Excel instances...")
        result = subprocess.run(
            ['tasklist', '/FI', 'IMAGENAME eq EXCEL.EXE'],
            capture_output=True,
            text=True,
            timeout=5
        )

        if 'EXCEL.EXE' in result.stdout:
            print("Found running Excel instances. Killing them...")
            subprocess.run(
                ['taskkill', '/F', '/IM', 'EXCEL.EXE'],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=5
            )
            import time
            time.sleep(2)
            print("Excel instances killed successfully")
        else:
            print("No Excel instances found")
    except Exception as e:
        print(f"Error killing Excel instances: {e}")


def _append_run_history(
    frequency: str,
    picked_date: str,
    scripts: list[str],
    checklist: list[dict],
    output_dir: Path | None
) -> None:
    LOGS_DIR.mkdir(parents=True, exist_ok=True)
    
    if HISTORY_XLSX.exists():
        wb = openpyxl.load_workbook(HISTORY_XLSX)
    else:
        wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "History"

    # Add header if file is new
    if ws.max_row == 1 and ws.cell(1, 1).value is None:
        ws.append(["timestamp", "frequency", "picked_date", "reports", "statuses", "output_folder"])

    from datetime import datetime as _dt
    ts = _dt.now().strftime("%Y-%m-%d %H:%M:%S")

    reports = ", ".join(Path(s).stem for s in scripts)
    statuses = ", ".join(f"{it['name']}={it['status']}" for it in checklist)
    out_name = output_dir.name if output_dir else ""

    ws.append([ts, frequency, picked_date, reports, statuses, out_name])

    # Keep last 8 records + header
    while ws.max_row > 9:
        ws.delete_rows(2)

    wb.save(HISTORY_XLSX)
    wb.close()



BASE_DIR = Path(__file__).resolve().parent
OUTPUTS_DIR = BASE_DIR / "outputs"
WORKING_DIR = BASE_DIR / "working"
REPORT_AUTOMATIONS_DIR = BASE_DIR / "report_automations"
LOGS_DIR = BASE_DIR / "logs"
HISTORY_XLSX = LOGS_DIR / "run_history.xlsx"
_status_lock = Lock()
_status: dict[str, str | bool] = {
    "running": False,
    "stage": "idle",
    "frequency": "",
    "date": "",
    "message": "",
    "eta_total_secs": 0,
    "eta_started_iso": "",
}

_log_lock = Lock()
_log_buffer: deque[str] = deque(maxlen=2000)
_proc_lock = Lock()
_current_proc: subprocess.Popen | None = None
_stop_requested: bool = False
_running_processes: dict[str, subprocess.Popen] = {}  # Track individual processes by script name
_thread_pool: ThreadPoolExecutor | None = None
_running_threads: dict[str, threading.Thread] = {}  # Track running threads by name

# Create timestamped log file for each run
from datetime import datetime as _dt
_timestamp_str = _dt.now().strftime('%Y%m%d_%H%M%S')
_run_log_dir = LOGS_DIR / f"run_{_timestamp_str}"
_run_log_dir.mkdir(parents=True, exist_ok=True)
RUN_LOG_FILE = _run_log_dir / "web_run.log"
_status_feed: deque[str] = deque(maxlen=200)
_report_checklist: list[dict] = []
_report_messages: dict[str, deque[str]] = {}
_last_output_dir: Path | None = None


def _append_gui(line: str) -> None:
    with _log_lock:
        _log_buffer.append(line)


def _append_file(line: str) -> None:
    try:
        with RUN_LOG_FILE.open("a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        # Avoid UI crashes if log file cannot be written
        pass


def _timestamp() -> str:
    return datetime.now().strftime("%H:%M:%S")


def _is_important(msg: str) -> bool:
    text = msg.lower()
    keywords = [
        "starting", "ensuring", "copying", "running", "saving",
        "completed", "done", "stopping", "stop requested",
        "error", "failed", "warning", "destination already exists",
    ]
    return any(k in text for k in keywords)


def _emit_log(message: str, *, important: bool = False, skip_web_log: bool = False) -> None:
    """
    Emit log message.

    Args:
        message: Log message text
        important: If True, always show in GUI feed
        skip_web_log: If True, skip writing to web_run.log (for verbose report output)
    """
    line = f"[{_timestamp()}] {message}"

    # Only write to web_run.log if not skipped (skip verbose report output)
    if not skip_web_log:
        _append_file(line)

    # Show in GUI if important or matches keywords
    if important or _is_important(message):
        _append_gui(line)


def _detect_log_file_creation(line: str) -> None:
    """Detect when a descriptive log file is created and emit status update."""
    if "run log file:" in line.lower() or "run log saved to:" in line.lower():
        # Extract log file path from the line
        import re
        match = re.search(r'([A-Za-z0-9_\-\.]+\.log)', line)
        if match:
            log_file = match.group(1)
            _emit_log(f"Created descriptive log file: {log_file}", important=True)


def _log(message: str) -> None:
    # Backward compatible: treat direct _log calls as important UI messages
    _emit_log(message, important=True)


def _clear_logs() -> None:
    with _log_lock:
        _log_buffer.clear()


def _initialize_new_run() -> None:
    """Initialize a new run with fresh timestamp and log directory."""
    global _timestamp_str, _run_log_dir, RUN_LOG_FILE, _thread_pool
    _timestamp_str = _dt.now().strftime('%Y%m%d_%H%M%S')
    _run_log_dir = LOGS_DIR / f"run_{_timestamp_str}"
    _run_log_dir.mkdir(parents=True, exist_ok=True)
    RUN_LOG_FILE = _run_log_dir / "web_run.log"
    _thread_pool = ThreadPoolExecutor(max_workers=4)  # C6..C2, IA, NBD-MF-23-C4, and GA-IS
    _emit_log(f"Starting new run: {_timestamp_str}", important=True)


def _request_stop() -> None:
    """Request to stop all processes and threads."""
    global _stop_requested, _thread_pool
    _stop_requested = True
    
    with _proc_lock:
        # Stop all running processes
        for report_name, proc in _running_processes.items():
            try:
                proc.terminate()
                _emit_log(f"Stopped {report_name}", important=True)
            except Exception:
                pass
        _running_processes.clear()
        
        # Also stop the current process if it exists
        if _current_proc:
            try:
                _current_proc.terminate()
            except Exception:
                pass
    
    # Shutdown thread pool
    if _thread_pool:
        try:
            _thread_pool.shutdown(wait=False)
        except Exception:
            pass
        _thread_pool = None


def _stop_individual_script(report_name: str) -> bool:
    """Stop a specific script by name. Returns True if stopped, False if not found."""
    with _proc_lock:
        if report_name in _running_processes:
            try:
                _running_processes[report_name].terminate()
                del _running_processes[report_name]
                _emit_log(f"Stopped {report_name}", important=True)
                return True
            except Exception:
                return False
    return False


def _set_status(**kwargs) -> None:
    with _status_lock:
        prev = _status.copy()
        _status.update(kwargs)
        # Record status changes in a compact feed for the UI
        stage = _status.get("stage", "") or ""
        msg = _status.get("message", "") or ""
        freq = _status.get("frequency", "") or ""
        date = _status.get("date", "") or ""
        parts = []
        if stage:
            parts.append(stage.capitalize())
        if freq:
            parts.append(f"{freq}")
        if date:
            parts.append(f"{date}")
        if msg:
            parts.append(msg)
        summary = " | ".join(parts)
        if summary:
            ts_line = f"[{_timestamp()}] {summary}"
            _status_feed.append(ts_line)


def _reset_status() -> None:
    _set_status(running=False, stage="idle", message="", frequency="", date="", eta_total_secs=0, eta_started_iso="")
    # Do not clear report checklist here so UI can still show last results until next start


def _init_report_checklist(script_names: list[str]) -> None:
    global _report_checklist
    _report_checklist = [{"name": name.replace(".py", ""), "status": "pending"} for name in script_names]
    # Initialize per-report message buffers
    global _report_messages
    _report_messages = {name.replace(".py", ""): deque(maxlen=200) for name in script_names}


def _ensure_reports_in_checklist(script_names: list[str]) -> None:
    """Ensure provided scripts exist in the checklist and message buffers without resetting existing entries."""
    existing_names = {it["name"] for it in _report_checklist}
    for name in script_names:
        stem = name.replace(".py", "")
        if stem not in existing_names:
            _report_checklist.append({"name": stem, "status": "pending"})
            _report_messages.setdefault(stem, deque(maxlen=200))

def _ensure_checklist_item(name: str) -> None:
    stem = name.replace(".py", "")
    exists = any(it.get("name") == stem for it in _report_checklist)
    if not exists:
        _report_checklist.append({"name": stem, "status": "pending"})
    if stem not in _report_messages:
        _report_messages[stem] = deque(maxlen=200)


def _set_report_status(script_name: str, status: str) -> None:
    """Set status by script name instead of index"""
    for item in _report_checklist:
        if item["name"] == script_name.replace(".py", ""):
            item["status"] = status
            break


def _append_report_message(report_stem: str, message: str) -> None:
    # Filter and simplify: include only notable lines, trim length
    txt = message.strip()
    lower = txt.lower()
    include = (
        " - info - " in lower
        or " - warning - " in lower
        or " - error - " in lower
        or "running" in lower
        or "opened workbook" in lower
        or "saved" in lower
        or "completed" in lower
        or "failed" in lower
    )
    if not include:
        return
    # Remove noisy prefixes if present
    for marker in [" - info - ", " - warning - ", " - error - "]:
        idx = lower.find(marker)
        if idx != -1:
            txt = txt[idx + len(marker):]
            break
    if len(txt) > 180:
        txt = txt[:177] + "..."
    ts = _timestamp()
    line = f"[{ts}] {txt}"
    buf = _report_messages.get(report_stem)
    if buf is not None:
        buf.append(line)



def ensure_directories_exist() -> None:
    for sub in ["weekly", "monthly", "quarterly", "annually"]:
        target = OUTPUTS_DIR / sub
        target.mkdir(parents=True, exist_ok=True)
        completed_file = target / "completed_dates.txt"
        if not completed_file.exists():
            completed_file.write_text("")
    for sub in ["weekly", "monthly", "quarterly", "annually"]:
        (WORKING_DIR / sub).mkdir(parents=True, exist_ok=True)
    LOGS_DIR.mkdir(parents=True, exist_ok=True)


def _estimate_script_duration_seconds(script_name: str) -> int:
    """Estimate duration from past descriptive logs. Defaults to 600 seconds."""
    try:
        logs_dir = REPORT_AUTOMATIONS_DIR / "logs"
        if not logs_dir.exists():
            return 600
        import re
        import statistics
        durations: list[float] = []
        # Filter logs for this specific script by filename prefix
        prefix = Path(script_name).stem + "_"
        candidates = [p for p in sorted(logs_dir.glob("*.log")) if p.name.startswith(prefix)]
        for p in candidates[-20:]:
            try:
                txt = p.read_text(encoding="utf-8", errors="ignore")
            except Exception:
                continue
            # Heuristic: look for lines like "Run log saved to:" and first INFO line timestamp
            lines = txt.splitlines()
            if not lines:
                continue
            # Parse timestamps at line starts: 2025-09-17 22:11:47,123 - INFO - ...
            ts_pattern = re.compile(r"^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})")
            first_ts = None
            last_ts = None
            for ln in lines:
                m = ts_pattern.match(ln)
                if m:
                    from datetime import datetime as _dt
                    t = _dt.strptime(m.group(1), "%Y-%m-%d %H:%M:%S")
                    if first_ts is None:
                        first_ts = t
                    last_ts = t
            if first_ts and last_ts and last_ts > first_ts:
                durations.append((last_ts - first_ts).total_seconds())
        if durations:
            return int(statistics.median(durations))
    except Exception:
        pass
    return 600


def list_report_names() -> list[str]:
    # No longer used for creating placeholders; leaving for possible future use.
    if not REPORT_AUTOMATIONS_DIR.exists():
        return []
    names: list[str] = []
    for item in REPORT_AUTOMATIONS_DIR.iterdir():
        if item.is_file() and item.suffix == ".py" and not item.name.startswith("__"):
            names.append(item.stem)
    names.sort()
    return names


def parse_ui_date(date_str: str) -> datetime:
    """Accepts either yyyy-mm-dd (HTML date input) or mm/dd/yyyy (legacy)."""
    date_str = date_str.strip()
    for fmt in ("%Y-%m-%d", "%m/%d/%Y"):
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    raise ValueError(f"Unsupported date format: {date_str}")


def ui_date_to_folder_name(ui_date: str) -> str:
    dt = parse_ui_date(ui_date)
    return dt.strftime("%m-%d-%Y")


def ui_date_to_display(ui_date: str) -> str:
    dt = parse_ui_date(ui_date)
    return dt.strftime("%m/%d/%Y")


def append_completed_date(frequency: str, ui_date: str) -> None:
    completed_path = OUTPUTS_DIR / frequency / "completed_dates.txt"
    with completed_path.open("a", encoding="utf-8") as f:
        f.write(ui_date_to_display(ui_date) + "\n")


def get_last_completed_date(frequency: str) -> str | None:
    completed_path = OUTPUTS_DIR / frequency / "completed_dates.txt"
    if not completed_path.exists():
        return None
    lines = [ln.strip() for ln in completed_path.read_text(encoding="utf-8").splitlines() if ln.strip()]
    if not lines:
        return None
    return lines[-1]


def find_latest_versioned_folder(parent: Path, base_name: str) -> Path | None:
    candidates: list[tuple[int, Path]] = []
    for child in parent.iterdir():
        if not child.is_dir():
            continue
        name = child.name
        if name == base_name:
            candidates.append((0, child))
        elif name.startswith(base_name + "(") and name.endswith(")"):
            middle = name[len(base_name) + 1 : -1]
            if middle.isdigit():
                candidates.append((int(middle), child))
    if not candidates:
        return None
    candidates.sort(key=lambda t: t[0])
    return candidates[-1][1]


def create_reports_structure(target_dir: Path) -> None:
    # Intentionally no-op: no python files or report subfolders are created inside date folders.
    return


def copy_latest_weekly_to_working() -> tuple[bool, str]:
    last_ui_date = get_last_completed_date("weekly")
    if not last_ui_date:
        return False, "No completed weekly date found."
    base_name = ui_date_to_folder_name(last_ui_date)
    latest_folder = find_latest_versioned_folder(OUTPUTS_DIR / "weekly", base_name)
    if latest_folder is None:
        return False, f"No folder found for {last_ui_date}."
    dest = WORKING_DIR / "weekly" / latest_folder.name
    if dest.exists():
        shutil.rmtree(dest)
    shutil.copytree(latest_folder, dest)
    return True, f"Copied {latest_folder.name} to working/weekly."


def copy_latest_for_date_to_working(frequency: str, ui_date: str, *, clean_working: bool = True) -> tuple[bool, str]:
    """Copy latest versioned folder for the picked date into working/<frequency>.

    If clean_working is True, attempt to clear existing items under working/<frequency> first.
    """
    base_name = ui_date_to_folder_name(ui_date)
    parent = OUTPUTS_DIR / frequency
    if not parent.exists():
        return False, f"Source base folder not found: {parent}"
    latest_folder = find_latest_versioned_folder(parent, base_name)
    if latest_folder is None:
        try:
            existing = ", ".join(sorted([p.name for p in parent.iterdir() if p.is_dir()])) or "<none>"
        except Exception:
            existing = "<unknown>"
        return False, (
            f"No folder found for {ui_date_to_display(ui_date)} (expected '{base_name}' or versioned). "
            f"Existing under {parent.name}: {existing}"
        )
    working_freq_dir = WORKING_DIR / frequency
    if clean_working:
        try:
            if working_freq_dir.exists():
                for child in working_freq_dir.iterdir():
                    if child.is_dir():
                        shutil.rmtree(child)
                    else:
                        try:
                            child.unlink()
                        except Exception:
                            pass
            else:
                working_freq_dir.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            return False, f"Failed to clean working/{frequency}: {exc}"
    else:
        # Ensure dir exists, then remove existing subfolder(s) only
        working_freq_dir.mkdir(parents=True, exist_ok=True)
        try:
            for child in working_freq_dir.iterdir():
                if child.is_dir():
                    shutil.rmtree(child, ignore_errors=True)
        except Exception:
            # Best-effort; continue
            pass

    dest = working_freq_dir / latest_folder.name
    try:
        if dest.exists():
            shutil.rmtree(dest, ignore_errors=True)
        shutil.copytree(latest_folder, dest)
    except Exception as exc:
        return False, f"Failed to copy '{latest_folder.name}' to '{dest}': {exc}"
    return True, f"Copied {latest_folder.name} to working/{frequency}."


def copy_from_last_completed_to_working(frequency: str) -> tuple[bool, str]:
    last_ui_date = get_last_completed_date(frequency)
    if not last_ui_date:
        return False, f"No completed {frequency} date found."
    return copy_latest_for_date_to_working(frequency, last_ui_date)


def _run_script_direct(script_name: str, date_folder: Path, report_month: str, report_year: str) -> tuple[bool, str]:
    """Run a script directly in the current thread using importlib (for COM compatibility)."""
    import importlib.util
    import sys

    script_path = REPORT_AUTOMATIONS_DIR / script_name
    if not script_path.exists():
        return False, f"Script not found: {script_path}"

    _ensure_checklist_item(script_name)
    _set_report_status(script_name, "running")

    # Create individual log file
    script_log_file = _run_log_dir / f"{script_name.replace('.py', '')}.log"

    # Set up environment
    original_dir = os.getcwd()
    original_argv = sys.argv.copy()

    try:
        # Determine working directory
        desired_working = date_folder
        if script_name in ["NBD_MF_23_IA.py", "NBD_MF_23_C1C2.py"]:
            ia_folder = date_folder / "NBD_MF_23_IA"
            if ia_folder.exists():
                desired_working = ia_folder
        elif script_name == "NBD_MF_10_GA_IS.py":
            ga_is_folder = date_folder / "NBD-MF-10-GA & NBD-MF-11-IS"
            if ga_is_folder.exists():
                desired_working = ga_is_folder

        # Change to working directory
        os.chdir(str(desired_working))
        _emit_log(f"Running {script_name} in {desired_working}", important=True)

        # Set up sys.argv for argparse
        sys.argv = [
            str(script_path),
            "--working-dir", str(desired_working),
            "--month", report_month,
            "--year", report_year
        ]

        # Set log directory environment variable
        os.environ['REPORT_LOG_DIR'] = str(_run_log_dir)

        # Redirect stdout/stderr to capture logs
        import io
        from contextlib import redirect_stdout, redirect_stderr

        log_buffer = io.StringIO()

        with redirect_stdout(log_buffer), redirect_stderr(log_buffer):
            # Load and execute the script module
            spec = importlib.util.spec_from_file_location(f"script_{script_name.replace('.', '_')}", str(script_path))
            module = importlib.util.module_from_spec(spec)

            # Execute the module (this runs the script)
            spec.loader.exec_module(module)

        # Get captured output
        output = log_buffer.getvalue()

        # Write to log file and emit
        with script_log_file.open('w', encoding='utf-8') as f:
            f.write(output)

        for line in output.split('\n'):
            if line.strip():
                _emit_log(line, important=False, skip_web_log=True, script=script_name)

        _set_report_status(script_name, "completed")
        return True, "Completed"

    except Exception as e:
        error_msg = f"Error: {str(e)}"
        _emit_log(error_msg, important=True)
        import traceback
        _emit_log(traceback.format_exc(), important=False, skip_web_log=True)
        _set_report_status(script_name, "failed")
        return False, error_msg

    finally:
        # Restore original state
        os.chdir(original_dir)
        sys.argv = original_argv


def _run_single_script(script_name: str, date_folder: Path, report_month: str, report_year: str) -> tuple[bool, str]:
    """Run a single report script and return (success, message)."""
    script_path = REPORT_AUTOMATIONS_DIR / script_name
    if not script_path.exists():
        msg = f"Script not found: {script_path}; skipping"
        _emit_log(msg, important=True)
        _set_report_status(script_name, "skipped")
        return True, msg
    
    try:
        # Special handling for C4 and C6 scripts - run them directly without --working-dir
        if script_name in ["NBD_MF_20_C4.py", "NBD_MF_20_C6.py", "NBD_MF_20_C5.py", "NBD_MF_20_C2.py", "NBD_MF_20_C3_report.py"]:
            # For C6, C5, C4, C2, and C3, run directly with python command
            cmd = [
                "python", 
                script_name,
                "--month",
                report_month,
                "--year",
                report_year,
            ]
            # Set working directory to the report_automations folder for these scripts
            run_cwd = str(REPORT_AUTOMATIONS_DIR)
        else:
            # For other scripts like IA, use the full path and working directory
            # Special case: GA_IS uses a different folder name
            if script_name == "NBD_MF_10_GA_IS.py":
                desired_working = date_folder / "NBD-MF-10-GA & NBD-MF-11-IS"
                if not desired_working.exists():
                    desired_working = date_folder
            else:
                desired_working = date_folder / script_name.replace(".py", "")
                if not desired_working.exists():
                    desired_working = date_folder
            cmd = [
                sys.executable,
                str(script_path),
                "--working-dir",
                str(desired_working),
                "--month",
                report_month,
                "--year",
                report_year,
            ]
            run_cwd = str(REPORT_AUTOMATIONS_DIR)
        
        # Estimate duration based on past runs
        estimated_secs = max(60, _estimate_script_duration_seconds(script_name))
        from datetime import datetime as _dt
        
        _emit_log(f"Running: {' '.join(cmd)} (cwd={run_cwd})", important=True)
        _set_report_status(script_name, "running")
        
        # Create individual log file for this script
        script_log_file = _run_log_dir / f"{script_name.replace('.py', '')}.log"
        _emit_log(f"Creating individual log file: {script_log_file.name}", important=True)
        
        # Add environment setup for C6 and C4 to ensure they can find required files
        env = os.environ.copy()
        if script_name in ["NBD_MF_20_C4.py", "NBD_MF_20_C6.py", "NBD_MF_20_C5.py", "NBD_MF_20_C2.py", "NBD_MF_20_C3_report.py"]:
            # Ensure Python can find the script and any dependencies
            env["PYTHONPATH"] = str(REPORT_AUTOMATIONS_DIR)
        
        with subprocess.Popen(
            cmd,
            cwd=run_cwd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding='utf-8',
            errors='replace',  # Replace undecodable characters instead of crashing
            bufsize=1,
            env=env,
        ) as proc:
            # Track individual process and global stop requests
            global _current_proc, _stop_requested, _running_processes
            report_name = Path(script_name).stem
            
            with _proc_lock:
                _current_proc = proc
                _running_processes[report_name] = proc
            
            assert proc.stdout is not None
            # Open individual script log file
            with open(script_log_file, 'w', encoding='utf-8') as script_log:
                for line in proc.stdout:
                    # Check for individual script stop or global stop
                    if _stop_requested or report_name not in _running_processes:
                        break
                    raw = line.rstrip()
                    
                    # Write to individual script log file
                    script_log.write(raw + '\n')
                    script_log.flush()

                    # Skip verbose report output from web_run.log (already in individual log)
                    _emit_log(raw, important=False, skip_web_log=True)
                    # Check for log file creation
                    _detect_log_file_creation(raw)
                    # Append compact message to the per-report feed (lightly filtered)
                    _append_report_message(report_name, raw[:500])
            
            # Clean up process tracking
            with _proc_lock:
                # Check if script was stopped BEFORE removing from running processes
                was_stopped = _stop_requested or report_name not in _running_processes
                if report_name in _running_processes:
                    del _running_processes[report_name]
                if was_stopped:
                    try:
                        proc.terminate()
                    except Exception:
                        pass
                ret = proc.wait()
                _current_proc = None
                
    except Exception as exc:
        _emit_log(f"Failed to start {script_name}: {exc}", important=True)
        _set_report_status(script_name, "failed")
        return False, f"Failed to start {script_name}: {exc}"

    # Check if script was stopped vs failed
    if was_stopped:
        _emit_log(f"Script {script_name} stopped by user", important=True)
        _set_report_status(script_name, "stopped")
        return False, f"Script {script_name} stopped by user"
    elif ret != 0:
        _emit_log(f"Script {script_name} failed with exit code {ret}", important=True)
        _set_report_status(script_name, "failed")
        return False, f"Script {script_name} failed with exit code {ret}"

    _set_report_status(script_name, "completed")
    return True, f"Script {script_name} completed successfully"


def _run_download_bot(date_folder: Path) -> tuple[bool, str]:
    """Run the Selenium download bot before automations to fetch latest source files."""
    try:
        bot_path = BASE_DIR / "bots" / "report_download_bot_IA_C1C2.py"
        if not bot_path.exists():
            _emit_log(f"Download bot not found: {bot_path}", important=True)
            return False, f"Bot not found: {bot_path}"

        cmd = [
            sys.executable,
            str(bot_path),
        ]
        run_cwd = str(bot_path.parent)
        _emit_log(f"Running download bot: {' '.join(cmd)} (cwd={run_cwd})", important=True)
        _set_status(stage="running", message="Downloading latest reports (bot)")
        # Add bot to checklist (first time) and mark running
        if not any(it["name"] == "Download_Bot" for it in _report_checklist):
            _report_checklist.insert(0, {"name": "Download_Bot", "status": "running"})
            _report_messages["Download_Bot"] = deque(maxlen=200)
        else:
            _set_report_status("Download_Bot.py", "running")

        with subprocess.Popen(
            cmd,
            cwd=run_cwd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
        ) as proc:
            assert proc.stdout is not None
            for line in proc.stdout:
                raw = line.rstrip()
                _emit_log(raw, important=False, skip_web_log=True)
                # Mirror bot output into IA and Bot message feeds
                _append_report_message("NBD_MF_23_IA", raw[:500])
                _append_report_message("Download_Bot", raw[:500])
            ret = proc.wait()

        if ret != 0:
            _emit_log(f"Download bot failed with exit code {ret}", important=True)
            _set_report_status("Download_Bot.py", "failed")
            return False, f"Download bot failed with exit code {ret}"

        _emit_log("Download bot completed successfully", important=True)
        _append_report_message("NBD_MF_23_IA", "Download bot completed successfully")
        _set_report_status("Download_Bot.py", "completed")
        return True, "Download bot completed"

    except Exception as exc:
        _emit_log(f"Failed to run download bot: {exc}", important=True)
        return False, f"Failed to run download bot: {exc}"

def _run_c6_c4_thread(date_folder: Path, report_month: str, report_year: str) -> Dict[str, Tuple[bool, str]]:
    """Run C6 -> C5 -> C4 -> C2 -> C3 sequentially in a single thread."""
    results = {}
    thread_name = "C6_C4_Thread"
    _running_threads[thread_name] = threading.current_thread()

    try:
        _emit_log("Starting C6->C5->C4->C2->C3 thread", important=True)
        _emit_log(f"Working folder: {date_folder}", important=True)
        _emit_log(f"Report month: {report_month}, year: {report_year}", important=True)

        sequence = [
            ("NBD_MF_20_C6.py", "Running C6"),
            ("NBD_MF_20_C5.py", "Running C5 (after C6)"),
            ("NBD_MF_20_C4.py", "Running C4 (after C5)"),
            ("NBD_MF_20_C2.py", "Running C2 (after C4)"),
            ("NBD_MF_20_C3_report.py", "Running C3 (after C2)"),
        ]
        for script, status_msg in sequence:
            script_path = REPORT_AUTOMATIONS_DIR / script
            if not script_path.exists():
                msg = f"{script} not found at: {script_path}; skipping"
                _emit_log(msg, important=True)
                _set_report_status(script, "skipped")
                results[script.replace('.py','')] = (True, msg)
                continue
            _set_status(stage="running", message=status_msg)
            ok, msg = _run_single_script(script, date_folder, report_month, report_year)
            results[script.replace('.py','')] = (ok, msg)
            if not ok:
                _emit_log(f"{script} failed, stopping sequence: {msg}", important=True)
                return results
        _emit_log("C6->C5->C4->C2->C3 thread completed successfully", important=True)
            
    except Exception as e:
        _emit_log(f"C6+C4+C2+C3 thread failed: {e}", important=True)
        import traceback
        _emit_log(f"C6+C4+C2+C3 thread traceback: {traceback.format_exc()}", important=False, skip_web_log=True)
        results["NBD_MF_20_C6"] = results.get("NBD_MF_20_C6", (False, str(e)))
        results["NBD_MF_20_C4"] = results.get("NBD_MF_20_C4", (False, str(e)))
        results["NBD_MF_20_C2"] = results.get("NBD_MF_20_C2", (False, str(e)))
        results["NBD_MF_20_C3_report"] = results.get("NBD_MF_20_C3_report", (False, str(e)))
    finally:
        if thread_name in _running_threads:
            del _running_threads[thread_name]
    
    return results


def _run_ia_thread(date_folder: Path, report_month: str, report_year: str) -> Dict[str, Tuple[bool, str]]:
    """Run IA script followed by C1C2 script sequentially in a separate thread."""
    results = {}
    thread_name = "IA_C1C2_Thread"
    _running_threads[thread_name] = threading.current_thread()

    try:
        _emit_log("Starting IA → C1C2 thread (sequential)", important=True)
        _set_status(stage="running", message="Running IA → C1C2 (parallel)")

        # Run IA first
        success_ia, msg_ia = _run_single_script("NBD_MF_23_IA.py", date_folder, report_month, report_year)
        results["NBD_MF_23_IA"] = (success_ia, msg_ia)

        if success_ia:
            _emit_log("IA completed successfully, starting C1C2", important=True)

            # Run C1C2 after IA succeeds
            success_c1c2, msg_c1c2 = _run_single_script("NBD_MF_23_C1C2.py", date_folder, report_month, report_year)
            results["NBD_MF_23_C1C2"] = (success_c1c2, msg_c1c2)

            if success_c1c2:
                _emit_log("IA → C1C2 thread completed successfully", important=True)
            else:
                _emit_log("IA succeeded but C1C2 failed", important=True)
        else:
            _emit_log("IA failed, skipping C1C2", important=True)
            results["NBD_MF_23_C1C2"] = (False, "Skipped due to IA failure")
            _set_report_status("NBD_MF_23_C1C2.py", "skipped")

    except Exception as e:
        _emit_log(f"IA → C1C2 thread failed: {e}", important=True)
        results["NBD_MF_23_IA"] = (False, str(e))
        results["NBD_MF_23_C1C2"] = (False, "Skipped due to thread error")
    finally:
        if thread_name in _running_threads:
            del _running_threads[thread_name]

    return results


def _run_23_c4_thread(date_folder: Path, report_month: str, report_year: str) -> Dict[str, Tuple[bool, str]]:
    """Run NBD-MF-23-C4.py as subprocess."""
    results = {}
    thread_name = "NBD_MF_23_C4_Thread"
    _running_threads[thread_name] = threading.current_thread()
    script_filename = "NBD-MF-23-C4.py"

    try:
        _emit_log("Starting NBD_MF_23_C4 (subprocess)", important=True)
        _ensure_checklist_item(script_filename)
        _set_report_status(script_filename, "running")

        script_path = REPORT_AUTOMATIONS_DIR / script_filename
        if not script_path.exists():
            msg = f"{script_filename} not found at: {script_path}; skipping"
            _emit_log(msg, important=True)
            _set_report_status(script_filename, "skipped")
            results["NBD_MF_23_C4"] = (True, msg)
            return results

        # Determine working directory
        c4_folder = date_folder / "NBD_MF_23_C4"
        if not c4_folder.exists():
            c4_folder = date_folder

        _emit_log(f"Working directory for C4: {c4_folder}", important=True)

        # Run as subprocess
        cmd = [sys.executable, str(script_path)]
        _emit_log(f"Running: {' '.join(cmd)} (cwd={c4_folder})", important=True)

        # Create individual log file
        log_file = _run_log_dir / script_filename.replace(".py", ".log")
        _emit_log(f"Creating individual log file: {log_file.name}", important=True)

        with subprocess.Popen(
            cmd,
            cwd=str(c4_folder),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding='utf-8',
            errors='replace'
        ) as proc:
            with open(log_file, "w", encoding="utf-8", errors="replace") as lf:
                for line in proc.stdout:
                    line = line.rstrip()
                    if line:
                        lf.write(line + "\n")
                        _emit_log(line, important=False, skip_web_log=True)

                proc.wait()

                if proc.returncode == 0:
                    _set_report_status(script_filename, "completed")
                    results["NBD_MF_23_C4"] = (True, "Completed")
                    _emit_log("NBD_MF_23_C4 completed successfully", important=True)
                else:
                    msg = f"Script failed with exit code {proc.returncode}"
                    _emit_log(f"Script {script_filename} {msg}", important=True)
                    _set_report_status(script_filename, "failed")
                    results["NBD_MF_23_C4"] = (False, msg)

    except Exception as e:
        _emit_log(f"NBD_MF_23_C4 thread failed: {e}", important=True)
        import traceback
        _emit_log(traceback.format_exc(), important=False, skip_web_log=True)
        _set_report_status(script_filename, "failed")
        results["NBD_MF_23_C4"] = (False, str(e))
    finally:
        if thread_name in _running_threads:
            del _running_threads[thread_name]

    return results


def _run_ga_is_thread(date_folder: Path, report_month: str, report_year: str) -> Dict[str, Tuple[bool, str]]:
    """Run NBD_MF_10_GA_IS.py as subprocess to avoid COM threading issues with long operations."""
    results = {}
    thread_name = "NBD_MF_10_GA_IS_Thread"
    _running_threads[thread_name] = threading.current_thread()
    script_filename = "NBD_MF_10_GA_IS.py"

    try:
        _emit_log("Starting NBD_MF_10_GA_IS (subprocess)", important=True)
        _ensure_checklist_item(script_filename)
        _set_report_status(script_filename, "running")
        _set_status(stage="running", message="Running NBD_MF_10_GA_IS (parallel)")

        script_path = REPORT_AUTOMATIONS_DIR / script_filename
        if not script_path.exists():
            msg = f"{script_filename} not found at: {script_path}; skipping"
            _emit_log(msg, important=True)
            _set_report_status(script_filename, "skipped")
            results["NBD_MF_10_GA_IS"] = (True, msg)
            return results

        # Determine working directory
        ga_is_folder = date_folder / "NBD-MF-10-GA & NBD-MF-11-IS"
        if not ga_is_folder.exists():
            ga_is_folder = date_folder

        _emit_log(f"Working directory for GA_IS: {ga_is_folder}", important=True)

        # Create individual log file
        log_file = _run_log_dir / script_filename.replace(".py", ".log")
        _emit_log(f"Creating individual log file: {log_file.name}", important=True)

        # Run as subprocess
        env = os.environ.copy()
        env['REPORT_LOG_DIR'] = str(_run_log_dir)

        cmd = [sys.executable, str(script_path)]
        _emit_log(f"Running: {' '.join(cmd)} (cwd={ga_is_folder})", important=True)

        with open(log_file, "w", encoding="utf-8", errors="replace") as lf:
            process = subprocess.Popen(
                cmd,
                cwd=str(ga_is_folder),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
                env=env
            )

            # Stream output line by line
            for line in process.stdout:
                line = line.rstrip()
                if line:
                    lf.write(line + "\n")
                    lf.flush()
                    _emit_log(line, important=False, skip_web_log=True)

            process.wait()

        if process.returncode == 0:
            _set_report_status(script_filename, "completed")
            results["NBD_MF_10_GA_IS"] = (True, "Completed")
            _emit_log("NBD_MF_10_GA_IS completed successfully", important=True)
        else:
            _set_report_status(script_filename, "failed")
            results["NBD_MF_10_GA_IS"] = (False, f"Exit code {process.returncode}")
            _emit_log(f"NBD_MF_10_GA_IS failed with exit code {process.returncode}", important=True)

    except Exception as e:
        _emit_log(f"NBD_MF_10_GA_IS thread failed: {e}", important=True)
        import traceback
        _emit_log(traceback.format_exc(), important=False, skip_web_log=True)
        _set_report_status(script_filename, "failed")
        results["NBD_MF_10_GA_IS"] = (False, str(e))
    finally:
        if thread_name in _running_threads:
            del _running_threads[thread_name]

    return results


def _run_single_script_in_dir(script_name: str, working_dir: Path, log_file: Path, env: dict = None) -> tuple[bool, str]:
    """Run a single script in a specific working directory with custom environment."""
    script_path = REPORT_AUTOMATIONS_DIR / script_name
    if not script_path.exists():
        return False, f"Script not found: {script_path}"

    _ensure_checklist_item(script_name)
    _set_report_status(script_name, "running")

    if env is None:
        env = os.environ.copy()

    try:
        with log_file.open("w", encoding="utf-8") as lf:
            proc = subprocess.Popen(
                [sys.executable, str(script_path)],
                cwd=str(working_dir),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding='utf-8',
                errors='replace',
                bufsize=1,
                env=env
            )

            with _proc_lock:
                _running_processes[script_name] = proc

            for line in iter(proc.stdout.readline, ''):
                if not line:
                    break
                _emit_log(line.rstrip(), important=False, skip_web_log=True, script=script_name)
                lf.write(line)
                lf.flush()

                if _stop_requested:
                    proc.terminate()
                    try:
                        proc.wait(timeout=5)
                    except subprocess.TimeoutExpired:
                        proc.kill()
                    _set_report_status(script_name, "stopped")
                    return False, "Stopped by user"

            proc.wait()

            with _proc_lock:
                if script_name in _running_processes:
                    del _running_processes[script_name]

            if proc.returncode == 0:
                _set_report_status(script_name, "completed")
                return True, "Completed"
            else:
                _set_report_status(script_name, "failed")
                return False, f"Exit code {proc.returncode}"

    except Exception as e:
        _set_report_status(script_name, "failed")
        return False, str(e)


def run_reports_for_frequency(frequency: str, ui_date: str) -> tuple[bool, str]:
    """Run report automation scripts using threading for the given frequency.

    For monthly reports:
    - Thread 1 (sequential): C6 → C5 → C4 → C2 → C3
    - Thread 2 (sequential): IA → C1C2
    - Thread 3 (parallel): NBD-MF-23-C4
    - Thread 4 (parallel): GA-IS

    Returns (ok, message).
    """
    scripts_by_frequency: dict[str, list[str]] = {
        "monthly": [
           "NBD_MF_20_C6.py", "NBD_MF_20_C5.py", "NBD_MF_20_C4.py", "NBD_MF_20_C2.py", "NBD_MF_20_C3_report.py",
           "NBD_MF_23_IA.py", "NBD_MF_23_C1C2.py", "NBD-MF-23-C4.py", "NBD_MF_10_GA_IS.py",
        ],
        "weekly": [],
        "quarterly": [],
        "annually": [],
    }

    scripts = scripts_by_frequency.get(frequency, [])
    if not scripts:
        return True, f"No automation scripts configured for {frequency}."

    # Initialize checklist for UI only if not already initialized; otherwise ensure all scripts are present
    if not _report_checklist:
        _init_report_checklist(scripts)
    else:
        _ensure_reports_in_checklist(scripts)

    # Determine working dir: there should be exactly one subfolder inside working/<frequency>
    working_freq_dir = WORKING_DIR / frequency
    if not working_freq_dir.exists():
        return False, f"Working directory not found: {working_freq_dir}"
    subdirs = [p for p in working_freq_dir.iterdir() if p.is_dir()]
    if len(subdirs) != 1:
        return False, f"Expected exactly one folder in {working_freq_dir}, found {len(subdirs)}."
    date_folder = subdirs[0]

    # Derive reporting month (previous month) and year from picked date
    dt = parse_ui_date(ui_date)
    prev_month_anchor = dt.replace(day=1) - timedelta(days=1)
    report_month = prev_month_anchor.strftime("%b")  # e.g., Jul
    report_year = str(prev_month_anchor.year)

    global _thread_pool
    if not _thread_pool:
        _thread_pool = ThreadPoolExecutor(max_workers=4)

    try:
        # Submit both thread tasks
        future_c6_c4 = _thread_pool.submit(_run_c6_c4_thread, date_folder, report_month, report_year)
        future_ia = _thread_pool.submit(_run_ia_thread, date_folder, report_month, report_year)
        future_23_c4 = _thread_pool.submit(_run_23_c4_thread, date_folder, report_month, report_year)
        future_ga_is = _thread_pool.submit(_run_ga_is_thread, date_folder, report_month, report_year)
        
        # Wait for all threads to complete
        results_c6_c4 = future_c6_c4.result()
        results_ia = future_ia.result()
        results_23_c4 = future_23_c4.result()
        results_ga_is = future_ga_is.result()
        
        # Combine results
        all_results = {**results_c6_c4, **results_ia, **results_23_c4, **results_ga_is}
        
        # Check if all scripts succeeded
        all_success = all(success for success, _ in all_results.values())
        
        # Build summary message
        executed = []
        failed = []
        for script_key, (success, msg) in all_results.items():
            script_name = script_key + ".py"
            if success:
                executed.append(script_name)
            else:
                failed.append(f"{script_name}: {msg}")
        
        if all_success:
            return True, f"All scripts completed successfully: {', '.join(executed)}"
        else:
            failure_msg = "; ".join(failed)
            return False, f"Some scripts failed - Successful: {', '.join(executed)}; Failed: {failure_msg}"
            
    except Exception as e:
        _emit_log(f"Threading execution failed: {e}", important=True)
        return False, f"Threading execution failed: {e}"


app = Flask(__name__)
# Require a proper secret key in production; default only for local dev
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "change-me-in-prod")


@app.route("/", methods=["GET"]) 
def index():
    ensure_directories_exist()
    today_iso = datetime.today().strftime("%Y-%m-%d")
    return render_template("index.html", today_iso=today_iso)


@app.route("/start", methods=["POST"]) 
def start_automation():
    ensure_directories_exist()
    frequency = request.form.get("frequency", "").strip().lower()
    date_str = request.form.get("date", "").strip()
    _set_status(running=True, stage="starting", frequency=frequency, date=date_str, message="Starting automation")
    # Initialize new run with timestamped logs
    _initialize_new_run()
    # Reset per-run buffers
    _clear_logs()
    _status_feed.clear()
    global _report_checklist, _report_messages, _last_output_dir, _stop_requested
    _report_checklist = []
    _report_messages = {}
    _last_output_dir = None
    _stop_requested = False
    # Clear running processes and threads
    with _proc_lock:
        _running_processes.clear()
    _running_threads.clear()
    
    if frequency not in {"weekly", "monthly", "quarterly", "annually"}:
        flash("Invalid frequency.", "error")
        return redirect(url_for("index"))
    
    # Get the last completed date for this frequency
    last_ui_date = get_last_completed_date(frequency)
    if not last_ui_date:
        flash(f"No completed {frequency} date found. Use Complete first.", "warning")
        _set_status(stage="error", running=False, message="No completed date found")
        return redirect(url_for("index"))

    safe_name = ui_date_to_folder_name(last_ui_date)
    _set_status(stage="copying", date=last_ui_date, message=f"Copying latest {safe_name} folder to working")

    # Always copy the latest versioned folder for the last completed date to working
    ok, message = copy_latest_for_date_to_working(frequency, last_ui_date, clean_working=True)
    flash(message, "success" if ok else "warning")
    if not ok:
        _set_status(stage="error", running=False, message=message)
        return redirect(url_for("index"))

    # Determine working date folder for bot/report steps
    working_freq_dir = WORKING_DIR / frequency
    subdirs = [p for p in working_freq_dir.iterdir() if p.is_dir()]
    if len(subdirs) != 1:
        _set_status(stage="error", running=False, message=f"Expected one folder under {working_freq_dir}, found {len(subdirs)}")
        return redirect(url_for("index"))
    date_folder = subdirs[0]

    # 1) Run download bot first
    # ok_bot, msg_bot = _run_download_bot(date_folder)
    # flash(msg_bot, "success" if ok_bot else "warning")
    # if not ok_bot:
    #     _set_status(stage="error", running=False, message=msg_bot)
    #     return redirect(url_for("index"))

    # 2) Run reports after bot completes
    _set_status(stage="running", message="Running automation scripts in parallel")
    ok_reports, msg_reports = run_reports_for_frequency(frequency, last_ui_date)
    flash(msg_reports, "success" if ok_reports else "warning")

    # Compute references to the single working date folder (for rename/download setup in both success/failure)
    safe_name = ui_date_to_folder_name(last_ui_date)
    parent = OUTPUTS_DIR / frequency
    working_freq_dir = WORKING_DIR / frequency
    subdirs = [p for p in working_freq_dir.iterdir() if p.is_dir()]

    if ok_reports:
        # Save and download only if reports were successful
        if len(subdirs) == 1:
            source_folder = subdirs[0]
            # 1) Rename working folder to picked date base name first (so download matches picked date)
            try:
                desired_name = safe_name
                if source_folder.name != desired_name:
                    desired_path = working_freq_dir / desired_name
                    if desired_path.exists():
                        shutil.rmtree(desired_path, ignore_errors=True)
                    old_name = source_folder.name
                    source_folder = source_folder.rename(desired_path)
                    _emit_log(f"Renamed working folder '{old_name}' to '{source_folder.name}'", important=True)
            except Exception:
                _emit_log("Rename of working folder failed; proceeding with existing name", important=True)

            # 2) Signal browser download is ready from the renamed working folder
            _last_output_dir = source_folder
            _set_status(stage="download_ready", message="Preparing download")

            # 3) Compute outputs target and copy after signaling download readiness
            existing = find_latest_versioned_folder(parent, safe_name)
            if existing is None:
                target = parent / safe_name
            else:
                last = existing.name
                if last == safe_name:
                    target = parent / f"{safe_name}(1)"
                else:
                    middle = last[len(safe_name) + 1 : -1]
                    next_n = int(middle) + 1 if middle.isdigit() else 1
                    target = parent / f"{safe_name}({next_n})"

            try:
                shutil.copytree(source_folder, target)
                flash(
                    f"Saved working folder '{source_folder.name}' to outputs/{frequency} as '{target.name}'.",
                    "success",
                )
                _emit_log(f"Copied working folder to outputs: {target}", important=True)
                
                # Zip creation is handled via the /download-latest endpoint for browser-only download
            except FileExistsError:
                flash(f"Destination already exists: {target}.", "warning")
                _emit_log(f"Destination already exists: {target}", important=True)
            except Exception as exc:
                flash(f"Failed to save to outputs: {exc}", "warning")
                _emit_log(f"Failed to save to outputs: {exc}", important=True)
        else:
            flash(
                f"Expected one folder under {working_freq_dir}, found {len(subdirs)}. Skipped saving to outputs.",
                "warning",
            )
            _emit_log(f"Save skipped: expected one folder under {working_freq_dir}, found {len(subdirs)}", important=True)

        # Append history record (use working folder path set in _last_output_dir)
        try:
            script_names = [it["name"] + ".py" for it in _report_checklist]
            _append_run_history(frequency, last_ui_date, script_names, _report_checklist, _last_output_dir)
        except Exception as exc:
            _emit_log(f"Failed to write history: {exc}", important=True)

        # Mark completion
        _set_status(stage="done", running=False, message="Completed")
    else:
        # Even on failure, prepare the download from the working folder and record history
        if len(subdirs) == 1:
            source_folder = subdirs[0]
            try:
                desired_name = safe_name
                if source_folder.name != desired_name:
                    desired_path = working_freq_dir / desired_name
                    if desired_path.exists():
                        shutil.rmtree(desired_path, ignore_errors=True)
                    old_name = source_folder.name
                    source_folder = source_folder.rename(desired_path)
                    _emit_log(f"Renamed working folder '{old_name}' to '{source_folder.name}' (failed run)", important=True)
            except Exception:
                _emit_log("Rename of working folder failed on error path; proceeding with existing name", important=True)
            # Make the failed run available for download
            _last_output_dir = source_folder
            _set_status(stage="download_ready", message="Preparing download (partial)")
        else:
            _emit_log(f"Download prep skipped: expected one folder under {working_freq_dir}, found {len(subdirs)}", important=True)
        # Append history even for failed runs
        try:
            script_names = [it["name"] + ".py" for it in _report_checklist]
            _append_run_history(frequency, last_ui_date, script_names, _report_checklist, _last_output_dir)
        except Exception as exc:
            _emit_log(f"Failed to write history: {exc}", important=True)
        # Do not override stage so auto-download can trigger; only mark running False and include message
        _set_status(running=False, message=msg_reports)
    
    # Clean up thread pool
    global _thread_pool
    if _thread_pool:
        try:
            _thread_pool.shutdown(wait=True)
        except Exception:
            pass
        _thread_pool = None
    
    return redirect(url_for("index"))

@app.route("/complete", methods=["POST"]) 
def mark_complete():
    ensure_directories_exist()
    frequency = request.form.get("frequency", "").strip().lower()
    date_str = request.form.get("date", "").strip()
    if frequency not in {"weekly", "monthly", "quarterly", "annually"}:
        flash("Invalid frequency.", "error")
        return redirect(url_for("index"))
    if not date_str:
        flash("Please pick a date.", "error")
        return redirect(url_for("index"))
    
    # Only append the picked date to completed_dates.txt (no copying)
    append_completed_date(frequency, date_str)
    flash(f"Marked {frequency} as completed for {ui_date_to_display(date_str)}.", "success")

    # Reset status to idle (no copying or processing)
    _reset_status()
    return redirect(url_for("index"))
    

@app.route("/status", methods=["GET"]) 
def status():
    with _status_lock:
        return jsonify(_status)


@app.route("/logs", methods=["GET"]) 
def logs():
    with _log_lock:
        content = "\n".join(_log_buffer)
    return app.response_class(content, mimetype="text/plain")


@app.route("/status-feed", methods=["GET"]) 
def status_feed():
    content = "\n".join(_status_feed)
    return app.response_class(content, mimetype="text/plain")


@app.route("/report-checklist", methods=["GET"]) 
def report_checklist():
    return jsonify(_report_checklist)


@app.route("/report-messages/<report_stem>", methods=["GET"]) 
def report_messages(report_stem: str):
    buf = _report_messages.get(report_stem)
    if buf is None:
        return app.response_class("", mimetype="text/plain")
    return app.response_class("\n".join(buf), mimetype="text/plain")


@app.route("/history")
def history():
    if not HISTORY_XLSX.exists():
        return jsonify([])
    wb = openpyxl.load_workbook(HISTORY_XLSX, read_only=True)
    ws = wb.active
    data = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        data.append(dict(
            timestamp=row[0],
            frequency=row[1],
            picked_date=row[2],
            reports=row[3],
            statuses=row[4],
            output_folder=row[5]
        ))
    wb.close()
    return jsonify(data)


@app.route("/stop", methods=["POST"]) 
def stop():
    _request_stop()
    _set_status(stage="stopping", message="Stopping requested")
    return jsonify({"ok": True})


@app.route("/stop-script", methods=["POST"])
def stop_individual_script():
    data = request.get_json()
    report_name = data.get("report_name", "").strip()
    
    if not report_name:
        return jsonify({"success": False, "message": "Report name required"}), 400
    
    success = _stop_individual_script(report_name)
    if success:
        # Update the report status to stopped
        for item in _report_checklist:
            if item["name"] == report_name:
                item["status"] = "stopped"
                break
        return jsonify({"success": True, "message": f"Stopped {report_name}"})
    else:
        return jsonify({"success": False, "message": f"Report {report_name} not found or already stopped"})


@app.route("/download-latest", methods=["GET"])
def download_latest():
    # Create zip with specific output files only
    global _last_output_dir
    if _last_output_dir is None or not _last_output_dir.exists():
        return jsonify({"error": "No output to download."}), 400

    base_name = _last_output_dir.name
    temp_zip_dir = BASE_DIR / "temp_zips"
    try:
        temp_zip_dir.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass

    zip_path = temp_zip_dir / f"{base_name}.zip"

    # Define specific files to include in the zip
    # These paths are relative to the working directory
    files_to_zip = [
        ("NBD_MF_23_C4", "Loan CF Analysis*.xlsb"),
        ("NBD_MF_23_IA", "Prod. wise Class. of Loans*.xlsb"),
        ("NBD-MF-10-GA & NBD-MF-11-IS", "NBD-MF-10-GA & NBD-MF-11-IS*.xlsx"),
        ("", "NBD_MF_20_C2_*_report.xlsx"),  # Root of date folder
        ("", "NBD_NBL_MF_20_C5_*_report.xlsx"),  # Root of date folder
    ]

    import zipfile
    import glob as glob_module

    try:
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            files_added = 0
            for subfolder, pattern in files_to_zip:
                # Build search path
                if subfolder:
                    search_path = _last_output_dir / subfolder / pattern
                else:
                    search_path = _last_output_dir / pattern

                # Find matching files
                matches = glob_module.glob(str(search_path))

                for file_path in matches:
                    file_path = Path(file_path)
                    if file_path.exists() and file_path.is_file():
                        # Add to zip with just the filename (no folder structure)
                        arcname = file_path.name
                        zipf.write(file_path, arcname=arcname)
                        files_added += 1
                        _emit_log(f"Added to zip: {file_path.name}", important=False, skip_web_log=True)

            _emit_log(f"Created zip with {files_added} files", important=True)

            if files_added == 0:
                return jsonify({"error": "No output files found to download."}), 400

        return send_file(str(zip_path), as_attachment=True, download_name=f"{base_name}.zip")

    except Exception as e:
        _emit_log(f"Failed to create zip: {e}", important=True)
        return jsonify({"error": f"Failed to create zip: {e}"}), 500


if __name__ == "__main__":
    # Kill any running Excel instances to prevent COM errors
    kill_excel_instances()

    ensure_directories_exist()

    # Production-safe defaults
    # host = os.environ.get("FLASK_HOST", "0.0.0.0")
    # port = int(os.environ.get("PORT", 5000))

    app.run(host="127.0.0.1", port=5000, debug=True)

    # print(f"Starting Waitress server on http://{host}:{port}")
    # serve(app, host=host, port=port)