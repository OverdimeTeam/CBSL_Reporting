import os
import shutil
from datetime import datetime
from pathlib import Path

from flask import Flask, render_template, request, redirect, url_for, flash


BASE_DIR = Path(__file__).resolve().parent
OUTPUTS_DIR = BASE_DIR / "outputs"
WORKING_DIR = BASE_DIR / "working"
REPORT_AUTOMATIONS_DIR = BASE_DIR / "report_automations"


def ensure_directories_exist() -> None:
    for sub in ["weekly", "monthly", "quarterly", "annually"]:
        target = OUTPUTS_DIR / sub
        target.mkdir(parents=True, exist_ok=True)
        completed_file = target / "completed_dates.txt"
        if not completed_file.exists():
            completed_file.write_text("")
    for sub in ["weekly", "monthly", "quarterly", "annually"]:
        (WORKING_DIR / sub).mkdir(parents=True, exist_ok=True)


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


def copy_latest_for_date_to_working(frequency: str, ui_date: str) -> tuple[bool, str]:
    """Copy latest versioned folder for the picked date into working/<frequency>."""
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
    dest = WORKING_DIR / frequency / latest_folder.name
    try:
        if dest.exists():
            shutil.rmtree(dest)
        shutil.copytree(latest_folder, dest)
    except Exception as exc:
        return False, f"Failed to copy '{latest_folder.name}' to '{dest}': {exc}"
    return True, f"Copied {latest_folder.name} to working/{frequency}."


def copy_from_last_completed_to_working(frequency: str) -> tuple[bool, str]:
    last_ui_date = get_last_completed_date(frequency)
    if not last_ui_date:
        return False, f"No completed {frequency} date found."
    return copy_latest_for_date_to_working(frequency, last_ui_date)


app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-secret-key")


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
    if frequency not in {"weekly", "monthly", "quarterly", "annually"}:
        flash("Invalid frequency.", "error")
        return redirect(url_for("index"))
    if not date_str:
        flash("Please pick a date.", "error")
        return redirect(url_for("index"))

    ok, message = copy_from_last_completed_to_working(frequency)
    flash(message, "success" if ok else "warning")

    safe_name = ui_date_to_folder_name(date_str)
    parent = OUTPUTS_DIR / frequency

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

    target.mkdir(parents=True, exist_ok=True)

    flash(f"Created folder {target.name} under {frequency}.", "success")
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
    append_completed_date(frequency, date_str)
    flash(f"Marked {frequency} as completed for {ui_date_to_display(date_str)}.", "success")
    return redirect(url_for("index"))


if __name__ == "__main__":
    ensure_directories_exist()
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)


