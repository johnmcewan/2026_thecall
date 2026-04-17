"""
quark_batch_export.py
─────────────────────────────────────────────────────────────────────────────
Batch-exports QuarkXPress 2016 (.qzp / .qxp) files to PDF and JSON.

Pipeline per file:
  1. Open file in QuarkXPress 2016 via subprocess
  2. Auto-dismiss "Repair it" dialog
  3. Auto-dismiss "Missing fonts" dialog
  4. Export > Layout as PDF  →  <OUTPUT_ROOT>/<folder>/<db_id>.pdf
  5. Close file without saving
  6. Extract text from PDF via pdfplumber
  7. Build standardised JSON and update PostgreSQL

Dependencies:
    pip install pyautogui pygetwindow pdfplumber psycopg2-binary pillow

Run from your thecall2 conda env:
    python quark_batch_export.py

IMPORTANT: Keep QuarkXPress 2016 closed before running.
           Do not move the mouse or use the keyboard while the script runs.
           Use a separate machine or RDP session if you need to work
           while processing runs overnight.
"""

import os
import re
import sys
import json
import time
import shutil
import logging
import subprocess
from pathlib import Path

import psycopg2
from psycopg2 import sql, extras
import pdfplumber
import pyautogui
import pygetwindow as gw

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURATION  ── edit these before running
# ─────────────────────────────────────────────────────────────────────────────

QUARK_EXE       = r"C:\Program Files\Quark\QuarkXPress 2016.exe"
ASSEMBLED_ROOT  = r"E:\Callproject\assembled"          # source .qzp files
OUTPUT_ROOT     = r"E:\Callproject\c_out_quark"        # PDFs + JSON land here
PDF_SUBDIR      = "pdfs"                               # OUTPUT_ROOT\pdfs\
LOG_FILE        = r"E:\Callproject\quark_batch.log"

DB_PARAMS = {
    "dbname":   "thecall",
    "user":     "postgres",
    "password": "password",
    "host":     "localhost",
    "port":     "5432",
}
TABLE_SCHEMA = "articles"
TABLE_NAME   = "filelist"

# How many files to process per run (set to None for all)
BATCH_LIMIT = None

# Seconds to wait for QuarkXPress window to appear after launch
QUARK_LAUNCH_TIMEOUT = 30

# Seconds to wait between UI interactions (increase on slower machines)
UI_PAUSE = 0.6

# Seconds to wait for the PDF Save dialog to appear
PDF_DIALOG_TIMEOUT = 20

# ─────────────────────────────────────────────────────────────────────────────
# LOGGING
# ─────────────────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)

# ─────────────────────────────────────────────────────────────────────────────
# PYAUTOGUI SAFETY
# ─────────────────────────────────────────────────────────────────────────────

pyautogui.FAILSAFE  = True   # move mouse to top-left corner to abort
pyautogui.PAUSE     = 0.3    # small global pause between every pyautogui call

# ─────────────────────────────────────────────────────────────────────────────
# DATABASE HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def db_connect():
    conn = psycopg2.connect(**DB_PARAMS)
    conn.autocommit = False
    return conn


def db_fetch_pending(conn, limit=None):
    """Return rows that have not yet been processed (json_raw IS NULL)."""
    q = sql.SQL(
        "SELECT id, filename, folder "
        "FROM {}.{} "
        "WHERE json_raw IS NULL "
        "ORDER BY id"
    ).format(sql.Identifier(TABLE_SCHEMA), sql.Identifier(TABLE_NAME))

    if limit:
        q = sql.SQL("{} LIMIT {}").format(q, sql.Literal(limit))

    with conn.cursor() as cur:
        cur.execute(q)
        return cur.fetchall()


def db_update_row(conn, db_id, json_str):
    q = sql.SQL(
        "UPDATE {}.{} SET json_raw = %s::jsonb WHERE id = %s"
    ).format(sql.Identifier(TABLE_SCHEMA), sql.Identifier(TABLE_NAME))
    with conn.cursor() as cur:
        cur.execute(q, (json_str, db_id))
    conn.commit()


# ─────────────────────────────────────────────────────────────────────────────
# WINDOW HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def wait_for_window(title_fragment, timeout=QUARK_LAUNCH_TIMEOUT):
    """
    Poll until a window whose title contains title_fragment appears.
    Returns the window object or raises TimeoutError.
    """
    deadline = time.time() + timeout
    while time.time() < deadline:
        matches = [w for w in gw.getAllWindows()
                   if title_fragment.lower() in w.title.lower()]
        if matches:
            return matches[0]
        time.sleep(0.5)
    raise TimeoutError(f"Window containing '{title_fragment}' did not appear "
                       f"within {timeout}s")


def window_exists(title_fragment):
    return any(title_fragment.lower() in w.title.lower()
               for w in gw.getAllWindows())


def focus_window(title_fragment):
    """Bring window to front."""
    wins = [w for w in gw.getAllWindows()
            if title_fragment.lower() in w.title.lower()]
    if wins:
        try:
            wins[0].activate()
            time.sleep(0.4)
        except Exception:
            pass


def close_quark_completely():
    """Force-kill QuarkXPress if it is still running."""
    subprocess.run(
        ["taskkill", "/f", "/im", "QuarkXPress 2016.exe"],
        capture_output=True
    )
    time.sleep(2)


# ─────────────────────────────────────────────────────────────────────────────
# DIALOG DISMISSERS
# ─────────────────────────────────────────────────────────────────────────────

def dismiss_repair_dialog(timeout=15):
    """Click 'Repair it' if the repair dialog appears."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        if window_exists("repair") or window_exists("Repair"):
            log.info("    → Dismissing 'Repair it' dialog")
            pyautogui.press("enter")          # 'Repair it' is default button
            time.sleep(UI_PAUSE)
            return True
        # Also try to find it as a child of the main QuarkXPress window
        # by looking for the button text on screen
        try:
            loc = pyautogui.locateOnScreen("repair_it_btn.png", confidence=0.8)
            if loc:
                pyautogui.click(loc)
                time.sleep(UI_PAUSE)
                return True
        except Exception:
            pass
        time.sleep(0.4)
    return False  # dialog did not appear — that is fine


def dismiss_font_dialog(timeout=15):
    """Click 'Continue' on the missing-fonts warning."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        if window_exists("font") or window_exists("Font"):
            log.info("    → Dismissing missing-fonts dialog")
            pyautogui.press("enter")          # 'Continue' is default button
            time.sleep(UI_PAUSE)
            return True
        time.sleep(0.4)
    return False


def dismiss_save_changes_dialog(timeout=10):
    """Click 'No' on the 'Save changes to the project?' dialog on close."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        if window_exists("Save") or window_exists("save"):
            log.info("    → Dismissing 'Save changes' dialog")
            # Tab to 'No' button and press it
            # 'No' is typically the second button; Alt+N is the accelerator
            pyautogui.hotkey("alt", "n")
            time.sleep(UI_PAUSE)
            return True
        time.sleep(0.3)
    return False


# ─────────────────────────────────────────────────────────────────────────────
# QUARK AUTOMATION CORE
# ─────────────────────────────────────────────────────────────────────────────

def open_file_in_quark(filepath):
    """
    Launch QuarkXPress with the given file and wait for it to be ready.
    Handles the Repair and Font dialogs automatically.
    Returns True on success.
    """
    log.info(f"    Opening: {filepath.name}")

    # Launch QuarkXPress with the file
    subprocess.Popen([QUARK_EXE, str(filepath)])

    # Wait for the main QuarkXPress window
    try:
        wait_for_window("QuarkXPress", timeout=QUARK_LAUNCH_TIMEOUT)
    except TimeoutError:
        log.error("    QuarkXPress did not launch in time")
        return False

    time.sleep(2)  # let QuarkXPress fully initialise

    # Handle dialogs — they may appear in either order, or not at all
    # We check for both in a loop for up to 30 seconds total
    deadline = time.time() + 30
    repair_done = False
    font_done   = False

    while time.time() < deadline:
        if not repair_done:
            if window_exists("Repair") or window_exists("repair"):
                log.info("    → Clicking 'Repair it'")
                pyautogui.press("enter")
                time.sleep(UI_PAUSE)
                repair_done = True
                continue

        if not font_done:
            if window_exists("font") or window_exists("Font"):
                log.info("    → Clicking 'Continue' (fonts)")
                pyautogui.press("enter")
                time.sleep(UI_PAUSE)
                font_done = True
                continue

        # If neither dialog is present and QXP main window is focused, we're ready
        if window_exists("QuarkXPress"):
            break

        time.sleep(0.5)

    time.sleep(1)  # final settle
    return True


def export_pdf(pdf_path):
    """
    Trigger File > Export > Layout as PDF and save to pdf_path.
    Returns True on success.
    """
    focus_window("QuarkXPress")
    time.sleep(0.3)

    log.info(f"    Exporting PDF → {pdf_path.name}")

    # File menu
    pyautogui.hotkey("alt", "f")
    time.sleep(UI_PAUSE)

    # Export submenu  (keyboard: E then L for "Layout as PDF")
    pyautogui.press("e")
    time.sleep(UI_PAUSE)
    pyautogui.press("l")
    time.sleep(UI_PAUSE)

    # Wait for the Save dialog
    try:
        wait_for_window("PDF", timeout=PDF_DIALOG_TIMEOUT)
    except TimeoutError:
        log.error("    PDF Save dialog did not appear")
        # Escape out to recover
        pyautogui.press("escape")
        pyautogui.press("escape")
        return False

    time.sleep(0.5)

    # The dialog pre-fills the filename in the same folder.
    # We need to change the save location to our output folder.
    # Use Ctrl+A in the filename field to select all, then type full path.

    # First navigate to the filename field (it should already be focused)
    # Clear it and type the full destination path
    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.3)
    pyautogui.typewrite(str(pdf_path), interval=0.03)
    time.sleep(0.3)

    # Click Save
    pyautogui.press("enter")
    time.sleep(3)  # wait for PDF generation

    # Handle any "file exists — overwrite?" dialog
    if window_exists("already exists") or window_exists("Replace") or window_exists("Overwrite"):
        pyautogui.press("enter")
        time.sleep(1)

    return True


def close_document():
    """Close the current document, clicking 'No' on the save-changes dialog."""
    focus_window("QuarkXPress")
    time.sleep(0.3)

    log.info("    Closing document")
    pyautogui.hotkey("ctrl", "w")
    time.sleep(UI_PAUSE)

    dismiss_save_changes_dialog(timeout=8)
    time.sleep(0.5)


# ─────────────────────────────────────────────────────────────────────────────
# TEXT EXTRACTION + JSON BUILDER
# ─────────────────────────────────────────────────────────────────────────────

def clean_text(raw):
    """Strip non-printable characters and normalise whitespace."""
    if not raw:
        return ""
    text = re.sub(r"[^\x20-\x7E\n\r\t]", "", raw)
    text = re.sub(r" +", " ", text)
    return text.strip()


def pdf_to_json(pdf_path, db_id, folder, filename):
    """
    Open the exported PDF with pdfplumber and build the standard JSON structure.
    """
    with pdfplumber.open(pdf_path) as pdf:
        json_data = {
            "source_file":   f"{folder}/{filename}".replace("\\", "/"),
            "document_name": filename,
            "page_count":    len(pdf.pages),
            "metadata": {
                "source_id": db_id,
                "format":    "QuarkXPress 2016",
            },
            "pages": [],
        }

        for i, page in enumerate(pdf.pages):
            raw     = page.extract_text() or ""
            cleaned = clean_text(raw)
            json_data["pages"].append({
                "page_name": str(i + 1),
                "frames": [{
                    "bounds":     [0, 0, 0, 0],
                    "text":       cleaned,
                    "paragraphs": [
                        {"style": "Normal", "text": line}
                        for line in cleaned.split("\n")
                        if line.strip()
                    ],
                }],
            })

    return json_data


# ─────────────────────────────────────────────────────────────────────────────
# PROGRESS TRACKING  (so we can resume after a crash)
# ─────────────────────────────────────────────────────────────────────────────

PROGRESS_FILE = Path(OUTPUT_ROOT) / "progress.json"


def load_progress():
    if PROGRESS_FILE.exists():
        with open(PROGRESS_FILE) as f:
            return set(json.load(f))
    return set()


def save_progress(completed_ids):
    PROGRESS_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(PROGRESS_FILE, "w") as f:
        json.dump(list(completed_ids), f)


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    pdf_dir = Path(OUTPUT_ROOT) / PDF_SUBDIR
    pdf_dir.mkdir(parents=True, exist_ok=True)

    completed_ids = load_progress()
    log.info(f"Resuming — {len(completed_ids)} files already completed")

    conn = db_connect()
    log.info("Database connected")

    rows = db_fetch_pending(conn, limit=BATCH_LIMIT)
    log.info(f"Fetched {len(rows)} pending rows from database")

    # Filter out already-completed (in case DB update failed but file exists)
    rows = [r for r in rows if r[0] not in completed_ids]
    log.info(f"{len(rows)} rows to process after filtering")

    # Make sure QuarkXPress is not already running
    close_quark_completely()

    stats = {"success": 0, "failed": 0, "skipped": 0}

    for idx, (db_id, filename, folder) in enumerate(rows, 1):
        log.info(f"[{idx}/{len(rows)}]  ID {db_id}: {folder}/{filename}")

        # Skip non-Quark files (those with extensions other than .qzp/.qxp)
        suffix = Path(filename).suffix.lower()
        if suffix not in ("", ".qzp", ".qxp"):
            log.info(f"    Skipping — not a Quark file (suffix: '{suffix}')")
            stats["skipped"] += 1
            continue

        source_file = Path(ASSEMBLED_ROOT) / folder / filename
        if not source_file.exists():
            # Try with _9x suffix variants
            for candidate in [
                source_file.parent / (filename + "_9x.qzp"),
                source_file.parent / (filename + "_9x.qxp"),
                source_file.with_suffix(".qzp"),
                source_file.with_suffix(".qxp"),
            ]:
                if candidate.exists():
                    source_file = candidate
                    break
            else:
                log.warning(f"    File not found: {source_file}")
                stats["failed"] += 1
                continue

        pdf_path = pdf_dir / f"{db_id}.pdf"

        try:
            # ── Step 1: Open in QuarkXPress ──────────────────────────────────
            ok = open_file_in_quark(source_file)
            if not ok:
                raise RuntimeError("Failed to open file in QuarkXPress")

            # ── Step 2: Export PDF ───────────────────────────────────────────
            ok = export_pdf(pdf_path)
            if not ok:
                raise RuntimeError("PDF export failed")

            # ── Step 3: Close document ───────────────────────────────────────
            close_document()

            # ── Step 4: Verify PDF was created ───────────────────────────────
            if not pdf_path.exists():
                raise RuntimeError(f"PDF not found after export: {pdf_path}")

            # ── Step 5: Extract text and build JSON ──────────────────────────
            json_data = pdf_to_json(pdf_path, db_id, folder, filename)

            # ── Step 6: Update database ──────────────────────────────────────
            db_update_row(conn, db_id, json.dumps(json_data, ensure_ascii=False))

            completed_ids.add(db_id)
            save_progress(completed_ids)
            stats["success"] += 1
            log.info(f"    ✓ Done  ({json_data['page_count']} pages)")

        except Exception as e:
            log.error(f"    ✗ Failed: {e}")
            stats["failed"] += 1

            # Try to recover: close any open Quark windows and kill if needed
            try:
                pyautogui.hotkey("ctrl", "w")
                time.sleep(1)
                dismiss_save_changes_dialog(timeout=5)
            except Exception:
                pass

            # If Quark seems hung, kill and restart
            if not window_exists("QuarkXPress") or idx % 100 == 0:
                log.info("    Restarting QuarkXPress...")
                close_quark_completely()
                time.sleep(3)

            continue

        # Periodic hard restart every 500 files to prevent memory leaks
        if idx % 500 == 0:
            log.info("── Periodic restart of QuarkXPress ──")
            close_quark_completely()
            time.sleep(5)

    # ── Final cleanup ────────────────────────────────────────────────────────
    close_quark_completely()
    conn.close()

    log.info("═" * 60)
    log.info(f"COMPLETE  success={stats['success']}  "
             f"failed={stats['failed']}  skipped={stats['skipped']}")
    log.info(f"Progress saved to: {PROGRESS_FILE}")
    log.info(f"PDFs saved to:     {pdf_dir}")
    log.info(f"Log saved to:      {LOG_FILE}")


if __name__ == "__main__":
    main()
