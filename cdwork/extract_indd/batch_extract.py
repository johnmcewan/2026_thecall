"""
batch_extract.py

Drives InDesign CC 2023+ via Windows COM automation to extract text from
every .indd file under SOURCE_DIR and write one JSON file per document.

How it works
------------
- Opens InDesign once and keeps it running for the entire batch (fast).
- Reads extract_text.jsx as a template, injects the two file paths as
  JavaScript variables, then calls InDesign's DoScript method directly.
  No command-line arguments, no popup dialogs.
- Resumes automatically if interrupted (progress is logged to a file).

Prerequisites
-------------
    pip install pywin32 tqdm

Then run once to register COM types (in an elevated prompt if needed):
    python -m win32com.client.makepy "Adobe InDesign 2024"
  (Adjust the year to match your installed version.)

Usage
-----
    python batch_extract.py
"""

import os
import sys
import time
import json
import logging
import pythoncom
import win32com.client
from pathlib import Path
from tqdm import tqdm

# ---------------------------------------------------------------------------
# Configuration — edit these before running
# ---------------------------------------------------------------------------

SOURCE_DIR  = r"E:\Callproject\assembled"
OUTPUT_DIR  = r"E:\Callproject\extracted_json"

# Path to the JSX template (same folder as this script by default)
SCRIPT_TEMPLATE = Path(__file__).parent / "extract_text.jsx"

# Seconds to wait for InDesign to process one file before giving up
TIMEOUT_PER_FILE = 120

# Abort the whole run after this many consecutive failures
MAX_CONSECUTIVE_ERRORS = 10

# Resume log — lists .indd paths that have already been processed
PROGRESS_LOG = Path(__file__).parent / "batch_progress.log"

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[
        logging.FileHandler(
            Path(__file__).parent / "batch_extract.log", encoding="utf-8"
        ),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Progress helpers
# ---------------------------------------------------------------------------

def load_progress() -> set:
    if not PROGRESS_LOG.exists():
        return set()
    with PROGRESS_LOG.open("r", encoding="utf-8") as f:
        return set(line.strip() for line in f if line.strip())


def record_progress(path: str) -> None:
    with PROGRESS_LOG.open("a", encoding="utf-8") as f:
        f.write(path + "\n")


# ---------------------------------------------------------------------------
# File discovery
# ---------------------------------------------------------------------------

def find_indd_files(root: str) -> list:
    return sorted(Path(root).rglob("*.indd"))


def derive_output_path(indd: Path, src_root: Path, out_root: Path) -> Path:
    rel = indd.relative_to(src_root)
    out = out_root / rel.with_suffix(".json")
    out.parent.mkdir(parents=True, exist_ok=True)
    return out


# ---------------------------------------------------------------------------
# COM / InDesign helpers
# ---------------------------------------------------------------------------

def get_indesign() -> object:
    """
    Connect to a running InDesign instance or launch a new one.
    Uses the version-independent ProgID so it works with any CC version.
    """
    # Try to connect to an already-running instance first
    for progid in [
        "InDesign.Application.2026",
        "InDesign.Application.2025",
        "InDesign.Application",          # version-independent fallback
    ]:
        try:
            app = win32com.client.GetActiveObject(progid)
            log.info(f"Connected to running InDesign instance ({progid}).")
            return app
        except Exception:
            pass

    # Nothing running — launch it
    for progid in [
        "InDesign.Application.2026",
        "InDesign.Application.2025",
        "InDesign.Application",
    ]:
        try:
            app = win32com.client.Dispatch(progid)
            log.info(f"Launched InDesign ({progid}).")
            # Give InDesign a moment to finish starting up
            time.sleep(5)
            return app
        except Exception:
            pass

    raise RuntimeError(
        "Could not connect to or launch InDesign. "
        "Make sure it is installed and try running:\n"
        '  python -m win32com.client.makepy "Adobe InDesign 2024"'
    )


def build_script(template: str, indd_path: Path, json_path: Path) -> str:
    """
    Prepend the two path variables to the JSX template.
    Backslashes are doubled so they are valid inside JavaScript string literals.
    """
    indd_escaped = str(indd_path).replace("\\", "\\\\")
    json_escaped = str(json_path).replace("\\", "\\\\")
    header = (
        f'var inddPath = "{indd_escaped}";\n'
        f'var jsonPath = "{json_escaped}";\n\n'
    )
    return header + template


def run_script(app, script: str) -> tuple:
    """
    Send the script to InDesign via DoScript and return (ok, message).
    ScriptLanguage 1246973031 = JavaScript (ExtendScript).
    """
    try:
        # DoScript(script, ScriptLanguage, arguments, undoMode)
        app.DoScript(script, 1246973031)

        return True, "OK"
    except Exception as e:
        return False, str(e)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    log.info("=" * 70)
    log.info("InDesign batch text extractor (COM mode) starting")
    log.info(f"  Source : {SOURCE_DIR}")
    log.info(f"  Output : {OUTPUT_DIR}")
    log.info(f"  Script : {SCRIPT_TEMPLATE}")
    log.info("=" * 70)

    if not SCRIPT_TEMPLATE.exists():
        log.error(f"JSX template not found: {SCRIPT_TEMPLATE}")
        sys.exit(1)

    jsx_template = SCRIPT_TEMPLATE.read_text(encoding="utf-8")
    Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)

    # Discover files
    log.info("Scanning for .indd files …")
    all_files = find_indd_files(SOURCE_DIR)
    log.info(f"Found {len(all_files):,} .indd files.")

    done = load_progress()
    pending = [f for f in all_files if str(f) not in done]
    skipped = len(all_files) - len(pending)
    if skipped:
        log.info(f"Skipping {skipped:,} already-processed files.")
    log.info(f"Files to process: {len(pending):,}")

    # Initialise COM on this thread
    pythoncom.CoInitialize()

    try:
        app = get_indesign()
    except RuntimeError as e:
        log.error(str(e))
        sys.exit(1)

    src_root = Path(SOURCE_DIR)
    out_root = Path(OUTPUT_DIR)

    success_count      = 0
    fail_count         = 0
    consecutive_errors = 0
    failed_files       = []

    with tqdm(total=len(pending), unit="file", desc="Extracting") as bar:
        for indd_path in pending:
            json_path = derive_output_path(indd_path, src_root, out_root)
            bar.set_postfix_str(indd_path.name[:40])

            script = build_script(jsx_template, indd_path, json_path)
            ok, msg = run_script(app, script)

            # Secondary check: even if DoScript didn't throw, verify output
            if ok and not json_path.exists():
                ok, msg = False, "JSON output file not created"
            if ok:
                try:
                    with json_path.open("r", encoding="utf-8") as f:
                        json.load(f)
                except json.JSONDecodeError as e:
                    ok, msg = False, f"Malformed JSON output: {e}"

            if ok:
                record_progress(str(indd_path))
                success_count      += 1
                consecutive_errors  = 0
                log.debug(f"OK   {indd_path}")
            else:
                fail_count         += 1
                consecutive_errors += 1
                failed_files.append((str(indd_path), msg))
                log.warning(f"FAIL {indd_path}  —  {msg}")

                # If InDesign has gone away, try to reconnect once
                if consecutive_errors % 5 == 0:
                    log.warning("Multiple consecutive errors — attempting to reconnect to InDesign …")
                    try:
                        pythoncom.CoInitialize()
                        app = get_indesign()
                        log.info("Reconnected.")
                    except Exception as re:
                        log.error(f"Reconnect failed: {re}")

            bar.update(1)

            if consecutive_errors >= MAX_CONSECUTIVE_ERRORS:
                log.error(
                    f"Aborting: {consecutive_errors} consecutive failures. "
                    "Check InDesign is still running and responsive."
                )
                break

    log.info("=" * 70)
    log.info(f"Finished.  Success: {success_count:,}  |  Failed: {fail_count:,}")

    if failed_files:
        fail_log = Path(__file__).parent / "failed_files.log"
        with fail_log.open("w", encoding="utf-8") as f:
            for path, reason in failed_files:
                f.write(f"{path}\t{reason}\n")
        log.info(f"Failed file list written to: {fail_log}")

    pythoncom.CoUninitialize()
    log.info("=" * 70)


if __name__ == "__main__":
    main()
