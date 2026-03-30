"""
7zip_batch_extract.py
=====================
Reads the pycdlib errors.log, finds every OPEN_FAIL entry, and attempts
to extract each ISO with 7-Zip.  Results are written to a companion log
so you can see exactly what was recovered.

Usage
-----
    python 7zip_batch_extract.py

Configure the four variables in the CONFIGURATION block below, then run
from any Python 3 prompt (no extra packages needed).
"""

import os
import re
import subprocess
import time

# ── CONFIGURATION ──────────────────────────────────────────────────────────

# Full path to the errors.log produced by cdextract
ERROR_LOG = r"E:\c_out_march29\errors.log"

# Root folder where extracted files will be written.
# A sub-folder is created for each ISO, named {parent_number}_{iso_name}.
DEST_DIR = r"E:\c_out_march29\_7zip_recovered"

# 7-Zip executable
SEVENZIP = r"C:\Program Files\7-Zip\7z.exe"

# Log file written by this script
RESULT_LOG = os.path.join(DEST_DIR, "7zip_results.log")

# ── END CONFIGURATION ───────────────────────────────────────────────────────


# 7-Zip exit codes
EXIT_CODES = {
    0: "OK",
    1: "WARNING",   # non-fatal errors (e.g. some files locked) — data still extracted
    2: "FATAL",     # 7-Zip could not open or read the archive
    7: "BAD_ARGS",
    8: "OUT_OF_MEM",
    255: "ABORTED",
}


def log(msg, file=None):
    line = f"{time.strftime('%Y-%m-%d %H:%M:%S')} | {msg}"
    print(line)
    if file:
        file.write(line + "\n")
        file.flush()


def parse_open_fails(error_log_path):
    """
    Return a list of (iso_full_path, error_message) tuples for every
    OPEN_FAIL line in the pycdlib error log.
    """
    entries = []
    pattern = re.compile(r"OPEN_FAIL \| (.+?) \| (.+)$")
    with open(error_log_path, encoding="utf-8", errors="replace") as f:
        for line in f:
            line = line.rstrip("\r\n")
            if "| OPEN_FAIL |" not in line:
                continue
            m = pattern.search(line)
            if m:
                entries.append((m.group(1).strip(), m.group(2).strip()))
    return entries


def make_output_folder(iso_path, dest_root):
    """
    Mirror the naming convention used by cdextract:
        {parent_folder_name}_{iso_name_without_extension}
    e.g.  G:\\thecall\\done\\130\\Nov 3.iso  ->  130_Nov 3
    """
    parent = os.path.basename(os.path.dirname(iso_path))
    iso_stem = os.path.splitext(os.path.basename(iso_path))[0]
    folder_name = f"{parent}_{iso_stem}"
    # Remove characters illegal on Windows
    folder_name = re.sub(r'[<>:"/\\|?*]', "_", folder_name)
    return os.path.join(dest_root, folder_name)


def extract_with_7zip(sevenzip, iso_path, out_dir):
    """
    Run 7-Zip and return (exit_code, stderr_text).
    Uses the long-path prefix on the output directory to handle paths > 260 chars.
    """
    # The \\?\ prefix lets Windows accept paths longer than MAX_PATH
    safe_out = "\\\\?\\" + os.path.abspath(out_dir)

    cmd = [
        sevenzip,
        "x",            # extract with full paths
        iso_path,
        f"-o{safe_out}",
        "-y",           # yes to all prompts
        "-bb0",         # minimal output (errors only)
        "-bd",          # no progress bar (cleaner logs)
        "-scsUTF-8",    # treat filenames as UTF-8
    ]

    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=600,    # 10-minute timeout per ISO
        )
        return result.returncode, result.stdout + result.stderr
    except subprocess.TimeoutExpired:
        return -1, "TIMEOUT after 600s"
    except Exception as e:
        return -2, str(e)


def main():
    # Validate configuration
    if not os.path.isfile(ERROR_LOG):
        print(f"ERROR: Cannot find error log: {ERROR_LOG}")
        return
    if not os.path.isfile(SEVENZIP):
        print(f"ERROR: Cannot find 7-Zip: {SEVENZIP}")
        return

    os.makedirs(DEST_DIR, exist_ok=True)

    entries = parse_open_fails(ERROR_LOG)
    if not entries:
        print("No OPEN_FAIL entries found in the log.")
        return

    print(f"Found {len(entries)} OPEN_FAIL ISOs to attempt.\n")

    counts = {"ok": 0, "warning": 0, "failed": 0, "skipped": 0}

    with open(RESULT_LOG, "a", encoding="utf-8") as logfile:
        log(f"Starting batch — {len(entries)} ISOs", logfile)
        log(f"Output root : {DEST_DIR}", logfile)
        log(f"7-Zip       : {SEVENZIP}", logfile)
        log("-" * 72, logfile)

        for i, (iso_path, pycdlib_err) in enumerate(entries, 1):
            iso_name = os.path.basename(iso_path)
            out_dir  = make_output_folder(iso_path, DEST_DIR)

            prefix = f"[{i:>3}/{len(entries)}] {iso_name}"

            # Skip if already extracted (folder exists and is non-empty)
            if os.path.isdir(out_dir) and os.listdir(out_dir):
                log(f"{prefix} | SKIP (already extracted) -> {out_dir}", logfile)
                counts["skipped"] += 1
                continue

            if not os.path.isfile(iso_path):
                log(f"{prefix} | SKIP (ISO not found on disk: {iso_path})", logfile)
                counts["skipped"] += 1
                continue

            log(f"{prefix} | pycdlib error: {pycdlib_err}", logfile)
            log(f"{prefix} | Extracting to: {out_dir}", logfile)

            os.makedirs(out_dir, exist_ok=True)
            rc, output = extract_with_7zip(SEVENZIP, iso_path, out_dir)

            status = EXIT_CODES.get(rc, f"UNKNOWN({rc})")

            if rc == 0:
                log(f"{prefix} | RESULT: OK", logfile)
                counts["ok"] += 1
            elif rc == 1:
                # Warnings mean some files may have been skipped but the rest
                # extracted fine — still worth keeping
                log(f"{prefix} | RESULT: WARNING (partial — check output)", logfile)
                counts["warning"] += 1
            else:
                log(f"{prefix} | RESULT: FAILED ({status})", logfile)
                # Log 7-Zip's own output for diagnosis
                for line in output.strip().splitlines():
                    log(f"  7z> {line}", logfile)
                counts["failed"] += 1

            log("", logfile)   # blank line between entries

        log("-" * 72, logfile)
        log(
            f"Done.  OK={counts['ok']}  WARNING={counts['warning']}  "
            f"FAILED={counts['failed']}  SKIPPED={counts['skipped']}",
            logfile,
        )

    print(f"\nResults written to: {RESULT_LOG}")


if __name__ == "__main__":
    main()
