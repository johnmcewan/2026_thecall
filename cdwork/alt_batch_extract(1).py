"""
alt_batch_extract.py
====================
Reads the 7zip_results.log produced by the previous extraction pass and
retries every ISO that received RESULT: FAILED, using two alternative
approaches that are more tolerant of non-standard Mac-burned UDF discs:

  Method 1 – PowerShell Mount-DiskImage
      Mounts the ISO as a virtual Windows drive and robocopy's all files
      out.  The Windows UDF driver accepts the non-standard anchor layout
      that 7-Zip rejects.

  Method 2 – isoinfo / isoread  (cdrtools / cdrkit)
      Falls back to isoinfo -l to list the ISO directory tree, then
      isoread to extract each file individually.  Works on ISOs the
      Windows driver also cannot mount.

Only ISOs whose previous result was FAILED are retried; already-OK
entries are left alone.  Results go to a new alt_extract_results.log so
the original log is never modified.

Usage
-----
    python alt_batch_extract.py

Edit the CONFIGURATION block below, then run from any Python 3 prompt.
No third-party packages required.

Dependencies
------------
  • PowerShell 5+ (built into Windows 10/11)  — for Method 1
  • isoinfo.exe and isoread.exe from cdrtools  — for Method 2
      Download: https://sourceforge.net/projects/cdrtools/
      or install via Chocolatey:  choco install cdrtools
"""

import os
import re
import shutil
import subprocess
import tempfile
import time

# ── CONFIGURATION ──────────────────────────────────────────────────────────

# The 7zip_results.log written by the previous script
PREVIOUS_LOG = r"E:\c_out_march29\_7zip_recovered\7zip_results.log"

# Root folder where re-extracted files will be written.
# Each ISO gets a sub-folder: {parent_number}_{iso_stem}_alt
DEST_DIR = r"E:\c_out_march29\_alt_recovered"

# Path to isoinfo.exe  (Method 2).  Set to "" to skip Method 2 entirely.
ISOINFO  = r"C:\Program Files\cdrtools\isoinfo.exe"

# Path to isoread.exe  (Method 2).  Set to "" to skip Method 2 entirely.
ISOREAD  = r"C:\Program Files\cdrtools\isoread.exe"

# Timeout in seconds for a single ISO (mount + copy).  Large ISOs may need more.
TIMEOUT = 900   # 15 minutes

# Log file written by this script
RESULT_LOG = os.path.join(DEST_DIR, "alt_extract_results.log")

# ── END CONFIGURATION ───────────────────────────────────────────────────────


# ── LOGGING ────────────────────────────────────────────────────────────────

def log(msg, file=None):
    line = f"{time.strftime('%Y-%m-%d %H:%M:%S')} | {msg}"
    print(line)
    if file:
        file.write(line + "\n")
        file.flush()


# ── INPUT PARSING ──────────────────────────────────────────────────────────

def parse_failed_isos(previous_log_path):
    """
    Read the 7zip_results.log and return a list of (iso_path, pycdlib_err)
    tuples for every ISO whose result was FAILED.

    The log format produced by the previous script looks like:

        2026-03-29 11:37:56 | [ 18/236] Nov 3.iso | pycdlib error: …
        2026-03-29 11:37:56 | [ 18/236] Nov 3.iso | Extracting to: E:\…\130_Nov 3
        2026-03-29 11:37:56 | [ 18/236] Nov 3.iso | RESULT: FAILED (FATAL)
        2026-03-29 11:37:56 |   7z> …7-Zip output…

    We reconstruct the original ISO path from the "Extracting to:" line and
    the 7-Zip output block that contains the real path.
    """
    entries   = []
    # keyed by iso_name -> dict with keys: pycdlib_err, iso_path, result
    current   = {}
    failed_names = set()

    # Patterns
    re_pycdlib   = re.compile(r"\|\s+\[[\s\d/]+\]\s+(.+?)\s+\|\s+pycdlib error:\s+(.+)$")
    re_extracting = re.compile(r"\|\s+\[[\s\d/]+\]\s+(.+?)\s+\|\s+Extracting to:\s+(.+)$")
    re_result    = re.compile(r"\|\s+\[[\s\d/]+\]\s+(.+?)\s+\|\s+RESULT:\s+(\S+)")
    re_7zpath    = re.compile(r"7z>\s+(?:Extracting archive|Path)\s*[=:]\s*(.+\.iso)", re.IGNORECASE)

    with open(previous_log_path, encoding="utf-8", errors="replace") as f:
        for raw in f:
            line = raw.rstrip("\r\n")

            m = re_pycdlib.search(line)
            if m:
                name = m.group(1).strip()
                current.setdefault(name, {})["pycdlib_err"] = m.group(2).strip()
                continue

            m = re_extracting.search(line)
            if m:
                name    = m.group(1).strip()
                out_dir = m.group(2).strip()
                current.setdefault(name, {})["out_dir"] = out_dir
                continue

            m = re_result.search(line)
            if m:
                name   = m.group(1).strip()
                result = m.group(2).strip()
                current.setdefault(name, {})["result"] = result
                if result == "FAILED":
                    failed_names.add(name)
                continue

            # 7-Zip verbose output embedded after a FAILED entry — grab the real path
            m = re_7zpath.search(line)
            if m:
                iso_path = m.group(1).strip()
                iso_name = os.path.basename(iso_path)
                if iso_name in current:
                    current[iso_name]["iso_path"] = iso_path

    for name in failed_names:
        info = current.get(name, {})
        iso_path   = info.get("iso_path", "")
        pycdlib_err = info.get("pycdlib_err", "unknown error")
        if iso_path:
            entries.append((iso_path, pycdlib_err))
        else:
            # iso_path not found in 7-Zip output; record with empty path so
            # the main loop can log the skip clearly
            entries.append(("", f"[{name}] iso path not found in log"))

    return entries


# ── OUTPUT FOLDER ──────────────────────────────────────────────────────────

def make_output_folder(iso_path, dest_root):
    """
    {parent_folder_name}_{iso_stem}_alt
    e.g. G:\\thecall\\done\\130\\Nov 3.iso  ->  130_Nov 3_alt
    """
    parent     = os.path.basename(os.path.dirname(iso_path))
    iso_stem   = os.path.splitext(os.path.basename(iso_path))[0]
    folder     = f"{parent}_{iso_stem}_alt"
    folder     = re.sub(r'[<>:"/\\|?*]', "_", folder)
    return os.path.join(dest_root, folder)


# ── METHOD 1 — PowerShell Mount-DiskImage ──────────────────────────────────

_PS = r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"

def _run_ps(script, timeout=60):
    """Run a PowerShell one-liner and return (returncode, combined_output)."""
    try:
        r = subprocess.run(
            [_PS, "-NoProfile", "-NonInteractive", "-Command", script],
            capture_output=True, text=True,
            encoding="utf-8", errors="replace",
            timeout=timeout,
        )
        return r.returncode, (r.stdout + r.stderr).strip()
    except subprocess.TimeoutExpired:
        return -1, f"PowerShell timeout after {timeout}s"
    except Exception as e:
        return -2, str(e)


def extract_via_mount(iso_path, out_dir, timeout=TIMEOUT):
    """
    Mount the ISO with PowerShell, robocopy everything, dismount.

    Returns (success: bool, message: str, files_copied: int).
    """
    # ── Mount ──────────────────────────────────────────────────────────────
    mount_script = (
        f"$img = Mount-DiskImage -ImagePath '{iso_path}' -PassThru; "
        f"($img | Get-Volume).DriveLetter"
    )
    rc, out = _run_ps(mount_script, timeout=60)
    if rc != 0 or not out.strip():
        return False, f"Mount-DiskImage failed (rc={rc}): {out}", 0

    drive_letter = out.strip().splitlines()[-1].strip().rstrip(":")
    if not drive_letter or len(drive_letter) != 1:
        # Try to dismount just in case, then bail
        _run_ps(f"Dismount-DiskImage -ImagePath '{iso_path}'", timeout=30)
        return False, f"Could not determine drive letter from output: {out!r}", 0

    src = f"{drive_letter}:\\"

    # ── Copy ───────────────────────────────────────────────────────────────
    os.makedirs(out_dir, exist_ok=True)
    # /E  = all sub-dirs including empty  /COPYALL = preserve metadata
    # /R:1 /W:1 = minimal retry on locked files  /NP = no progress %
    # /NFL /NDL = suppress file/dir lists from stdout (keeps log tidy)
    robocopy_cmd = [
        "robocopy", src, out_dir,
        "/E", "/COPYALL", "/R:1", "/W:1", "/NP", "/NFL", "/NDL",
    ]
    try:
        r = subprocess.run(
            robocopy_cmd,
            capture_output=True, text=True,
            encoding="utf-8", errors="replace",
            timeout=timeout,
        )
        # robocopy exit codes: 0=no files, 1=files copied, 2=extra files,
        # 3=1+2, 4=mismatches, 8=fail, 16=fatal.  0-7 are all non-fatal.
        copy_ok     = r.returncode < 8
        copy_output = (r.stdout + r.stderr).strip()

        # Parse how many files were copied from robocopy summary line
        files_copied = 0
        m = re.search(r"Files\s*:\s*(\d+)", copy_output)
        if m:
            files_copied = int(m.group(1))

    except subprocess.TimeoutExpired:
        copy_ok, copy_output, files_copied = False, f"robocopy timeout after {timeout}s", 0
    except Exception as e:
        copy_ok, copy_output, files_copied = False, str(e), 0

    # ── Dismount ───────────────────────────────────────────────────────────
    _run_ps(f"Dismount-DiskImage -ImagePath '{iso_path}'", timeout=30)

    if not copy_ok:
        return False, f"robocopy failed (rc={r.returncode}): {copy_output}", 0

    return True, f"robocopy rc={r.returncode}, files copied: {files_copied}", files_copied


# ── METHOD 2 — isoinfo / isoread (cdrtools) ────────────────────────────────

def _isoinfo_list(isoinfo_exe, iso_path, timeout=120):
    """
    Run  isoinfo -l -i <iso>  and return the raw output string,
    or raise RuntimeError on failure.
    """
    r = subprocess.run(
        [isoinfo_exe, "-l", "-i", iso_path],
        capture_output=True, text=True,
        encoding="utf-8", errors="replace",
        timeout=timeout,
    )
    if r.returncode not in (0, 1):   # isoinfo exits 1 on minor issues
        raise RuntimeError(f"isoinfo exited {r.returncode}: {r.stderr.strip()}")
    return r.stdout


def _parse_isoinfo_listing(listing):
    """
    Parse the isoinfo -l output and return a list of ISO 9660 path strings
    like  /DIR/SUBDIR/FILENAME.EXT;1
    isoinfo prints directory headers as:
        Directory listing of /DIR/SUBDIR/
    and file lines as:
        -rwxr-xr-x   1    0    0    12345  Jan  1 2007 FILENAME.EXT;1
    """
    paths    = []
    cur_dir  = "/"
    re_dir   = re.compile(r"^Directory listing of (.+)$")
    re_file  = re.compile(r"^[-d].{9}\s+\d+\s+\d+\s+\d+\s+\d+\s+\S+\s+\d+\s+\d{4}\s+(.+)$")

    for line in listing.splitlines():
        line = line.rstrip()
        m = re_dir.match(line)
        if m:
            cur_dir = m.group(1).rstrip("/")
            continue
        m = re_file.match(line)
        if m:
            fname = m.group(1).strip()
            if fname in (".", ".."):
                continue
            # Drop the version number suffix (;1) for the destination path
            iso_path = f"{cur_dir}/{fname}" if cur_dir != "/" else f"/{fname}"
            paths.append(iso_path)

    return paths


def _isoread_file(isoread_exe, iso_path, iso_file_path, dest_file, timeout=120):
    """
    Extract a single file from the ISO using isoread.
    iso_file_path is the full path inside the ISO, e.g. /DIR/FILE.JPG;1
    """
    os.makedirs(os.path.dirname(dest_file), exist_ok=True)
    with open(dest_file, "wb") as fh:
        r = subprocess.run(
            [isoread_exe, "-i", iso_path, "-x", iso_file_path],
            stdout=fh,
            stderr=subprocess.PIPE,
            timeout=timeout,
        )
    if r.returncode not in (0, 1):
        os.remove(dest_file)
        raise RuntimeError(
            f"isoread exited {r.returncode}: {r.stderr.decode('utf-8', errors='replace').strip()}"
        )


def extract_via_isoinfo(isoinfo_exe, isoread_exe, iso_path, out_dir, timeout=TIMEOUT):
    """
    List and extract an ISO with isoinfo + isoread.

    Returns (success: bool, message: str, files_copied: int).
    """
    if not os.path.isfile(isoinfo_exe):
        return False, f"isoinfo not found: {isoinfo_exe}", 0
    if not os.path.isfile(isoread_exe):
        return False, f"isoread not found: {isoread_exe}", 0

    try:
        listing = _isoinfo_list(isoinfo_exe, iso_path, timeout=min(120, timeout))
    except Exception as e:
        return False, f"isoinfo listing failed: {e}", 0

    file_paths = _parse_isoinfo_listing(listing)
    if not file_paths:
        return False, "isoinfo returned no files", 0

    os.makedirs(out_dir, exist_ok=True)
    errors    = []
    extracted = 0
    deadline  = time.time() + timeout

    for iso_file_path in file_paths:
        if time.time() > deadline:
            errors.append("timeout reached mid-extraction")
            break

        # Strip version suffix for the on-disk filename
        dest_rel  = iso_file_path.lstrip("/").replace(";1", "")
        dest_file = os.path.join(out_dir, dest_rel.replace("/", os.sep))

        # Skip directory entries (end with /)
        if iso_file_path.endswith("/"):
            continue

        try:
            _isoread_file(isoread_exe, iso_path, iso_file_path, dest_file,
                          timeout=min(300, int(deadline - time.time()) + 1))
            extracted += 1
        except Exception as e:
            errors.append(f"{iso_file_path}: {e}")

    if errors:
        err_summary = f"{len(errors)} error(s): " + "; ".join(errors[:3])
        if len(errors) > 3:
            err_summary += f" … (+{len(errors)-3} more)"
        if extracted == 0:
            return False, err_summary, 0
        return True, f"partial — {extracted} files ok, {err_summary}", extracted

    return True, f"{extracted} files extracted", extracted


# ── MAIN ───────────────────────────────────────────────────────────────────

def main():
    # ── Pre-flight checks ──────────────────────────────────────────────────
    if not os.path.isfile(PREVIOUS_LOG):
        print(f"ERROR: Cannot find previous results log: {PREVIOUS_LOG}")
        return

    ps_available = os.path.isfile(_PS)
    isoinfo_available = (
        os.path.isfile(ISOINFO) and os.path.isfile(ISOREAD)
    )

    if not ps_available:
        print("WARNING: PowerShell not found — Method 1 (Mount-DiskImage) will be skipped.")
    if not isoinfo_available:
        print(
            "WARNING: isoinfo/isoread not found — Method 2 (cdrtools) will be skipped.\n"
            f"  Expected: {ISOINFO}\n"
            f"            {ISOREAD}\n"
            "  Install via: choco install cdrtools"
        )
    if not ps_available and not isoinfo_available:
        print("ERROR: No extraction methods available.  Aborting.")
        return

    os.makedirs(DEST_DIR, exist_ok=True)

    # ── Parse failed ISOs from previous log ───────────────────────────────
    print(f"Reading failed ISOs from: {PREVIOUS_LOG}")
    entries = parse_failed_isos(PREVIOUS_LOG)
    if not entries:
        print("No FAILED entries found in the previous log.  Nothing to do.")
        return
    print(f"Found {len(entries)} FAILED ISOs to retry.\n")

    counts = {"ok": 0, "partial": 0, "failed": 0, "skipped": 0}

    with open(RESULT_LOG, "a", encoding="utf-8") as logfile:
        log(f"Starting alt-extraction batch — {len(entries)} ISOs", logfile)
        log(f"Output root  : {DEST_DIR}", logfile)
        log(f"Method 1     : {'Mount-DiskImage (PowerShell)' if ps_available else 'DISABLED'}", logfile)
        log(f"Method 2     : {'isoinfo + isoread (cdrtools)' if isoinfo_available else 'DISABLED'}", logfile)
        log("-" * 72, logfile)

        for i, (iso_path, pycdlib_err) in enumerate(entries, 1):
            iso_name = os.path.basename(iso_path) if iso_path else "(unknown)"
            prefix   = f"[{i:>3}/{len(entries)}] {iso_name}"

            # ── Skip: no path recovered from log ──────────────────────────
            if not iso_path:
                log(f"{prefix} | SKIP — {pycdlib_err}", logfile)
                counts["skipped"] += 1
                log("", logfile)
                continue

            # ── Skip: ISO file missing on disk ────────────────────────────
            if not os.path.isfile(iso_path):
                log(f"{prefix} | SKIP — ISO not found on disk: {iso_path}", logfile)
                counts["skipped"] += 1
                log("", logfile)
                continue

            out_dir = make_output_folder(iso_path, DEST_DIR)

            # ── Skip: already successfully extracted in a previous alt run ─
            if os.path.isdir(out_dir) and os.listdir(out_dir):
                log(f"{prefix} | SKIP — already extracted to {out_dir}", logfile)
                counts["skipped"] += 1
                log("", logfile)
                continue

            log(f"{prefix} | pycdlib error : {pycdlib_err}", logfile)
            log(f"{prefix} | Extracting to : {out_dir}", logfile)

            method_used = None
            success     = False
            partial     = False
            message     = ""
            files       = 0

            # ── Method 1: Mount-DiskImage ──────────────────────────────────
            if ps_available:
                log(f"{prefix} | Trying Method 1: Mount-DiskImage ...", logfile)
                success, message, files = extract_via_mount(iso_path, out_dir)
                method_used = "Mount-DiskImage"
                if success:
                    partial = "partial" in message.lower()

            # ── Method 2: isoinfo + isoread ────────────────────────────────
            if not success and isoinfo_available:
                log(f"{prefix} | Method 1 failed ({message}) — trying Method 2: isoinfo/isoread ...", logfile)
                success, message, files = extract_via_isoinfo(
                    ISOINFO, ISOREAD, iso_path, out_dir
                )
                method_used = "isoinfo/isoread"
                if success:
                    partial = "partial" in message.lower()

            # ── Record result ──────────────────────────────────────────────
            if success and not partial:
                log(f"{prefix} | RESULT: OK  [{method_used}]  {message}", logfile)
                counts["ok"] += 1
            elif success and partial:
                log(f"{prefix} | RESULT: PARTIAL  [{method_used}]  {message}", logfile)
                counts["partial"] += 1
            else:
                log(f"{prefix} | RESULT: FAILED  (all methods exhausted)", logfile)
                log(f"{prefix} |   Last error: {message}", logfile)
                counts["failed"] += 1

            log("", logfile)

        log("-" * 72, logfile)
        log(
            f"Done.  OK={counts['ok']}  PARTIAL={counts['partial']}  "
            f"FAILED={counts['failed']}  SKIPPED={counts['skipped']}",
            logfile,
        )

    print(f"\nResults written to: {RESULT_LOG}")


if __name__ == "__main__":
    main()
