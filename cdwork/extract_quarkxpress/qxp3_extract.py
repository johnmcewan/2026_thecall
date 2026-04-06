#!/usr/bin/env python3
"""
qxp3_extract.py — Batch text extractor for QuarkXPress 3 documents.

QuarkXPress 3 stores text as MacRoman-encoded bytes embedded in a proprietary
binary layout format (magic: MMXPR3, big-endian).  Text chains are separated
from binary layout/image data by control characters (0x00–0x1F, excluding TAB
0x09 and Mac CR 0x0D which are valid text).

This script:
  1. Locates all story text chains in a QXP3 file.
  2. Filters out binary-garbage blocks (detected via MacRoman high-byte density).
  3. Classifies each line as prose, tabular (predictions grid), header, or
     photo-caption production markup.
  4. Writes a UTF-8 .txt file (clean prose + table annotations) and a
     structured .json sidecar ready for NER / entity-linking pipelines.

Usage
-----
  # Single file
  python3 qxp3_extract.py path/to/file

  # Batch: whole directory (recurses, auto-detects QXP3 by magic bytes)
  python3 qxp3_extract.py path/to/archive/

  # Custom output directory
  python3 qxp3_extract.py path/to/archive/ --out path/to/output/

  # Quiet (no per-file progress lines)
  python3 qxp3_extract.py path/to/archive/ --quiet

Output per file
---------------
  <name>.txt   Clean UTF-8 plain text; tables rendered as TSV-style lines.
  <name>.json  Structured extraction with blocks[], each block having:
                 type       : "prose" | "table" | "header" | "photo_caption"
                              | "fragment" | "title"
                 text       : raw cleaned text of the block
                 lines      : list of individual lines (prose blocks)
                 rows       : list of cell-lists (table blocks)
                 persons    : list of person names found in photo captions
               Plus top-level fields:
                 source, file_size, block_count, full_prose_text

NER note
--------
For downstream NER, use the `full_prose_text` field from the JSON, which
concatenates all prose + header + title blocks.  Table rows contain named
entities too (person names as column headers, team/score mentions in cells);
iterate `blocks` where type == "table" to process them separately.
"""

import argparse
import json
import re
import sys
from pathlib import Path


# ── Constants ────────────────────────────────────────────────────────────────

QXP3_MAGIC = b"MMXPR3"

# Split only on runs of 2+ control characters.
# Single bytes (especially 0x00) frequently appear as inline formatting tags
# *inside* a text run; splitting on them creates sentence fragments.
# Keeping TAB (0x09) and Mac CR (0x0D) as text characters.
CTRL_SPLIT_RE = re.compile(r"[\x00-\x08\x0a-\x0c\x0e-\x1f]{2,}")

# A block is garbage if its MacRoman high-byte (0x80–0xFF) density exceeds
# this threshold.  Real text has near-zero high-byte density (only isolated
# smart quotes, em-dashes, ellipses, etc.).  Binary halftone / image data has
# 30–60%.
MAX_HIGH_BYTE_RATIO = 0.08   # 8%

# Minimum word count and alpha ratio to even consider a run as a text block.
MIN_WORD_COUNT = 5
MIN_ALPHA_RATIO = 0.28
MIN_RUN_LENGTH  = 35         # chars after stripping

# Column separator: 5+ spaces denotes a tabular column boundary.
COL_SEP_RE = re.compile(r" {5,}")

# Photo/image caption marker used by QXP production workflow.
PHOTO_RE = re.compile(r"(?:CALL\s+PHOTOS?|Big\s+Boy)\s*:", re.IGNORECASE)

# MacRoman typography characters to normalise to ASCII equivalents.
MACROMAN_NORMALIZE = str.maketrans({
    "\u2018": "'",   # left single quotation mark  (0xD4)
    "\u2019": "'",   # right single quotation mark (0xD5)
    "\u201C": '"',   # left double quotation mark  (0xD2)
    "\u201D": '"',   # right double quotation mark (0xD3)
    "\u2014": "--",  # em dash                     (0xD0)
    "\u2013": "-",   # en dash                     (0xD1)
    "\u2026": "...", # ellipsis                    (0xC9)
    "\u00B4": "'",   # acute accent used as apostrophe
    "\u0060": "'",   # grave accent used as apostrophe
})


# ── File-level helpers ────────────────────────────────────────────────────────

def is_qxp3(path: Path) -> bool:
    try:
        header = path.read_bytes()[:16]
    except OSError:
        return False
    return QXP3_MAGIC in header


def strip_binary_suffix(s: str) -> str:
    """
    Truncate a decoded run at the point where real text ends.

    QXP3 appends paragraph-attribute records (kerning tables, style tags,
    etc.) directly after the last text character in a story chain.  These
    records start with byte sequences like 0x03, 0x7F, 0xFF that are
    decoded as MacRoman characters but are not text.  Because our
    control-char split only breaks on *two or more* consecutive control
    bytes, single-byte tags survive into the run.

    Strategy: work line by line (text is already CR→LF normalised).  The
    last line that has meaningful word density is the last real line.
    Within that line, cut at the first control byte (0x00–0x1F excluding
    LF, or 0x7F DEL which QXP uses as a tag marker).
    """
    lines = s.split("\n")
    last_good = -1
    for i, line in enumerate(lines):
        # Evaluate word density after removing obvious tag bytes
        clean = re.sub(r"[\x00-\x08\x0b-\x0c\x0e-\x1f\x7f]", "", line)
        words = re.findall(r"[a-zA-Z]{3,}", clean)
        alpha = sum(c.isalpha() and ord(c) < 0x80 for c in clean)
        if words and (alpha / max(len(clean), 1) > 0.30):
            last_good = i

    if last_good < 0:
        return s.strip()

    result_lines = lines[:last_good]
    last_line = lines[last_good]
    # On the last good line, cut at the first control/tag byte
    m = re.search(r"[\x00-\x08\x0b-\x0c\x0e-\x1f\x7f]", last_line)
    if m:
        last_line = last_line[: m.start()].rstrip()
    result_lines.append(last_line)
    return "\n".join(result_lines).strip()


def high_byte_ratio(s: str) -> float:
    """Fraction of characters with ord > 127 (MacRoman high bytes)."""
    if not s:
        return 0.0
    return sum(1 for c in s if ord(c) > 127) / len(s)


def max_high_byte_run(s: str) -> int:
    """Longest consecutive run of MacRoman high-byte characters."""
    best = cur = 0
    for c in s:
        if ord(c) > 127:
            cur += 1
            if cur > best:
                best = cur
        else:
            cur = 0
    return best


def is_garbage(block: str) -> bool:
    """
    Return True if the block is not real story text.
    Catches three kinds of junk:
      1. MacRoman high-byte–dense blocks (halftone/image data decoded as text).
      2. Halftone bitmap data with very low ASCII character diversity.
      3. PostScript font name registry entries (metadata, not story text).
    """
    # ── Test 1: MacRoman high-byte density ───────────────────────────────────
    # Check the HEAD (first 120 chars) — legitimate blocks have near-zero high
    # bytes in their text-bearing portion; binary-image blocks have high bytes
    # throughout.  We deliberately avoid checking the tail because legitimate
    # story blocks often have a trailing binary tag suffix (kerning tables,
    # paragraph attributes) that inflates the tail ratio.
    head = block[:120]
    if high_byte_ratio(head) > MAX_HIGH_BYTE_RATIO:
        return True
    if max_high_byte_run(block) > 3:
        return True

    # ── Test 2: Halftone / dithered bitmap data ───────────────────────────────
    # These blocks decode to sequences of {f, w, U, D, 3, ", space} — the
    # greyscale byte values of QXP's in-memory halftone representation.
    # Real text uses 20+ distinct ASCII chars; halftone uses < 15.
    ascii_chars = set(c for c in block if 0x20 <= ord(c) <= 0x7E)
    if len(ascii_chars) < 15:
        return True
    halftone_chars = sum(1 for c in block if c in 'fwUD3" \t\n')
    if len(block) and halftone_chars / len(block) > 0.75:
        return True

    # ── Test 3: PostScript font name registry ────────────────────────────────
    # Font records look like: "Souvenir-Demi\x14Souvenir-LightItalic\x13..."
    # They pass alpha-ratio and word-count tests but have no prose sentence structure.
    clean = re.sub(r"[\x00-\x1f]", " ", block)
    tokens = clean.split()
    if tokens:
        ps_like = sum(1 for t in tokens if "-" in t and len(t) > 5)
        if ps_like / len(tokens) > 0.40:
            return True

    return False


# ── Line-level classification ─────────────────────────────────────────────────

def classify_line(line: str) -> str:
    """
    Return one of: 'empty', 'tabular', 'header', 'photo_caption', 'prose'.
    """
    stripped = line.strip()
    if not stripped:
        return "empty"

    # Photo/image caption production markup — also catch lines where the
    # leading "PHOTOS" prefix was truncated by a binary tag.
    clean = re.sub(r"^[^\w:]+", "", stripped)  # strip leading junk chars
    if PHOTO_RE.search(clean) or re.search(r"\d-col\s+[A-Z]", clean, re.IGNORECASE):
        return "photo_caption"

    # Split on 5+ spaces to find columnar cells
    cells = [c.strip() for c in COL_SEP_RE.split(stripped) if c.strip()]

    # Tabular: 3+ substantial cells
    if len(cells) >= 3:
        substantial = sum(1 for c in cells if len(c) >= 2)
        if substantial >= 3:
            return "tabular"

    # Tabular: exactly 2 cells both long enough to be real content
    # (e.g. partial prediction rows like "Chiefs by 10   Chiefs b")
    if len(cells) == 2 and all(len(c) > 3 for c in cells):
        return "tabular"

    # Centered single-item header (significant padding on both sides, short)
    if len(cells) == 1:
        leading  = len(line) - len(line.lstrip())
        trailing = len(line) - len(line.rstrip())
        if leading > 8 and trailing > 8 and len(stripped) < 80:
            return "header"

    return "prose"


def clean_caption_line(line: str) -> str:
    """Strip leading binary-rendered MacRoman junk chars from caption lines."""
    # Remove leading non-ASCII and non-alphanumeric chars (e.g. Ω, ¿, √, ∆, digits)
    return re.sub(r"^[^\w:]+", "", line).strip()


def extract_caption_person(line: str) -> str | None:
    """Extract the person name from a QXP photo caption line."""
    clean = clean_caption_line(line)
    parts = clean.split(":")
    if not parts:
        return None
    last = parts[-1].strip()
    # Strip column spec like "1-col " or "1-COL "
    name = re.sub(r"^\d+-col\s+", "", last, flags=re.IGNORECASE).strip()
    # Must look like a real name: Title Case, ALLCAPS, or mixed
    if re.search(
        r"[A-Z][a-z]+ [A-Z][a-z]+"      # Title Case: Jim Nunnelly
        r"|[A-Z]{2,} [A-Z][a-z]+"        # ALLCAPS first: EMANUEL Cleaver
        r"|[A-Z][a-z]+ [A-Z]{2,}"        # ALLCAPS last: Emanuel CLEAVER
        r"|[A-Z]{2,} [A-Z]{2,}",          # All allcaps: EMANUEL CLEAVER
        name
    ):
        return name
    return None


# ── Block extraction ──────────────────────────────────────────────────────────

def extract_raw_blocks(data: bytes) -> list[str]:
    """
    Split the raw file bytes into candidate text runs and return those that
    pass the garbage and content-density filters.
    """
    full = data.decode("mac_roman", errors="replace")
    runs = CTRL_SPLIT_RE.split(full)

    blocks: list[str] = []
    for run in runs:
        words = re.findall(r"[a-zA-Z]{3,}", run)
        if len(words) < MIN_WORD_COUNT:
            continue

        alpha = sum(c.isalpha() for c in run)
        if not run or alpha / len(run) < MIN_ALPHA_RATIO:
            continue

        # Normalise line endings and clean up whitespace
        cleaned = run.replace("\r", "\n")
        cleaned = re.sub(r"\n{4,}", "\n\n", cleaned)
        cleaned = re.sub(r"[ \t]+\n", "\n", cleaned)
        cleaned = cleaned.strip()

        if len(cleaned) < MIN_RUN_LENGTH:
            continue

        # Strip binary tag suffix that QXP3 appends to every story chain.
        cleaned = strip_binary_suffix(cleaned)

        if not cleaned.strip() or len(cleaned) < MIN_RUN_LENGTH:
            continue

        if is_garbage(cleaned):
            continue

        # Normalise MacRoman typography to plain ASCII
        cleaned = cleaned.translate(MACROMAN_NORMALIZE)

        blocks.append(cleaned)

    return blocks


# ── Block structuring ─────────────────────────────────────────────────────────

def is_prose_padding(cells: list[str]) -> bool:
    """
    Return True when a 2-cell 'tabular' line is actually a prose sentence
    split by large internal whitespace (editorial layout padding) rather than
    a real column boundary.

    Real 2-cell table rows begin with a proper noun or short label
    (e.g. "Chiefs by 10", "Lloyd Greenfield").
    Prose-padding rows begin with a lowercase word or a common verb/preposition
    (e.g. "led by", "just wants the Chiefs to win those").
    """
    if len(cells) != 2:
        return False
    first_word = cells[0].split()[0] if cells[0].split() else ""
    if not first_word:
        return False
    # Starts lowercase → sentence continuation
    if first_word[0].islower():
        return True
    # Starts with a common sentence-opening verb or preposition
    prose_verbs = re.compile(
        r"^(led|just|been|have|has|had|will|would|could|should|may|might|want|said|told)\b",
        re.IGNORECASE,
    )
    if prose_verbs.match(cells[0]):
        return True
    return False


def parse_table_rows(tabular_lines: list[str]) -> list[list[str]]:
    """Convert a set of tabular lines into a list of cell-lists."""
    rows = []
    for line in tabular_lines:
        cells = [c.strip() for c in COL_SEP_RE.split(line.strip()) if c.strip()]
        if cells:
            # Reclassify 2-cell rows that look like prose with padding
            if is_prose_padding(cells):
                # Rejoin as a single prose string and skip this row
                continue
            rows.append(cells)
    return rows


def structure_block(raw: str) -> dict:
    """
    Parse a raw text block into a structured dict with type classification.
    A single raw block may contain mixed prose and tabular sections, which are
    split into sub-blocks here.
    """
    lines = raw.split("\n")
    segments: list[tuple[str, list[str]]] = []  # (type, lines)
    cur_type = None
    cur_lines: list[str] = []

    for line in lines:
        ltype = classify_line(line)
        # Merge 'empty' into current segment
        if ltype == "empty":
            cur_lines.append(line)
            continue
        if ltype != cur_type:
            if cur_type and cur_lines:
                segments.append((cur_type, cur_lines))
            cur_type = ltype
            cur_lines = [line]
        else:
            cur_lines.append(line)

    if cur_type and cur_lines:
        segments.append((cur_type, cur_lines))

    # Convert segments to structured blocks
    result_blocks = []
    for stype, slines in segments:
        non_empty = [l for l in slines if l.strip()]
        if not non_empty:
            continue

        if stype == "tabular":
            rows = parse_table_rows(non_empty)
            if rows:
                result_blocks.append({
                    "type": "table",
                    "text": "\n".join("\t".join(row) for row in rows),
                    "rows": rows,
                })

        elif stype == "photo_caption":
            persons = []
            cleaned_lines = [clean_caption_line(l) for l in non_empty]
            for l in cleaned_lines:
                p = extract_caption_person(l)
                if p:
                    persons.append(p)
            result_blocks.append({
                "type": "photo_caption",
                "text": "\n".join(cleaned_lines),
                "lines": cleaned_lines,
                "persons": persons,
            })

        elif stype == "header":
            result_blocks.append({
                "type": "header",
                "text": "\n".join(l.strip() for l in non_empty),
                "lines": [l.strip() for l in non_empty],
            })

        else:  # prose
            prose_text = "\n".join(non_empty)
            # Normalise leading tabs (QXP paragraph indents) to spaces
            prose_text = re.sub(r"^\t+", "    ", prose_text, flags=re.MULTILINE)
            prose_text = prose_text.strip()
            if not prose_text:
                continue
            result_blocks.append({
                "type": "prose",
                "text": prose_text,
                "lines": non_empty,
            })

    return result_blocks


def is_title_block(block: dict) -> bool:
    """Heuristically identify the document title block."""
    if block["type"] != "prose":
        return False
    text = block["text"]
    line_count = len(block.get("lines", []))
    word_count = len(re.findall(r"\w+", text))
    # Title blocks: short, high proportion of capitalised words
    if line_count <= 6 and word_count <= 30:
        caps = len(re.findall(r"\b[A-Z][a-z]+\b", text))
        if caps / max(word_count, 1) > 0.35:
            return True
    return False


def assemble_document(raw_blocks: list[str]) -> list[dict]:
    """
    Convert all raw text blocks into a flat list of structured block dicts.
    The first qualifying prose block is promoted to type 'title'.
    """
    all_blocks: list[dict] = []
    title_assigned = False

    for raw in raw_blocks:
        sub = structure_block(raw)
        for block in sub:
            if not title_assigned and is_title_block(block):
                block["type"] = "title"
                title_assigned = True
            all_blocks.append(block)

    return all_blocks


# ── Output builders ───────────────────────────────────────────────────────────

def build_plain_text(blocks: list[dict]) -> str:
    """
    Build a clean UTF-8 plain text string from structured blocks.
    Tables are rendered as aligned TSV rows with a clear header separator.
    Photo captions are collapsed to a single annotated line.
    """
    parts: list[str] = []

    for block in blocks:
        btype = block["type"]

        if btype in ("title", "header"):
            text = block["text"].strip()
            border_char = "=" if btype == "title" else "-"
            border = border_char * min(len(text.split("\n")[0]), 72)
            parts.append(f"{border}\n{text}\n{border}")

        elif btype == "prose":
            parts.append(block["text"].strip())

        elif btype == "table":
            # Check if this is a direct continuation of the previous table
            # (QXP3 linked text-chain split: same column count, no intervening prose)
            prev = blocks[blocks.index(block) - 1] if block in blocks[1:] else None
            is_continuation = (
                prev is not None
                and prev["type"] == "table"
                and max(len(r) for r in prev["rows"]) == max(len(r) for r in block["rows"])
            )
            tsv_lines = ["\t".join(row) for row in block["rows"]]
            header = "[TABLE CONTINUED]" if is_continuation else "[TABLE]"
            parts.append(f"{header}\n" + "\n".join(tsv_lines) + "\n[/TABLE]")

        elif btype == "photo_caption":
            persons = block.get("persons", [])
            if persons:
                parts.append("[PHOTO CAPTION — persons: " + "; ".join(persons) + "]")
            else:
                parts.append("[PHOTO CAPTION]\n" + block["text"])

    return "\n\n".join(parts)


def build_prose_only(blocks: list[dict]) -> str:
    """Return only prose + header + title text, concatenated with newlines."""
    prose_types = {"prose", "title", "header"}
    return "\n\n".join(
        block["text"].strip()
        for block in blocks
        if block["type"] in prose_types and block["text"].strip()
    )


# ── Per-file pipeline ─────────────────────────────────────────────────────────

def process_file(src: Path, out_dir: Path, quiet: bool = False) -> dict:
    data = src.read_bytes()

    raw_blocks = extract_raw_blocks(data)
    structured  = assemble_document(raw_blocks)

    plain_text     = build_plain_text(structured)
    full_prose_text = build_prose_only(structured)

    result = {
        "source":          str(src),
        "file_size":       len(data),
        "block_count":     len(structured),
        "blocks":          structured,
        "full_prose_text": full_prose_text,
    }

    out_dir.mkdir(parents=True, exist_ok=True)
    txt_path  = out_dir / (src.stem + ".txt")
    json_path = out_dir / (src.stem + ".json")

    txt_path.write_text(plain_text, encoding="utf-8")
    json_path.write_text(
        json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    if not quiet:
        prose_chars = len(full_prose_text)
        persons = [
            p
            for b in structured
            if b["type"] == "photo_caption"
            for p in b.get("persons", [])
        ]
        table_count = sum(1 for b in structured if b["type"] == "table")
        print(
            f"  ✓  {src.name:<40}  "
            f"{len(structured):>3} blocks  "
            f"{prose_chars:>6} prose chars  "
            f"{table_count} table(s)  "
            f"{len(persons)} photo caption(s)"
        )

    return result


# ── CLI ───────────────────────────────────────────────────────────────────────

def build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Extract text from QuarkXPress 3 documents.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    p.add_argument(
        "inputs", nargs="+", metavar="FILE_OR_DIR",
        help="One or more QXP3 files or directories to process.",
    )
    p.add_argument(
        "--out", metavar="DIR", default=None,
        help="Output directory (default: same directory as each source file).",
    )
    p.add_argument(
        "--quiet", action="store_true",
        help="Suppress per-file progress output.",
    )
    p.add_argument(
        "--force", action="store_true",
        help="Process files even if they don't pass the MMXPR3 magic check.",
    )
    return p


def main(argv: list[str] | None = None) -> None:
    parser = build_arg_parser()
    args = parser.parse_args(argv)

    global_out = Path(args.out) if args.out else None
    processed = skipped = errors = 0

    for input_str in args.inputs:
        target = Path(input_str)

        if target.is_dir():
            candidates = sorted(p for p in target.rglob("*") if p.is_file())
        elif target.is_file():
            candidates = [target]
        else:
            print(f"ERROR: '{target}' not found.", file=sys.stderr)
            errors += 1
            continue

        for src in candidates:
            if not args.force and not is_qxp3(src):
                if target.is_file():
                    # User explicitly named this file — warn but try anyway
                    print(
                        f"WARNING: '{src.name}' lacks MMXPR3 magic; "
                        "use --force to process anyway.",
                        file=sys.stderr,
                    )
                skipped += 1
                continue

            out_dir = global_out if global_out else src.parent
            try:
                process_file(src, out_dir=out_dir, quiet=args.quiet)
                processed += 1
            except Exception as exc:
                print(f"ERROR processing '{src}': {exc}", file=sys.stderr)
                errors += 1

    if not args.quiet:
        print(
            f"\nDone. {processed} processed, {skipped} skipped "
            f"(non-QXP3), {errors} error(s)."
        )


if __name__ == "__main__":
    main()
