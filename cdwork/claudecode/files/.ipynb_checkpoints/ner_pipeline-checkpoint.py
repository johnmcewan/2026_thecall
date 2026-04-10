#!/usr/bin/env python3
"""
ner_pipeline.py — Example NER + entity-linking pipeline for qxp3_extract output.

Reads the .json sidecar produced by qxp3_extract.py and runs spaCy NER
over the prose text and table cells.  Produces a consolidated entity list
with document source, block index, entity text, label, and start/end offsets.

Requirements:
    pip install spacy
    python -m spacy download en_core_web_sm   # or en_core_web_trf for accuracy

Usage:
    python3 ner_pipeline.py path/to/file.json [path/to/file2.json ...]
    python3 ner_pipeline.py path/to/output_dir/   # process all .json files
"""

import json
import sys
from pathlib import Path

try:
    import spacy
    nlp = spacy.load("en_core_web_sm")
except OSError:
    print("Run: python -m spacy download en_core_web_sm", file=sys.stderr)
    sys.exit(1)


def text_segments(doc: dict) -> list[tuple[str, str, int]]:
    """
    Yield (segment_text, block_type, block_index) for all text-bearing blocks.
    Includes prose, titles, headers, and table cells (flattened).
    Photo caption person names are pre-extracted and yielded directly.
    """
    segments = []
    for i, block in enumerate(doc["blocks"]):
        btype = block["type"]

        if btype in ("prose", "title", "header"):
            segments.append((block["text"], btype, i))

        elif btype == "table":
            # Flatten all cells into one text blob, tab/newline separated
            flat = "
".join("	".join(row) for row in block["rows"])
            segments.append((flat, btype, i))

        elif btype == "photo_caption":
            # Person names already extracted; yield as PERSON entities directly
            for name in block.get("persons", []):
                segments.append((name, "photo_caption_person", i))

    return segments


def run_ner(json_path: Path) -> list[dict]:
    doc_data = json.loads(json_path.read_text(encoding="utf-8"))
    source = doc_data["source"]
    entities = []

    for text, block_type, block_idx in text_segments(doc_data):
        if not text.strip():
            continue

        # Photo caption persons are already entities — no NLP needed
        if block_type == "photo_caption_person":
            entities.append({
                "source": source,
                "block_idx": block_idx,
                "block_type": "photo_caption",
                "text": text,
                "label": "PERSON",
                "start": 0,
                "end": len(text),
                "confidence": "high",  # extracted by structural rule, not model
            })
            continue

        spacy_doc = nlp(text)
        for ent in spacy_doc.ents:
            entities.append({
                "source": source,
                "block_idx": block_idx,
                "block_type": block_type,
                "text": ent.text,
                "label": ent.label_,
                "start": ent.start_char,
                "end": ent.end_char,
                "confidence": "model",
            })

    return entities


def main(argv: list[str]) -> None:
    if len(argv) < 2:
        print(__doc__)
        sys.exit(0)

    inputs = [Path(a) for a in argv[1:]]
    all_entities: list[dict] = []

    for inp in inputs:
        if inp.is_dir():
            paths = sorted(inp.glob("*.json"))
        elif inp.suffix == ".json":
            paths = [inp]
        else:
            print(f"Skipping {inp} (not .json)", file=sys.stderr)
            continue

        for p in paths:
            print(f"Processing {p.name} ...", file=sys.stderr)
            ents = run_ner(p)
            all_entities.extend(ents)
            print(f"  {len(ents)} entities found", file=sys.stderr)

    # Write consolidated output as JSON Lines
    out_path = Path("entities.jsonl")
    with out_path.open("w", encoding="utf-8") as f:
        for ent in all_entities:
            f.write(json.dumps(ent, ensure_ascii=False) + "\n")

    print(f"\nWrote {len(all_entities)} entities to {out_path}")

    # Print a summary table
    from collections import Counter
    label_counts = Counter(e["label"] for e in all_entities)
    print("\nEntity label breakdown:")
    for label, count in label_counts.most_common():
        print(f"  {label:12} {count}")


if __name__ == "__main__":
    main(sys.argv)
