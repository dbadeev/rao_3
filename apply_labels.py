#!/usr/bin/env python3
"""apply_labels.py - write field names from a labels JSON into a COPY of the PPTX (also for shapes inside groups).
Usage: python apply_labels.py template.pptx labels.json marked_copy.pptx [--selected-only]
Only renames shapes (shape.name); never edits text, geometry or layouts. First shape of a role gets the role name, further shapes get ROLE_2, ROLE_3, ...
--selected-only : rename shapes only on slides with selection == "included".
"""
import json, sys
from pptx import Presentation
from shapes_util import index_shapes

def main(src, labels, dst, selected_only=False):
    prs = Presentation(src); data = json.load(open(labels, encoding="utf-8")); renamed, skipped = 0, []
    for case in data["cases"]:
        if selected_only and case.get("selection", "included") != "included": continue
        by_id = index_shapes(prs.slides[case["json_index"]])
        for role, ids in case.get("fields", {}).items():
            for n, sid in enumerate(ids or []):
                if sid not in by_id: skipped.append((case["json_index"], role, sid)); continue
                by_id[sid][0].name = role if n == 0 else f"{role}_{n+1}"; renamed += 1
    prs.save(dst); print(f"Renamed shapes: {renamed}; skipped: {skipped}; saved: {dst}")
if __name__ == "__main__":
    a = [x for x in sys.argv[1:] if not x.startswith("--")]; main(*a[:3], selected_only="--selected-only" in sys.argv)
