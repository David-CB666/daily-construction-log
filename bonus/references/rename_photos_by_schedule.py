"""
rename_photos_by_schedule.py — Batch rename site photos by construction schedule.
=============================================================================
Renames site photos from scattered filenames to structured format:
  {date}_{phase_code}_{phase_desc}_{seq}.{ext}

Example: IMG_20250721_001.jpg → 07-21_前期_施工前現場_01.jpg

CONFIG: edit BASE (photo folder), LOG (output plan path), MAPPING (date → phases).
The MAPPING dict maps each date to a list of (phase_code, phase_desc) tuples,
one per photo in that date folder (sorted alphabetically by filename).

Usage:
  1. Edit BASE to point to your photo directory (with date subfolders like 07-21/, 07-22/)
  2. Edit MAPPING to match your schedule phases per date
  3. Run: python rename_photos_by_schedule.py
  4. Review the plan JSON output before executing rename
"""
import os, json
from pathlib import Path

BASE = Path(r"D:\Projects\site_photos")                    # <-- EDIT: photo root (date subfolders)
LOG = Path(r"D:\Projects\output\photo_rename_plan.json")    # <-- EDIT: plan output path

# Mapping: date -> list of (phase_code, phase_desc) in alphabetical filename order
# Each entry corresponds to one photo in that date folder
MAPPING = {
    "07-21": [
        ("前期", "施工前現場"),
        ("前期", "施工前現場"),
    ],
    "07-22": [
        ("A1", "地盤準備圍蔽"),
        ("A1", "地盤準備圍蔽"),
        ("A1", "地盤準備圍蔽"),
        ("A1", "地盤準備圍蔽"),
    ],
    "07-25": [
        ("B1", "一樓拆卸"),
        ("B1", "一樓拆卸"),
        ("B1", "一樓拆卸"),
    ],
    "07-27": [
        ("A1", "地盤保護"),
        ("B1", "一樓拆卸"),
        ("B2", "二樓拆卸"),
        ("材料", "材料進場"),
    ],
}

# Build plan with per-date per-phase sequential numbering
plan = []
errors = []
for date, rules in MAPPING.items():
    folder = BASE / date
    if not folder.is_dir():
        errors.append(f"Missing folder: {folder}")
        continue
    files = sorted([f for f in folder.iterdir() if f.is_file() and f.suffix.lower() in (".jpg", ".jpeg", ".heic", ".png")])
    if len(files) != len(rules):
        errors.append(f"{date}: {len(files)} files but {len(rules)} mapping entries")
        continue
    counters = {}
    for idx, src in enumerate(files):
        code, desc = rules[idx]
        key = (date, code, desc)
        counters[key] = counters.get(key, 0) + 1
        seq = f"{counters[key]:02d}"
        ext = src.suffix.lower()
        new_name = f"{date}_{code}_{desc}_{seq}{ext}"
        plan.append({
            "date": date,
            "src": str(src),
            "new_name": new_name,
            "code": code,
            "desc": desc,
        })

# Print plan
print(f"Rename plan: {len(plan)} photos")
if errors:
    print("\nErrors:")
    for e in errors:
        print("  !", e)
print("\nPreview (first 30):")
for p in plan[:30]:
    print(f"  {Path(p['src']).name} -> {p['new_name']}")
if len(plan) > 30:
    print(f"  ... {len(plan)-30} more")

# Save plan
LOG.parent.mkdir(parents=True, exist_ok=True)
LOG.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")
print(f"\nPlan saved: {LOG}")

# Execute rename
if not errors:
    for p in plan:
        src = Path(p["src"])
        dst = src.parent / p["new_name"]
        if dst.exists():
            errors.append(f"Target exists: {dst}")
            continue
        src.rename(dst)
    print(f"\nRenamed: {len(plan)} photos")
    if errors:
        print("Partial errors:")
        for e in errors:
            print("  !", e)
else:
    print("\nRename skipped due to mapping errors.")
