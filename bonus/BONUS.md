# Bonus Tool — Construction Photo Schedule Naming

Rename scattered site photos into structured filenames based on the construction schedule.

## The Problem

Site photos come in with random filenames like `IMG_20250721_001.jpg`. You need them named by:
- Date
- Phase code (A1/B1/C1 from schedule)
- Phase description
- Sequential number

Example: `IMG_20250721_001.jpg` → `07-21_A1_地盤準備圍蔽_01.jpg`

## The Solution

A 5-step workflow:

```
Schedule (Excel) → Scan photos → Contact sheets (visual ID) → Build mapping → Batch rename
```

### Step 1: Read Construction Schedule

Parse `施工進度表.xlsx` to extract phase codes, descriptions, start/end dates. Build a `date → active phases` lookup.

### Step 2: Scan & Backup Photos

- Scan photo directory (date subfolders)
- Backup entire folder before renaming
- Count photos per date and format (.jpg / .jpeg / .heic / .png)

### Step 3: Generate Contact Sheets

`generate_contact_sheets.py` — creates visual contact sheet PNGs:

- .heic conversion via ffmpeg (Pillow can't read .heic natively)
- 3-column grid with index badges and original filenames
- One contact sheet per date

### Step 4: Build Rename Mapping

Manually review contact sheets and identify which photos belong to which phase (especially on days with overlapping phases).

### Step 5: Execute Batch Rename

`rename_photos_by_schedule.py` — executes the rename:

- Format: `{date}_{phase_code}_{phase_desc}_{seq}.{ext}`
- Sequential numbering per date+phase combination
- Outputs plan JSON for audit trail before executing

## Phase Code Reference

| Code | Phase |
|------|-------|
| 前期 | Pre-construction / site survey |
| A1 | Site prep, hoarding, safety measures |
| B1/B2/B3 | Floor 1/2/3 demolition |
| C1/C2/C3 | Floor 1/2/3 water supply & drainage |
| D1/D2/D3 | Floor 1/2/3 waterproofing & leveling |
| E1/E2/E3 | Floor 1/2/3 tiling |
| 材料 | Material delivery / staging |

## Usage

```bash
# Step 1: Generate contact sheets for visual identification
python bonus/references/generate_contact_sheets.py

# Step 2: Review contact sheets, then edit MAPPING in rename script
# Step 3: Execute rename
python bonus/references/rename_photos_by_schedule.py
```

### Configuration

Edit the following variables at the top of each script:

| Variable | Description |
|----------|-------------|
| `BASE` | Photo root directory (with date subfolders) |
| `OUT` | Contact sheet output directory |
| `LOG` | Rename plan JSON output path |
| `MAPPING` | Date → phase assignments (one tuple per photo) |
| `FFMPEG` | ffmpeg path (for .heic conversion) |

## Key Notes

- **Don't rename blindly by date**: Days with overlapping phases (e.g., demolition + plumbing) require visual identification
- **.heic handling**: Windows Pillow can't read .heic; use ffmpeg to convert
- **Backup first**: Always backup the entire photo folder before renaming
- **Chinese filenames safe**: Keep output within workspace, avoid system directories

## Dependencies

```bash
pip install Pillow
# ffmpeg required for .heic photos (install: scoop install ffmpeg)
```
