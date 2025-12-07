#!/usr/bin/env python3
"""
Command-line version of the Streamlit exam seating generator.
Performs:
    1. Reads and converts workbook (in_timetable, in_course_roll_mapping, etc.)
    2. Runs process_master_excel()
    3. Generates PDFs
    4. Packs final ZIP into output directory
"""

import argparse
import tempfile
import os
import io
import zipfile
import shutil
from pathlib import Path
import pandas as pd
from datetime import datetime
import logging
import re
import sys

from modules.seating import process_master_excel
from modules.pdf_gen import generate_pdfs_for_output_dir
from modules.utils import safe_extract_photos_zip, create_final_zip, setup_output_dir

# ----------------------------------------------------------
# Logging: console + file (room_allocator.log next to this script)
# ----------------------------------------------------------
SCRIPT_DIR = Path(__file__).resolve().parent
LOG_FILE = SCRIPT_DIR / "room_allocator.log"
logger = logging.getLogger("room_allocator")
logger.setLevel(logging.INFO)

# remove existing handlers if any (helpful when reloading)
if logger.hasHandlers():
    logger.handlers.clear()

# console handler
ch = logging.StreamHandler(sys.stdout)
ch.setLevel(logging.INFO)
ch_fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
ch.setFormatter(ch_fmt)
logger.addHandler(ch)

# file handler (overwrite each run)
fh = logging.FileHandler(str(LOG_FILE), mode="w", encoding="utf-8")
fh.setLevel(logging.INFO)
fh.setFormatter(ch_fmt)
logger.addHandler(fh)

logger.info("Logger initialized. Writing logs to %s", LOG_FILE)

# ----------------------------------------------------------
# Regex for timetable splitting
# ----------------------------------------------------------
SUB_SPLIT_RE = re.compile(r"[;,|\n]")

# ----------------------------------------------------------
# Conversion helpers (EXACTLY same logic as Streamlit version)
# (I kept these functions in the script so CLI is self-contained)
# ----------------------------------------------------------

def find_best_col(df, candidates):
    cols_low = {c.lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in cols_low:
            return cols_low[cand.lower()]
    for cand in candidates:
        for c in df.columns:
            if cand.lower() in str(c).lower():
                return c
    return None

def expand_timetable(df_timetable, log):
    rows = []
    if df_timetable is None:
        log.append("in_timetable sheet not found.")
        return pd.DataFrame(columns=["Date", "Session", "Subject"])

    df = df_timetable.copy()

    date_col = find_best_col(df, ["Date", "date"])
    morning_col = find_best_col(df, ["Morning"])
    evening_col = find_best_col(df, ["Evening"])

    if date_col is None:
        log.append("No Date column in in_timetable.")
        return pd.DataFrame(columns=["Date", "Session", "Subject"])

    def split(cell):
        if pd.isna(cell):
            return []
        s = str(cell).strip()
        if s == "":
            return []
        return [x.strip() for x in SUB_SPLIT_RE.split(s) if x.strip()]

    for _, row in df.iterrows():
        raw_date = row.get(date_col)
        if pd.isna(raw_date):
            continue

        try:
            date_val = pd.to_datetime(raw_date).date()
        except Exception:
            date_val = str(raw_date).strip()

        if morning_col:
            for sub in split(row.get(morning_col)):
                rows.append({"Date": date_val, "Session": "Morning", "Subject": sub})

        if evening_col:
            for sub in split(row.get(evening_col)):
                rows.append({"Date": date_val, "Session": "Evening", "Subject": sub})

    return pd.DataFrame(rows, columns=["Date", "Session", "Subject"])

def convert_enrollment(df_course_roll, log):
    if df_course_roll is None:
        log.append("in_course_roll_mapping sheet not found.")
        return pd.DataFrame(columns=["Roll", "Subject"])

    df = df_course_roll.copy()
    roll_col = find_best_col(df, ["rollno", "roll"])
    subj_col = find_best_col(df, ["course_code", "subject", "course"])

    if roll_col is None or subj_col is None:
        if len(df.columns) >= 2:
            roll_col, subj_col = df.columns[:2]
        else:
            return pd.DataFrame(columns=["Roll", "Subject"])

    out = []
    for _, r in df.iterrows():
        roll = r.get(roll_col)
        sub = r.get(subj_col)
        if pd.isna(roll) or pd.isna(sub):
            continue
        out.append({"Roll": str(roll).strip(), "Subject": str(sub).strip()})

    return pd.DataFrame(out, columns=["Roll", "Subject"])

def convert_student_map(df_roll_name, log):
    if df_roll_name is None:
        log.append("in_roll_name_mapping missing.")
        return pd.DataFrame(columns=["Roll", "Name"])

    df = df_roll_name.copy()
    roll_col = find_best_col(df, ["roll", "rollno"])
    name_col = find_best_col(df, ["name", "studentname"])

    if roll_col is None or name_col is None:
        if len(df.columns) >= 2:
            roll_col, name_col = df.columns[:2]
        else:
            return pd.DataFrame(columns=["Roll", "Name"])

    rows = []
    for _, r in df.iterrows():
        roll = r.get(roll_col)
        name = r.get(name_col)
        if pd.isna(roll) or pd.isna(name):
            continue
        rows.append({"Roll": str(roll).strip(), "Name": str(name).strip()})
    return pd.DataFrame(rows, columns=["Roll", "Name"])

def convert_rooms(df_room_capacity, log):
    if df_room_capacity is None:
        log.append("in_room_capacity sheet not found.")
        return pd.DataFrame(columns=["Room", "Capacity", "Block"])

    df = df_room_capacity.copy()
    room_col = find_best_col(df, ["Room No.", "Room"])
    cap_col = find_best_col(df, ["Exam Capacity", "Capacity"])
    block_col = find_best_col(df, ["Block", "Building"])

    if room_col is None or cap_col is None:
        if len(df.columns) >= 2:
            room_col, cap_col = df.columns[:2]
        else:
            return pd.DataFrame(columns=["Room", "Capacity", "Block"])

    rows = []
    for _, r in df.iterrows():
        room = r.get(room_col)
        cap = r.get(cap_col)
        block = r.get(block_col) if block_col else ""

        if pd.isna(room):
            continue

        try:
            cap_i = int(float(cap)) if not pd.isna(cap) else 0
        except Exception:
            cap_i = 0

        rows.append({"Room": str(room).strip(), "Capacity": cap_i, "Block": str(block).strip()})

    return pd.DataFrame(rows, columns=["Room", "Capacity", "Block"])

def convert_workbook_bytes(input_bytes):
    """
    EXACT same logic as Streamlit version.
    """
    log = []
    bio = io.BytesIO(input_bytes)

    try:
        xls = pd.ExcelFile(bio, engine="openpyxl")
    except Exception as e:
        log.append(f"Failed to load XLS: {e}")
        return None, log

    sheets = {}
    for name in xls.sheet_names:
        try:
            sheets[name] = pd.read_excel(xls, sheet_name=name, engine="openpyxl")
        except Exception:
            sheets[name] = None

    required = {"Schedule", "Enrollment", "Rooms", "Student_Mapping"}
    # check presence case-insensitively
    present = set([s for s in sheets.keys() if sheets[s] is not None])
    if required.issubset(present):
        log.append("Workbook already contains required sheets.")
        return input_bytes, log

    # pick sheets
    def pick(pref):
        for p in pref:
            for s in sheets:
                if s and s.lower() == p.lower():
                    return sheets[s]
        for p in pref:
            for s in sheets:
                if s and p.lower() in s.lower():
                    return sheets[s]
        return None

    df_timetable = pick(["in_timetable", "timetable"])
    df_course_roll = pick(["in_course_roll_mapping", "course_roll_mapping", "in_course_roll"])
    df_roll_name = pick(["in_roll_name_mapping", "roll_name_mapping", "in_roll_name_mapping"])
    df_room_capacity = pick(["in_room_capacity", "room_capacity", "in_room_capacity"])

    schedule_df = expand_timetable(df_timetable, log)
    enroll_df = convert_enrollment(df_course_roll, log)
    stud_map_df = convert_student_map(df_roll_name, log)
    rooms_df = convert_rooms(df_room_capacity, log)

    out_bio = io.BytesIO()
    with pd.ExcelWriter(out_bio, engine="openpyxl") as writer:
        schedule_df.to_excel(writer, sheet_name="Schedule", index=False)
        enroll_df.to_excel(writer, sheet_name="Enrollment", index=False)
        rooms_df.to_excel(writer, sheet_name="Rooms", index=False)
        stud_map_df.to_excel(writer, sheet_name="Student_Mapping", index=False)

    out_bio.seek(0)
    log.append("Conversion completed.")
    return out_bio.read(), log

# ----------------------------------------------------------
# CLI main
# ----------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Command-line exam seating generator (equivalent to Streamlit app).")
    parser.add_argument("--master", required=True, help="Master Excel file (.xlsx)")
    parser.add_argument("--photos", default=None, help="Photos ZIP (optional)")
    parser.add_argument("--outdir", required=True, help="Directory to store final ZIP")
    parser.add_argument("--buffer", type=int, default=0, help="Buffer seats per room")
    parser.add_argument("--filling_mode", choices=["Dense", "Sparse"], default="Dense")
    parser.add_argument("--name", default="Output_Zip", help="Output ZIP base name")
    parser.add_argument("--keep_tmp", action="store_true", help="Keep temporary working directory for debugging")
    args = parser.parse_args()

    tmpdir = Path(tempfile.mkdtemp(prefix="room_alloc_"))
    logger.info("Using temporary directory: %s", tmpdir)
    photos_dir = tmpdir / "photos"
    photos_dir.mkdir(parents=True, exist_ok=True)

    # Extract photos if any
    if args.photos:
        logger.info("Extracting photos from %s -> %s", args.photos, photos_dir)
        try:
            # try the project's helper first
            safe_extract_photos_zip(args.photos, str(photos_dir))
        except Exception as e:
            logger.warning("safe_extract_photos_zip failed (%s); falling back to zipfile extraction.", e)
            try:
                with zipfile.ZipFile(args.photos, "r") as zf:
                    for member in zf.namelist():
                        if member.endswith("/") or member.startswith("__MACOSX") or ".." in Path(member).parts:
                            continue
                        outp = photos_dir / Path(member).name
                        with zf.open(member) as src, open(outp, "wb") as dst:
                            dst.write(src.read())
                logger.info("Photos extracted (fallback).")
            except Exception as e2:
                logger.exception("Failed extracting photos: %s", e2)

    # Setup output root (temporary folder structure)
    output_root = setup_output_dir(str(tmpdir))
    logger.info("Output root (temporary): %s", output_root)

    # Read master file
    try:
        with open(args.master, "rb") as f:
            master_bytes = f.read()
    except Exception as e:
        logger.exception("Failed to read master file: %s", e)
        return 2

    # Convert workbook (if needed)
    converted_bytes, logs = convert_workbook_bytes(master_bytes)
    for line in logs:
        logger.info(line)

    if converted_bytes is None:
        logger.error("Workbook conversion failed.")
        return 3

    # Run allocation
    try:
        logger.info("Running seating allocation...")
        process_master_excel(io.BytesIO(converted_bytes), output_root,
                             buffer=args.buffer, filling_mode=args.filling_mode)
        logger.info("Seating allocation completed.")
    except Exception as e:
        logger.exception("Seating allocation failed: %s", e)
        if not args.keep_tmp:
            try:
                shutil.rmtree(tmpdir)
            except Exception:
                pass
        return 4

    # Generate PDFs
    try:
        logger.info("Generating PDFs...")
        generate_pdfs_for_output_dir(output_root, str(photos_dir), cols=3, paper_format="A4")
        logger.info("PDF generation finished.")
    except Exception as e:
        logger.exception("PDF generation failed: %s", e)

    # Create final ZIP
    try:
        logger.info("Creating final ZIP...")
        final_zip = create_final_zip(output_root, args.name)
        dest = Path(args.outdir)
        dest.mkdir(parents=True, exist_ok=True)
        final_path = dest / Path(final_zip).name
        shutil.move(final_zip, final_path)
        logger.info("Success! Final ZIP created at: %s", final_path)
        print(final_path)
    except Exception as e:
        logger.exception("Failed to create/move final ZIP: %s", e)
        if not args.keep_tmp:
            try:
                shutil.rmtree(tmpdir)
            except Exception:
                pass
        return 5

    # Cleanup unless requested to keep
    if not args.keep_tmp:
        try:
            shutil.rmtree(tmpdir)
            logger.info("Removed temporary directory.")
        except Exception:
            logger.warning("Failed to remove temporary directory: %s", tmpdir)

    return 0

if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)
