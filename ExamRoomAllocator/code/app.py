# app.py (modified with logging to app.log)
import streamlit as st
import tempfile
import zipfile
import os
import io
import re
from pathlib import Path
import pandas as pd
from datetime import datetime
import logging

# -------------------- Logging configuration --------------------
# write logs to console AND to app.log next to this file
BASE_DIR = Path(__file__).resolve().parent
LOG_FILE = BASE_DIR / "app.log"

# Remove existing handlers (useful when reloading in dev)
for h in logging.root.handlers[:]:
    logging.root.removeHandler(h)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[
        logging.StreamHandler(),                       # console
        logging.FileHandler(str(LOG_FILE), mode="w"),  # file (overwrites each run)
    ],
)

LOG = logging.getLogger(__name__)
LOG.info("Logging initialized. Writing logs to %s", LOG_FILE)
# --------------------------------------------------------------

# local module imports (assumes modules package exists)
from modules.seating import process_master_excel
from modules.pdf_gen import generate_pdfs_for_output_dir
from modules.utils import safe_extract_photos_zip, create_final_zip, setup_output_dir

# Pillow used to create placeholder image at runtime if missing
try:
    from PIL import Image, ImageDraw, ImageFont
    PIL_AVAILABLE = True
except Exception:
    PIL_AVAILABLE = False

# --- Streamlit UI setup ---
st.set_page_config(page_title="Exam Seating & Attendance Generator", layout="wide")
st.title("Automated Exam Seating & Attendance System (with auto-convert)")
st.markdown(
    "Upload the master Excel file (with required sheets). Optionally upload a ZIP of photos "
    "(each file named exactly as the roll number, e.g. 1401ZZ16.jpg)."
)

# File inputs
master_file = st.file_uploader("Master Excel (.xlsx)", type=["xlsx", "xls"])
photos_zip = st.file_uploader("ZIP of photos (each file named ROLL.jpg)", type=["zip"])
PHOTOS_DIR = BASE_DIR / "photos"


col1, col2 = st.columns(2)
with col1:
    buffer = st.number_input("Buffer seats per room (reduce usable capacity)", min_value=0, value=0, step=1)
    filling_mode = st.selectbox("Filling Mode", ["Dense", "Sparse"])  # Dense tries to fill rooms tightly
with col2:
    output_name = st.text_input("Output ZIP name (without .zip)", value="Output_Zip")

# ---------- Converter helpers (your original code) ----------
SUB_SPLIT_RE = re.compile(r"[;,|\n]")  # split separators for timetable cells

def find_best_col(df, candidates):
    cols_low = {c.lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in cols_low:
            return cols_low[cand.lower()]
    # fallback: try substring match
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
    date_col = find_best_col(df, ["Date", "date", "DATE"])
    morning_col = find_best_col(df, ["Morning", "morning", "MORNING"])
    evening_col = find_best_col(df, ["Evening", "evening", "EVENING"])

    if date_col is None:
        log.append("Could not find a Date column in in_timetable; Schedule will be empty.")
        return pd.DataFrame(columns=["Date", "Session", "Subject"])

    def split_subjects(cell):
        if pd.isna(cell):
            return []
        s = str(cell).strip()
        if s == "":
            return []
        parts = [p.strip() for p in SUB_SPLIT_RE.split(s) if p and p.strip()]
        return parts

    for _, row in df.iterrows():
        raw_date = row.get(date_col)
        if pd.isna(raw_date):
            continue
        # try to normalize date
        date_val = None
        try:
            date_val = pd.to_datetime(raw_date).date()
        except Exception:
            try:
                date_val = datetime.strptime(str(raw_date).strip(), "%d/%m/%y").date()
            except Exception:
                try:
                    date_val = datetime.strptime(str(raw_date).strip(), "%d/%m/%Y").date()
                except Exception:
                    date_val = str(raw_date).strip()

        if morning_col:
            for sub in split_subjects(row.get(morning_col)):
                rows.append({"Date": date_val, "Session": "Morning", "Subject": sub})
        if evening_col:
            for sub in split_subjects(row.get(evening_col)):
                rows.append({"Date": date_val, "Session": "Evening", "Subject": sub})

    df_schedule = pd.DataFrame(rows, columns=["Date", "Session", "Subject"])
    try:
        df_schedule["Date"] = pd.to_datetime(df_schedule["Date"]).dt.date
    except Exception:
        pass
    if df_schedule.empty:
        log.append("Expanded Schedule is empty after parsing in_timetable.")
    return df_schedule

def convert_enrollment(df_course_roll, log):
    if df_course_roll is None:
        log.append("in_course_roll_mapping sheet not found.")
        return pd.DataFrame(columns=["Roll", "Subject"])
    df = df_course_roll.copy()
    roll_col = find_best_col(df, ["rollno", "roll", "Roll", "ROLL"])
    subj_col = find_best_col(df, ["course_code", "course", "coursecode", "Subject", "subject"])
    if roll_col is None or subj_col is None:
        log.append(f"in_course_roll_mapping missing expected columns; found: {list(df.columns)}")
        if len(df.columns) >= 2:
            roll_col = df.columns[0]
            subj_col = df.columns[1]
            log.append(f"Falling back to first two columns: {roll_col}, {subj_col}")
        else:
            return pd.DataFrame(columns=["Roll", "Subject"])
    out_rows = []
    for _, r in df.iterrows():
        roll = r.get(roll_col)
        subj = r.get(subj_col)
        if pd.isna(roll) or pd.isna(subj):
            continue
        out_rows.append({"Roll": str(roll).strip(), "Subject": str(subj).strip()})
    df_enroll = pd.DataFrame(out_rows, columns=["Roll", "Subject"])
    if df_enroll.empty:
        log.append("Enrollment table empty after conversion from in_course_roll_mapping.")
    return df_enroll

def convert_student_map(df_roll_name, log):
    if df_roll_name is None:
        log.append("in_roll_name_mapping sheet not found.")
        return pd.DataFrame(columns=["Roll", "Name"])
    df = df_roll_name.copy()
    roll_col = find_best_col(df, ["roll", "rollno", "Roll", "ROLL"])
    name_col = find_best_col(df, ["name", "Name", "StudentName", "student_name"])
    if roll_col is None or name_col is None:
        log.append(f"in_roll_name_mapping missing expected columns; found: {list(df.columns)}")
        if len(df.columns) >= 2:
            roll_col = df.columns[0]
            name_col = df.columns[1]
            log.append(f"Falling back to first two columns: {roll_col}, {name_col}")
        else:
            return pd.DataFrame(columns=["Roll", "Name"])
    rows = []
    for _, r in df.iterrows():
        roll = r.get(roll_col)
        name = r.get(name_col)
        if pd.isna(roll) or pd.isna(name):
            continue
        rows.append({"Roll": str(roll).strip(), "Name": str(name).strip()})
    df_map = pd.DataFrame(rows, columns=["Roll", "Name"])
    if df_map.empty:
        log.append("Student_Mapping is empty after conversion.")
    return df_map

def convert_rooms(df_room_capacity, log):
    if df_room_capacity is None:
        log.append("in_room_capacity sheet not found.")
        return pd.DataFrame(columns=["Room", "Capacity", "Block"])
    df = df_room_capacity.copy()
    room_col = find_best_col(df, ["Room No.", "Room No", "Room", "roomno", "room no"])
    cap_col = find_best_col(df, ["Exam Capacity", "ExamCapacity", "Capacity", "capacity", "Exam Cap"])
    block_col = find_best_col(df, ["Block", "block", "Building", "block name"])
    if room_col is None or cap_col is None:
        log.append(f"in_room_capacity missing expected columns; found: {list(df.columns)}")
        if len(df.columns) >= 2:
            room_col = df.columns[0]
            cap_col = df.columns[1]
            log.append(f"Falling back to first two columns: {room_col}, {cap_col}")
        else:
            return pd.DataFrame(columns=["Room", "Capacity", "Block"])
    rows = []
    for _, r in df.iterrows():
        room = r.get(room_col)
        cap = r.get(cap_col)
        block = r.get(block_col) if block_col is not None else ""
        if pd.isna(room):
            continue
        room_s = str(room).strip()
        try:
            cap_i = int(float(cap)) if not pd.isna(cap) else 0
        except Exception:
            try:
                cap_i = int(re.findall(r"\d+", str(cap))[0])
            except Exception:
                cap_i = 0
        block_s = str(block).strip() if not pd.isna(block) else ""
        rows.append({"Room": room_s, "Capacity": cap_i, "Block": block_s})
    df_rooms = pd.DataFrame(rows, columns=["Room", "Capacity", "Block"])
    if df_rooms.empty:
        log.append("Rooms table empty after conversion.")
    return df_rooms

def convert_workbook_bytes(input_bytes):
    """
    Accepts bytes of an uploaded workbook and returns bytes for the converted
    master workbook (with sheets Schedule, Enrollment, Rooms, Student_Mapping).
    Also returns a log (list of strings).
    """
    log = []
    bio = io.BytesIO(input_bytes)
    try:
        xls = pd.ExcelFile(bio, engine="openpyxl")
    except Exception as e:
        log.append(f"Failed to read uploaded workbook: {e}")
        return None, log

    # load existing sheets into dict
    sheets = {}
    for name in xls.sheet_names:
        try:
            sheets[name] = pd.read_excel(xls, sheet_name=name, engine="openpyxl")
        except Exception:
            sheets[name] = None

    # check if required sheets already present
    required = {"Schedule", "Enrollment", "Rooms", "Student_Mapping"}
    if required.issubset(set([s for s in sheets.keys() if s is not None])):
        # no conversion needed; just return original bytes
        log.append("Uploaded workbook already has required sheets; no conversion done.")
        return input_bytes, log

    # pick likely candidate sheets (case-insensitive)
    def pick_sheet(pref_names):
        for p in pref_names:
            for s in sheets:
                if s is None:
                    continue
                if s.lower().strip() == p.lower().strip():
                    return sheets[s]
        for p in pref_names:
            for s in sheets:
                if s is None:
                    continue
                if p.lower().strip() in s.lower().strip():
                    return sheets[s]
        return None

    df_timetable = pick_sheet(["in_timetable", "timetable", "in timetable"])
    df_course_roll = pick_sheet(["in_course_roll_mapping", "course_roll_mapping", "in_course_roll"])
    df_roll_name = pick_sheet(["in_roll_name_mapping", "roll_name_mapping", "in_roll_name_mapping"])
    df_room_capacity = pick_sheet(["in_room_capacity", "room_capacity", "in_room_capacity"])

    # convert pieces
    schedule_df = expand_timetable(df_timetable, log)
    enrollment_df = convert_enrollment(df_course_roll, log)
    student_map_df = convert_student_map(df_roll_name, log)
    rooms_df = convert_rooms(df_room_capacity, log)

    # write to BytesIO
    out_bio = io.BytesIO()
    try:
        with pd.ExcelWriter(out_bio, engine="openpyxl") as writer:
            schedule_df.to_excel(writer, sheet_name="Schedule", index=False)
            enrollment_df.to_excel(writer, sheet_name="Enrollment", index=False)
            rooms_df.to_excel(writer, sheet_name="Rooms", index=False)
            student_map_df.to_excel(writer, sheet_name="Student_Mapping", index=False)
        out_bio.seek(0)
        log.append("Conversion succeeded; master workbook produced in-memory.")
        return out_bio.read(), log
    except Exception as e:
        log.append(f"Failed to write converted workbook: {e}")
        return None, log

# ---------- End converter helpers ----------

# Ensure assets placeholder exists at runtime (create simple placeholder if missing)

def ensure_placeholder_image():
    assets_dir = Path("assets")
    assets_dir.mkdir(exist_ok=True)
    placeholder = assets_dir / "nopic.jpg"

    if not placeholder.exists():
        if PIL_AVAILABLE:
            try:
                img = Image.new("RGB", (400, 400), color=(240, 240, 240))
                draw = ImageDraw.Draw(img)
                try:
                    fnt = ImageFont.load_default()
                except Exception:
                    fnt = None

                text = "No Image"

                # Pillow ≥10: use textbbox() instead of deprecated textsize()
                bbox = draw.textbbox((0, 0), text, font=fnt)
                w = bbox[2] - bbox[0]
                h = bbox[3] - bbox[1]

                draw.text(
                    ((400 - w) / 2, (400 - h) / 2),
                    text,
                    fill=(80, 80, 80),
                    font=fnt
                )

                img.save(placeholder, format="JPEG", quality=85)
                st.info("Created placeholder image at assets/nopic.jpg")

            except Exception as e:
                st.warning(f"Could not create placeholder image: {e}")
        else:
            st.warning("Pillow not available. Ensure assets/nopic.jpg exists manually.")

ensure_placeholder_image()

# Main button
if st.button("Generate"):
    if master_file is None:
        st.error("Please upload the master Excel file.")
    else:
        tmpdir = tempfile.mkdtemp()
        photos_dir = os.path.join(tmpdir, "photos")
        os.makedirs(photos_dir, exist_ok=True)

        # extract photos if provided
        if photos_zip is not None:
            try:
                # prefer your safe helper. It may accept Streamlit UploadedFile directly.
                try:
                    # safe_extract_photos_zip might accept file-like or path; try both
                    safe_extract_photos_zip(photos_zip, photos_dir)
                except TypeError:
                    # maybe safe_extract_photos_zip expects bytes/path - handle common cases
                    zbytes = photos_zip.read()
                    safe_extract_photos_zip(io.BytesIO(zbytes), photos_dir)
                st.info("Photos extracted to temporary folder.")
            except Exception as e:
                st.warning(f"safe_extract_photos_zip failed, attempting fallback extraction: {e}")
                # fallback: extract using zipfile safely (prevent absolute paths)
                try:
                    zbytes = photos_zip.read()
                    with zipfile.ZipFile(io.BytesIO(zbytes)) as zf:
                        for member in zf.namelist():
                            # skip directories and suspicious names
                            if member.endswith("/") or member.startswith("__MACOSX") or ".." in Path(member).parts:
                                continue
                            target_path = Path(photos_dir) / Path(member).name
                            with zf.open(member) as source, open(target_path, "wb") as dest:
                                dest.write(source.read())
                    st.info("Photos extracted (fallback) to temporary folder.")
                except Exception as ee:
                    st.error(f"Error extracting photos (fallback): {ee}")
                    st.stop()
        else:
            st.info("No photos ZIP provided — placeholder images will be used automatically.")

        # prepare output dir
        output_root = setup_output_dir(tmpdir)  # your helper returns path to output dir
        LOG.info("Using temporary output root: %s", output_root)

        # read uploaded master bytes
        try:
            master_bytes = master_file.read()
        except Exception as e:
            st.error(f"Failed to read uploaded master file: {e}")
            st.stop()

        # try to convert or pass-through
        converted_bytes, conv_log = convert_workbook_bytes(master_bytes)
        for line in conv_log:
            st.info(line)
            LOG.info(line)

        if converted_bytes is None:
            st.error("Conversion failed. See info messages above.")
            st.stop()

        # Pass the converted workbook (as BytesIO) to process_master_excel
        try:
            converted_bio = io.BytesIO(converted_bytes)
            # process_master_excel expects a file-like; we pass BytesIO
            with st.spinner("Running seating allocation (this may take a few seconds)..."):
                # preserve the buffer and mode inputs you already have
                process_master_excel(converted_bio, output_root, buffer=int(buffer), filling_mode=filling_mode)
            st.success("Seating allocation completed.")
            LOG.info("Seating allocation completed successfully.")
        except Exception as e:
            st.error(f"Seating allocation failed: {e}")
            LOG.exception("Seating allocation failed")
            st.stop()

        # Generate PDFs for all allocation files.
        # Try to pass cols=3 and paper_format='A4' if your function supports it,
        # otherwise call with the original signature.
        try:
            with st.spinner("Generating PDFs (A4, 3 columns)..."):
                try:
                    # preferred: explicit args
                    generate_pdfs_for_output_dir(output_root, str(PHOTOS_DIR), cols=3, paper_format="A4")
                except TypeError:
                    # fallback: maybe function signature doesn't accept kwargs
                    try:
                        generate_pdfs_for_output_dir(output_root, photos_dir, 3, "A4")
                    except TypeError:
                        # ultimate fallback: two-arg call (original behavior)
                        generate_pdfs_for_output_dir(output_root, photos_dir)
            st.success("PDF generation finished.")
            LOG.info("PDF generation finished.")
        except Exception as e:
            st.error(f"PDF generation failed: {e}")
            LOG.exception("PDF generation failed")
            st.stop()

        # Zip the output_root into a downloadable file
        try:
            zip_path = create_final_zip(output_root, output_name)
            with open(zip_path, "rb") as f:
                st.download_button("Download Output ZIP", data=f, file_name=os.path.basename(zip_path), mime="application/zip")
            st.success("Generation complete. Download the zip using the button above.")
            LOG.info("Created final ZIP: %s", zip_path)
        except Exception as e:
            st.error(f"Failed to create/download final ZIP: {e}")
            LOG.exception("Failed to create final ZIP")
            st.stop()
