from pathlib import Path
import math
import re
from fpdf import FPDF
from PIL import Image
import pandas as pd
import logging
from typing import Optional
import math
import textwrap

from modules.utils import get_photo_for_roll

LOG = logging.getLogger(__name__)

A4_W = 210.0
A4_H = 297.0


class AttendancePDF(FPDF):
    def __init__(self, orientation='P', unit='mm', format='A4'):
        super().__init__(orientation=orientation, unit=unit, format=format)
        self.set_auto_page_break(False)


def _draw_page_header(pdf: AttendancePDF, date_str: str, session: str, room: str, subject: str, student_count: int):
    margin = 10.0
    pdf.set_xy(margin, 10)
    pdf.set_font("Helvetica", "B", 14)
    pdf.cell(0, 8, "IITP Attendance System", ln=1, align="C")
    pdf.ln(1)
    pdf.set_font("Helvetica", size=10)
    pdf.cell(0, 6, f"Date: {date_str} | Session: {session} | Room: {room} | Student count: {student_count}", ln=1, align="L")
    pdf.cell(0, 6, f"Subject: {subject}", ln=1, align="L")
    y = pdf.get_y() + 2
    pdf.set_line_width(0.5)
    pdf.line(margin, y, A4_W - margin, y)
    pdf.ln(4)


def _fit_image_size_mm(img_path: Path, max_w_mm: float, max_h_mm: float):
    try:
        with Image.open(img_path) as im:
            w, h = im.size
    except Exception:
        return max_w_mm, max_h_mm
    aspect = w / h if h != 0 else 1.0
    w_mm = max_w_mm
    h_mm = w_mm / aspect
    if h_mm > max_h_mm:
        h_mm = max_h_mm
        w_mm = h_mm * aspect
    return w_mm, h_mm

def _generate_pdf_for_allocation_file(allocation_path: Path, photos_dir: str, cols: int = 3, paper_format: str = "A4"):
   
    df = pd.read_excel(allocation_path, engine="openpyxl")
    cols_map = {c.lower().strip(): c for c in df.columns}

    rows_mode = ("roll" in cols_map) and ("name" in cols_map or "subject" in cols_map)

    rolls = []
    names = {}
    date_str = ""
    session = ""
    room = ""
    subject = ""

    # infer metadata from parent directories if possible
    for p in allocation_path.parents:
        pname = p.name
        if not date_str and re.match(r"\d{4}[_-]\d{1,2}[_-]\d{1,2}", pname):
            date_str = pname
        if not session and pname.lower() in {"morning", "evening"}:
            session = pname

    if rows_mode:
        name_col = cols_map.get("name")
        roll_col = cols_map.get("roll")
        subj_col = cols_map.get("subject")
        room_col = cols_map.get("room")
        for _, r in df.iterrows():
            roll = str(r.get(roll_col) or "").strip()
            if not roll:
                continue
            rolls.append(roll)
            if name_col:
                nm = str(r.get(name_col) or "").strip()
                if nm:
                    names[roll] = nm
            if not subject and subj_col:
                subject = subject or str(r.get(subj_col) or "").strip()
            if not room and room_col:
                room = room or str(r.get(room_col) or "").strip()
        # fallback metadata from filename
        stem_parts = allocation_path.stem.split("_")
        if not date_str and len(stem_parts) >= 1:
            date_str = stem_parts[0]
        if not session and len(stem_parts) >= 2:
            session = stem_parts[1]
    else:
        required = ["date", "shift", "room", "subject", "roll_list"]
        missing = [r for r in required if r not in cols_map]
        if missing:
            raise ValueError(f"Allocation file {allocation_path} missing required columns: {missing}")
        date_val = df.iloc[0][cols_map["date"]]
        try:
            date_str = pd.to_datetime(date_val).strftime("%Y-%m-%d")
        except Exception:
            date_str = str(date_val)
        session = str(df.iloc[0][cols_map["shift"]])
        room = str(df.iloc[0][cols_map["room"]])
        subject = str(df.iloc[0][cols_map["subject"]])
        raw_roll_list = str(df.iloc[0][cols_map["roll_list"]] or "")
        rolls = [r.strip() for r in raw_roll_list.split(";") if r.strip()]

    student_count = len(rolls)

    # ---------------------- Layout tunables ----------------------
    left_margin = 10.0
    right_margin = 10.0
    top_margin = 18.0
    bottom_margin = 24.0
    cols = max(1, int(cols))
    col_gap = 0.0

    img_box_size = 20.0      # square image area side in mm
    img_inner_pad = 2.0
    name_max_lines = 2
    name_line_h = 4.2
    roll_line_h = 5.5
    extra_cell_vpad = 2.0
    section_spacing = 6.0
    # deterministic cell height so things don't overlap
    name_area_h = name_max_lines * name_line_h
    text_block_h = name_area_h + roll_line_h + (img_inner_pad * 2) + 4.0  # small extra room
    image_block_h = img_inner_pad * 2 + img_box_size
    cell_height = max(text_block_h, image_block_h) + extra_cell_vpad

    page_w = A4_W
    page_h = A4_H
    usable_w = page_w - left_margin - right_margin - (cols - 1) * col_gap
    col_w = usable_w / cols
    pdf = AttendancePDF(format=paper_format)
    pdf.set_auto_page_break(False)

    # ---------- SECTION 1: FIRST PAGE HEADER (dedicated top area on first page) ----------
    pdf.add_page()
    pdf.set_xy(left_margin, 20)
    pdf.set_font("Helvetica", "B", 18)
    pdf.cell(0, 10, "IITP Attendance System", ln=1, align="C")
    pdf.ln(2)
    pdf.set_font("Helvetica", size=11)
    meta_line = f"Date: {date_str}    Session: {session}    Room: {room}    Subject: {subject}    Students: {student_count}"
    pdf.cell(0, 6, meta_line, ln=1, align="C")
    pdf.ln(4)
    pdf.set_font("Helvetica", size=9)
    pdf.ln(section_spacing)

    # ---------- SECTION 2: STUDENT BOXES (start immediately after header on same page) ----------
    # Determine how many rows can fit on the current (first) page after header.
    first_page_available_h = page_h - pdf.get_y() - bottom_margin
    rows_per_col_first = max(1, int(first_page_available_h // cell_height))
    per_page_first = rows_per_col_first * cols

    # For subsequent pages (full pages), compute using top_margin
    full_page_available_h = page_h - top_margin - bottom_margin
    rows_per_col_full = max(1, int(full_page_available_h // cell_height))
    per_page_full = rows_per_col_full * cols

    last_box_bottom_on_page = None

    # We'll manually manage an index and page-by-page capacities (first page capacity differs)
    idx = 0
    page_index = 0
    while idx < len(rolls):
        if page_index == 0:
            # we are currently on the header page and may continue drawing boxes here
            items_on_page = min(per_page_first, len(rolls) - idx)
            rows_used = max(1, int(math.ceil(items_on_page / float(cols))))
            y_start = pdf.get_y() + 0.0
        else:
            pdf.add_page()
            items_on_page = min(per_page_full, len(rolls) - idx)
            rows_used = max(1, int(math.ceil(items_on_page / float(cols))))
            y_start = top_margin

        # draw rows_used x cols grid
        for r_idx in range(rows_used):
            for c_idx in range(cols):
                if idx >= len(rolls):
                    break
                roll = rolls[idx]
                x = left_margin + c_idx * (col_w + col_gap)
                y = y_start + r_idx * cell_height

                # outer border
                pdf.set_draw_color(0, 0, 0)
                pdf.set_line_width(0.6)
                pdf.rect(x, y, col_w, cell_height)

                # image box
                img_x = x + img_inner_pad
                img_y = y + img_inner_pad
                pdf.set_line_width(0.4)
                pdf.rect(img_x, img_y, img_box_size, img_box_size)

                # draw image to fixed size (keeps layout consistent)
                img_path = get_photo_for_roll(photos_dir, roll)
                try:
                    pdf.image(img_path, x=img_x + 1.0, y=img_y + 1.0, w=img_box_size - 2.0, h=img_box_size - 2.0)
                except Exception:
                    LOG.exception("Failed to draw image for %s; filling placeholder.", roll)
                    pdf.set_fill_color(245, 245, 245)
                    pdf.rect(img_x + 1, img_y + 1, img_box_size - 2, img_box_size - 2, style="F")

                # right text area
                text_x = img_x + img_box_size + 4.0
                text_w = x + col_w - 4.0 - text_x

                # Name (wrap explicitly)
                name_text = names.get(roll, "")
                pdf.set_font("Helvetica", size=9)
                pdf.set_xy(text_x, y + img_inner_pad)
                if name_text:
                    approx_char_w = pdf.get_string_width("a") or 2.6
                    max_chars = max(10, int(text_w / approx_char_w))
                    wrapped = textwrap.wrap(name_text, width=max_chars)
                    wrapped = wrapped[:name_max_lines]
                    for line_i, line in enumerate(wrapped):
                        pdf.set_xy(text_x, y + img_inner_pad + line_i * name_line_h)
                        pdf.cell(text_w, name_line_h, line, ln=0, align="L")

                # Roll (below name area)
                roll_y = y + img_inner_pad + (name_max_lines * name_line_h) + 1.0  # tighter spacing:
                pdf.set_xy(text_x, roll_y)
                pdf.set_font("Helvetica", "B", 10)
                max_roll_chars = max(6, int(text_w / (pdf.get_string_width("W") or 3.0)))
                roll_text = roll if len(roll) <= max_roll_chars else roll[:max_roll_chars - 3] + "..."
                pdf.cell(text_w, roll_line_h, roll_text, ln=1, align="L")

               
                # Signature underline + "Sign" text below it
                sig_line_y = roll_y + roll_line_h + 2.0   # reduced spacing
                sig_line_w = text_w * 0.6
                sig_x1 = text_x + (text_w - sig_line_w) / 2.0
                sig_x2 = sig_x1 + sig_line_w
                pdf.set_draw_color(0, 0, 0)
                pdf.set_line_width(0.5)
                pdf.line(sig_x1, sig_line_y, sig_x2, sig_line_y)

                # 'Sign' text on next line
                pdf.set_xy(text_x, sig_line_y + 1.8)
                pdf.set_font("Helvetica", size=8)
                pdf.cell(text_w, 4, "Sign", align="C")

                last_box_bottom_on_page = y + cell_height
                idx += 1
                page_index += 1

    rows_needed = student_count
    if rows_needed > 0:
        table_margin = left_margin
        table_w = page_w - 2 * table_margin
        col_si_w = table_w * 0.10
        col_name_w = table_w * 0.60
        col_sig_w = table_w * 0.30
        table_row_h = 9.0
        header_h = 8.0

        # Check if we can place table on the current (last) page right after last_box_bottom_on_page
        place_on_same_page = False
        if last_box_bottom_on_page is not None:
            remaining_space = page_h - last_box_bottom_on_page - bottom_margin
            if remaining_space >= (header_h + table_row_h + 4.0):
                place_on_same_page = True

        if place_on_same_page:
            start_y = last_box_bottom_on_page + section_spacing
            pdf.set_xy(table_margin, start_y)
        else:
            pdf.add_page()
            pdf.set_xy(table_margin, top_margin + 2.0)

        drawn = 0
        while drawn < rows_needed:
            # header
            pdf.set_font("Helvetica", "B", 10)
            pdf.cell(col_si_w, header_h, "Sl No.", border=1, align='C')
            pdf.cell(col_name_w, header_h, "Name", border=1, align='C')
            pdf.cell(col_sig_w, header_h, "Signature", border=1, align='C')
            pdf.ln(header_h)

            pdf.set_font("Helvetica", size=9)
            remaining_space = page_h - pdf.get_y() - bottom_margin
            per_page_rows = max(1, int(remaining_space // table_row_h))
            rows_this_page = min(per_page_rows, rows_needed - drawn)

            for i in range(rows_this_page):
                idx_no = drawn + 1
                pdf.cell(col_si_w, table_row_h, str(idx_no), border=1, align='C')
                pdf.cell(col_name_w, table_row_h, "", border=1, align='L')
                pdf.cell(col_sig_w, table_row_h, "", border=1, align='L')
                pdf.ln(table_row_h)
                drawn += 1

            if drawn < rows_needed:
                pdf.add_page()
                pdf.set_xy(table_margin, top_margin + 2.0)

    out_pdf = allocation_path.with_suffix(".pdf")
    pdf.output(str(out_pdf))
    LOG.info("Generated PDF: %s", out_pdf)
    return str(out_pdf)



def generate_pdfs_for_output_dir(output_root: str, photos_dir: str, cols: int = 3, paper_format: str = "A4"):
    """
    Recursively scan output_root for allocation .xlsx files and generate PDFs next to them.
    Skip summary/overall files (op_overall*, op_subjects_summary*, op_seats_left* etc).
    """
    root = Path(output_root)
    if not root.exists():
        raise FileNotFoundError(f"{output_root} does not exist")

    generated = []

    skip_prefixes = {
        "op_overall",          
        "op_subjects_summary", 
        "op_seating_arrangement",
        "op_subjects_summary",
        "op_seats_left",       
        "op_overall_seating_arrangement",
    }

    for p in sorted(root.rglob("*.xlsx")):
        name_lower = p.name.lower()
        # skip if file name starts with any of the summary prefixes OR is exactly a known summary name
        if any(name_lower.startswith(pref) for pref in skip_prefixes) or name_lower in {
            "op_seating_arrangement.xlsx", "op_overall_seating_arrangement.xlsx", "op_subjects_summary.xlsx"
        }:
            LOG.info("Skipping non-allocation file for PDF generation: %s", p.name)
            continue

        try:
            pdfp = _generate_pdf_for_allocation_file(p, photos_dir, cols=cols, paper_format=paper_format)
            generated.append(pdfp)
        except Exception as e:
            LOG.exception("Failed to generate PDF for %s: %s", p, e)
    return generated

