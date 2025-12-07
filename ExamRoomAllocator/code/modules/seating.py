from pathlib import Path
import pandas as pd
from typing import IO
import logging
import re
import math

LOG = logging.getLogger(__name__)


def _read_workbook_bytesio(bio: IO):
    try:
        xls = pd.ExcelFile(bio, engine="openpyxl")
    except Exception as e:
        LOG.exception("Failed to open workbook bytes: %s", e)
        raise
    sheets = {}
    for name in xls.sheet_names:
        try:
            sheets[name] = pd.read_excel(xls, sheet_name=name, engine="openpyxl")
        except Exception:
            sheets[name] = None
    return sheets


def _infer_block(room_name: str):
    if not room_name:
        return ""
    if "-" in room_name:
        left = room_name.split("-", 1)[0]
        if left.isalpha():
            return left
        left = re.sub(r'[^A-Za-z]', '', left)
        return left
    m = re.match(r"^([A-Za-z]+)", room_name)
    if m:
        return m.group(1)
    return ""


def _infer_floor(room_name: str):
    if not room_name:
        return ""
    m = re.match(r"^(\d{1,2})", room_name)
    if m:
        return m.group(1)
    return ""


def _room_numeric_value(room_name: str) -> int:
    if not room_name:
        return 10 ** 9
    m = re.match(r"^(\d+)", room_name)
    if m:
        try:
            return int(m.group(1))
        except Exception:
            pass
    m2 = re.search(r"(\d+)", room_name)
    if m2:
        try:
            return int(m2.group(1))
        except Exception:
            pass
    return 10 ** 9

def process_master_excel(master_bio: IO, output_root: str, buffer: int = 0, filling_mode: str = "Dense"):
    out_dir = Path(output_root)
    out_dir.mkdir(parents=True, exist_ok=True)

    sheets = _read_workbook_bytesio(master_bio)
    schedule_df = sheets.get("Schedule")
    enrollment_df = sheets.get("Enrollment")
    rooms_df = sheets.get("Rooms")
    student_map_df = sheets.get("Student_Mapping")

    if schedule_df is None or enrollment_df is None or rooms_df is None:
        raise ValueError("Master workbook missing one of Schedule/Enrollment/Rooms sheets")

    # --- build subj_map (ordered roll lists, deduped) ---
    enrollment_df = enrollment_df.rename(columns={c: c.strip() for c in enrollment_df.columns})
    col_rename_map = {}
    for c in enrollment_df.columns:
        if c.lower().strip() == "roll":
            col_rename_map[c] = "Roll"
        elif c.lower().strip() == "subject":
            col_rename_map[c] = "Subject"
    if col_rename_map:
        enrollment_df = enrollment_df.rename(columns=col_rename_map)

    subj_map = {}
    for _, r in enrollment_df.iterrows():
        roll = str(r.get("Roll") or r.get("roll") or "").strip()
        subj = str(r.get("Subject") or r.get("subject") or "").strip()
        if not roll or not subj:
            continue
        subj_map.setdefault(subj, []).append(roll)
    # dedupe while preserving order
    for subj, lst in list(subj_map.items()):
        seen = set(); new=[]
        for rr in lst:
            if rr not in seen:
                seen.add(rr); new.append(rr)
        subj_map[subj]=new

    # --- roll -> name mapping ---
    roll_to_name = {}
    if student_map_df is not None:
        cols_low = {c.lower(): c for c in student_map_df.columns}
        roll_col = cols_low.get("roll") or cols_low.get("rollno") or list(student_map_df.columns)[0]
        name_col = cols_low.get("name") or cols_low.get("studentname") or (list(student_map_df.columns)[1] if len(student_map_df.columns) > 1 else None)
        for _, r in student_map_df.iterrows():
            try:
                rr = r.get(roll_col)
                nm = r.get(name_col) if name_col is not None else None
                if pd.isna(rr): continue
                rr_s = str(rr).strip()
                nm_s = str(nm).strip() if nm is not None and not pd.isna(nm) else ""
                if rr_s: roll_to_name[rr_s]=nm_s
            except Exception:
                continue

    # --- rooms -> room_capacity & meta (capacity, block, floor) ---
    room_capacity = {}
    rooms_meta = {}
    block_col_name = None
    for c in rooms_df.columns:
        if str(c).lower().strip() in {"block", "building", "block name"}:
            block_col_name = c
            break

    for _, r in rooms_df.iterrows():
        room_name = str(r.get("Room") or r.get("room") or "").strip()
        cap=None
        for cand in ("Capacity","capacity","Exam Capacity","ExamCapacity","Exam Cap"):
            if cand in r.index:
                cap = r.get(cand); break
        try:
            cap_i = int(float(cap)) if cap is not None and not pd.isna(cap) else 0
        except Exception:
            try: cap_i = int(re.findall(r"\d+", str(cap or "0"))[0])
            except Exception: cap_i = 0
        
        try:
            buf_i = int(buffer) if buffer is not None else 0
        except Exception:
            buf_i = 0

        present_capacity = cap_i
        if str(filling_mode).strip().lower() == "sparse":
            # half the room capacity when Sparse
            present_capacity = int(math.floor(cap_i / 2.0))

        # apply buffer reduction after deciding present capacity
        remaining = max(0, present_capacity - buf_i)

        block = str(r.get(block_col_name) or "").strip() if block_col_name else ""
        floor = _infer_floor(room_name)
        room_capacity[room_name] = remaining
        rooms_meta[room_name] = {"capacity": cap_i, "block": block, "floor": floor}

    # preserve original capacities (after buffer) so we can reset per shift/day
    original_room_capacity = dict(room_capacity)

    # Build rooms_by_floor & floor ordering (total remaining desc)
    rooms_by_floor = {}
    for rn, meta in rooms_meta.items():
        fl = meta["floor"] or "_NO_FLOOR"
        rooms_by_floor.setdefault(fl, []).append(rn)
    # sort per-floor rooms by capacity desc initially, tie-break numeric ascending
    for fl in rooms_by_floor:
        rooms_by_floor[fl].sort(key=lambda rn: (rooms_meta[rn]["capacity"], -_room_numeric_value(rn)), reverse=True)
    floor_total = {fl: sum(original_room_capacity.get(rn, 0) for rn in rlist) for fl, rlist in rooms_by_floor.items()}
    floors_order = sorted(floor_total.keys(), key=lambda f: floor_total[f], reverse=True)

    per_date_shift_fill_order = {}
    per_date_output_map = {}
    overall_output_map = {}         # overall map aggregated across dates
    overall_rows = []               # for op_overall_seating_arrangement rows

    # Normalize schedule date & shift columns
    schedule_df = schedule_df.copy()
    if "Date" not in schedule_df.columns:
        raise ValueError("Schedule sheet must contain a Date column")
    def _normalize_date_cell(d):
        try:
            return pd.to_datetime(d).date()
        except Exception:
            return d
    schedule_df["__alloc_date"] = schedule_df["Date"].apply(_normalize_date_cell)

    # normalize shift text into a column to preserve order of shifts per date
    def _normalize_shift_cell(s):
        if pd.isna(s):
            return "Unknown"
        return str(s).strip() or "Unknown"
    # try to find shift/Session column names (case-insensitive)
    shift_col_candidates = [c for c in schedule_df.columns if str(c).lower().strip() in {"session", "shift", "session/shift"}]
    if shift_col_candidates:
        sched_shift_col = shift_col_candidates[0]
    else:
        sched_shift_col = None
    if sched_shift_col:
        schedule_df["__shift_norm"] = schedule_df[sched_shift_col].apply(_normalize_shift_cell)
    else:
        schedule_df["__shift_norm"] = schedule_df.apply(lambda _: "Unknown", axis=1)

    for date_key, date_group in schedule_df.groupby("__alloc_date", sort=False):
        shifts_series = date_group["__shift_norm"]
        shifts_in_order = []
        for v in shifts_series.tolist():
            if v not in shifts_in_order:
                shifts_in_order.append(v)

        # iterate shifts for this date (capacities reset at each shift start)
        for shift in shifts_in_order:
            # Reset capacities at start of each shift
            room_capacity = dict(original_room_capacity)

            # rows for this date & shift in original order
            rows_for_shift = date_group[date_group["__shift_norm"] == shift]

            # canonical folder names
            try:
                date_dt = pd.to_datetime(date_key)
                date_str_for_group = date_dt.strftime("%Y_%m_%d")
                date_for_sheet = date_dt.strftime("%d/%m/%Y")
            except Exception:
                date_str_for_group = str(date_key)
                date_for_sheet = str(date_key)
            safe_session = "".join(ch if ch.isalnum() or ch in ("-", "_") else "_" for ch in str(shift)) or "Unknown"

            # initialize maps for this date+shift
            per_date_shift_fill_order.setdefault(date_str_for_group, {})
            per_date_shift_fill_order[date_str_for_group].setdefault(shift, [])
            per_date_output_map.setdefault(date_str_for_group, {})
            per_date_output_map[date_str_for_group].setdefault(shift, {})

            # process each scheduled subject row for this shift
            for _, sched in rows_for_shift.iterrows():
                session = shift
                subject = sched.get("Subject") or sched.get("subject") or ""
                if pd.isna(subject) or subject == "":
                    continue
                subject = str(subject).strip()

                rolls = subj_map.get(subject, []).copy()
                if not rolls:
                    LOG.info("No enrolled students for subject %s on %s %s", subject, date_key, session)
                    continue
                remaining_rolls = rolls

                assigned_chunks = []

                # Phase A: floor-first allocation (prefer larger rooms on floors first)
                for fl in floors_order:
                    if not remaining_rolls:
                        break
                    floor_rooms = sorted(
                        [(rn, room_capacity.get(rn, 0), rooms_meta[rn]["capacity"]) for rn in rooms_by_floor.get(fl, [])],
                        key=lambda x: (x[1], x[2], -_room_numeric_value(x[0])),
                        reverse=True
                    )
                    for rname, rem_cap, cap in floor_rooms:
                        if not remaining_rolls:
                            break
                        if rem_cap <= 0:
                            continue
                        take = min(rem_cap, len(remaining_rolls))
                        if take <= 0:
                            continue
                        chunk = remaining_rolls[:take]
                        remaining_rolls = remaining_rolls[take:]
                        room_capacity[rname] = room_capacity.get(rname, 0) - take

                        assigned_chunks.append((rname, chunk))

                        # record fill order for date+shift
                        if rname not in per_date_shift_fill_order[date_str_for_group][shift]:
                            per_date_shift_fill_order[date_str_for_group][shift].append(rname)

                        # update per_date_output_map for date->shift->room
                        per_date_output_map[date_str_for_group].setdefault(shift, {})
                        per_date_output_map[date_str_for_group][shift].setdefault(rname, {"subjects": {}, "rolls": []})
                        per_date_output_map[date_str_for_group][shift][rname]["subjects"].setdefault(subject, [])
                        per_date_output_map[date_str_for_group][shift][rname]["subjects"][subject].extend(chunk)
                        per_date_output_map[date_str_for_group][shift][rname]["rolls"].extend(chunk)

                        # update overall_output_map aggregated by room
                        overall_output_map.setdefault(rname, {"subjects": {}, "rolls": []})
                        overall_output_map[rname]["subjects"].setdefault(subject, [])
                        overall_output_map[rname]["subjects"][subject].extend(chunk)
                        overall_output_map[rname]["rolls"].extend(chunk)

                # Phase B: global exhaustive fill (fill all rooms in passes; picks room with current largest remaining)
                if remaining_rolls:
                    while remaining_rolls:
                        current_rooms = sorted(
                            [(rn, room_capacity.get(rn, 0), rooms_meta[rn]["capacity"]) for rn in rooms_meta.keys()],
                            key=lambda x: (x[1], x[2], -_room_numeric_value(x[0])),
                            reverse=True
                        )
                        any_assigned_this_pass = False
                        for rname, rem_cap, cap in current_rooms:
                            if not remaining_rolls:
                                break
                            if rem_cap <= 0:
                                continue
                            take = min(rem_cap, len(remaining_rolls))
                            if take <= 0:
                                continue
                            chunk = remaining_rolls[:take]
                            remaining_rolls = remaining_rolls[take:]
                            room_capacity[rname] = room_capacity.get(rname, 0) - take

                            assigned_chunks.append((rname, chunk))

                            if rname not in per_date_shift_fill_order[date_str_for_group][shift]:
                                per_date_shift_fill_order[date_str_for_group][shift].append(rname)

                            per_date_output_map[date_str_for_group].setdefault(shift, {})
                            per_date_output_map[date_str_for_group][shift].setdefault(rname, {"subjects": {}, "rolls": []})
                            per_date_output_map[date_str_for_group][shift][rname]["subjects"].setdefault(subject, [])
                            per_date_output_map[date_str_for_group][shift][rname]["subjects"][subject].extend(chunk)
                            per_date_output_map[date_str_for_group][shift][rname]["rolls"].extend(chunk)

                            overall_output_map.setdefault(rname, {"subjects": {}, "rolls": []})
                            overall_output_map[rname]["subjects"].setdefault(subject, [])
                            overall_output_map[rname]["subjects"][subject].extend(chunk)
                            overall_output_map[rname]["rolls"].extend(chunk)

                            any_assigned_this_pass = True

                        if not any_assigned_this_pass:
                            break

                # Final exhaustive sweep (safety) before marking overflow
                overflow_rolls = []
                if remaining_rolls:
                    total_remaining_seats = sum(max(0, v) for v in room_capacity.values())
                    if total_remaining_seats > 0:
                        for rname, rem in sorted(room_capacity.items(), key=lambda x: (x[1], -_room_numeric_value(x[0])), reverse=True):
                            if not remaining_rolls:
                                break
                            if rem <= 0:
                                continue
                            take = min(rem, len(remaining_rolls))
                            chunk = remaining_rolls[:take]
                            remaining_rolls = remaining_rolls[take:]
                            room_capacity[rname] = room_capacity.get(rname, 0) - take

                            assigned_chunks.append((rname, chunk))
                            if rname not in per_date_shift_fill_order[date_str_for_group][shift]:
                                per_date_shift_fill_order[date_str_for_group][shift].append(rname)

                            per_date_output_map[date_str_for_group].setdefault(shift, {})
                            per_date_output_map[date_str_for_group].setdefault(rname, {"subjects": {}, "rolls": []})
                            per_date_output_map[date_str_for_group][shift].setdefault(rname, {"subjects": {}, "rolls": []})
                            per_date_output_map[date_str_for_group][shift][rname]["subjects"].setdefault(subject, [])
                            per_date_output_map[date_str_for_group][shift][rname]["subjects"][subject].extend(chunk)
                            per_date_output_map[date_str_for_group][shift][rname]["rolls"].extend(chunk)

                            overall_output_map.setdefault(rname, {"subjects": {}, "rolls": []})
                            overall_output_map[rname]["subjects"].setdefault(subject, [])
                            overall_output_map[rname]["subjects"][subject].extend(chunk)
                            overall_output_map[rname]["rolls"].extend(chunk)

                    if remaining_rolls:
                        overflow_rolls = remaining_rolls
                        remaining_rolls = []

                # Write per-room allocation files (in assigned order) for this subject & shift
                for rname, chunk in assigned_chunks:
                    target_dir = Path(out_dir) / date_str_for_group / (safe_session or "Unknown")
                    target_dir.mkdir(parents=True, exist_ok=True)
                    safe_subject = "".join(ch if ch.isalnum() or ch in ("-", "_") else "_" for ch in subject)
                    filename = f"{date_str_for_group}_{safe_session}_{rname}_{safe_subject}.xlsx"
                    out_path = target_dir / filename
                    out_rows = []
                    for rr in chunk:
                        out_rows.append({
                            "Roll": rr,
                            "Name": roll_to_name.get(rr, ""),
                            "Subject": subject,
                            "Room": rname,
                            "Date": pd.to_datetime(date_key).date() if pd.notna(date_key) else str(date_key),
                            "Shift": session
                        })
                    if out_rows:
                        try:
                            pd.DataFrame(out_rows, columns=["Roll", "Name", "Subject", "Room", "Date", "Shift"]).to_excel(out_path, index=False)
                        except Exception:
                            LOG.exception("Failed to write allocation file %s", out_path)

                        # record overall rows for final overall file
                        roll_list = ";".join([r["Roll"] for r in out_rows])
                        overall_rows.append({
                            "Date": str(out_rows[0]["Date"]),
                            "Day": (pd.to_datetime(out_rows[0]["Date"]).strftime("%A") if pd.notna(out_rows[0]["Date"]) else ""),
                            "course_code": subject,
                            "Room": rname,
                            "Allocated_students_count": len(out_rows),
                            "Roll_list (semicolon separated_)": roll_list
                        })

                # Write overflow file if any and update maps
                if overflow_rolls:
                    total_room_capacity = sum(original_room_capacity.values())
                    allocated_so_far = sum(len(v.get("rolls", [])) for shift_map in per_date_output_map.get(date_str_for_group, {}).values() for v in shift_map.values())
                    LOG.warning("Overflow for subject %s on %s %s: students_remaining=%d total_room_capacity=%d total_allocated_so_far=%d",
                                subject, date_str_for_group, session, len(overflow_rolls), total_room_capacity, allocated_so_far)

                    target_dir = Path(out_dir) / date_str_for_group / (safe_session or "Unknown")
                    target_dir.mkdir(parents=True, exist_ok=True)
                    safe_subject = "".join(ch if ch.isalnum() or ch in ("-", "_") else "_" for ch in subject)
                    filename = f"{date_str_for_group}_{safe_session}_OVERFLOW_{safe_subject}.xlsx"
                    out_path = target_dir / filename
                    out_rows = []
                    for rr in overflow_rolls:
                        out_rows.append({
                            "Roll": rr,
                            "Name": roll_to_name.get(rr, ""),
                            "Subject": subject,
                            "Room": "OVERFLOW",
                            "Date": pd.to_datetime(date_key).date() if pd.notna(date_key) else str(date_key),
                            "Shift": session
                        })
                    if out_rows:
                        try:
                            pd.DataFrame(out_rows, columns=["Roll", "Name", "Subject", "Room", "Date", "Shift"]).to_excel(out_path, index=False)
                        except Exception:
                            LOG.exception("Failed to write OVERFLOW file %s", out_path)

                        # update per_date_output_map & overall_output_map with overflow bucket
                        per_date_output_map[date_str_for_group].setdefault(shift, {})
                        per_date_output_map[date_str_for_group][shift].setdefault("OVERFLOW", {"subjects": {}, "rolls": []})
                        per_date_output_map[date_str_for_group][shift]["OVERFLOW"]["subjects"].setdefault(subject, [])
                        per_date_output_map[date_str_for_group][shift]["OVERFLOW"]["subjects"][subject].extend([r["Roll"] for r in out_rows])
                        per_date_output_map[date_str_for_group][shift]["OVERFLOW"]["rolls"].extend([r["Roll"] for r in out_rows])

                        overall_output_map.setdefault("OVERFLOW", {"subjects": {}, "rolls": []})
                        overall_output_map["OVERFLOW"]["subjects"].setdefault(subject, [])
                        overall_output_map["OVERFLOW"]["subjects"][subject].extend([r["Roll"] for r in out_rows])
                        overall_output_map["OVERFLOW"]["rolls"].extend([r["Roll"] for r in out_rows])

                        roll_list = ";".join([r["Roll"] for r in out_rows])
                        overall_rows.append({
                            "Date": str(out_rows[0]["Date"]),
                            "Day": (pd.to_datetime(out_rows[0]["Date"]).strftime("%A") if pd.notna(out_rows[0]["Date"]) else ""),
                            "course_code": subject,
                            "Room": "OVERFLOW",
                            "Allocated_students_count": len(out_rows),
                            "Roll_list (semicolon separated_)": roll_list
                        })

    # -------------------- Post-scan: per-date shift-wise seats left files --------------------
    for date_str in sorted(per_date_output_map.keys()):
        date_folder = Path(out_dir) / date_str
        date_folder.mkdir(parents=True, exist_ok=True)

        # Gather allocation files for this date (exclude summary files)
        alloc_files = [p for p in sorted(date_folder.rglob("*.xlsx")) if not p.name.lower().startswith("op_")]

        date_shift_room_map = {}
        shift_first_seen_rooms = {}
        for p in alloc_files:
            try:
                df = pd.read_excel(p, engine="openpyxl")
            except Exception:
                LOG.exception("Skipping unreadable allocation file %s", p)
                continue

            # determine shift for this allocation file
            shift_val = None
            if "Shift" in df.columns:
                shift_val = df.iloc[0].get("Shift")
            elif "Session" in df.columns:
                shift_val = df.iloc[0].get("Session")
            # fallback: parent folder name (session folder)
            if shift_val is None or (isinstance(shift_val, float) and pd.isna(shift_val)) or str(shift_val).strip() == "":
                # parent might be .../date/<session>/file.xlsx
                parent_dirs = list(p.parents)
                shift_val = next((d.name for d in parent_dirs if d.parent and d.parent.name == date_folder.name and d.name), None)
            if shift_val is None:
                shift_val = "Unknown"
            shift_val = str(shift_val).strip()

            room_col = None
            if "Room" in df.columns:
                room_col = "Room"
            elif "room" in df.columns:
                room_col = "room"

            if room_col and "Roll" in df.columns:
                # multi-row allocation
                for _, row in df.iterrows():
                    rn = str(row.get(room_col) or "").strip()
                    rl = str(row.get("Roll") or "").strip()
                    if not rn:
                        continue
                    date_shift_room_map.setdefault(shift_val, {})
                    date_shift_room_map[shift_val].setdefault(rn, {"rolls": [], "first_seen": None})
                    if rl:
                        date_shift_room_map[shift_val][rn]["rolls"].append(rl)
                        if date_shift_room_map[shift_val][rn]["first_seen"] is None:
                            # record order by file read order
                            date_shift_room_map[shift_val][rn]["first_seen"] = len(shift_first_seen_rooms.get(shift_val, []))
                            shift_first_seen_rooms.setdefault(shift_val, [])
                            if rn not in shift_first_seen_rooms[shift_val]:
                                shift_first_seen_rooms[shift_val].append(rn)
            else:
                # legacy single-row format with a roll-list cell
                lr = None
                for cand in ("Roll_List", "RollList", "roll_list", "roll list"):
                    if cand in df.columns:
                        lr = str(df.iloc[0].get(cand) or "")
                        break
                if not lr:
                    # can't extract rolls from this file - skip it
                    continue
                rolls = [s.strip() for s in re.split(r"[;,|\n]", lr) if s and s.strip()]
                # try to infer room from filename pattern: YYYY_session_room_subject.xlsx
                parts = p.stem.split("_")
                room_name = parts[2] if len(parts) >= 3 else ""
                rn = str(room_name).strip()
                if not rn:
                    continue
                date_shift_room_map.setdefault(shift_val, {})
                date_shift_room_map[shift_val].setdefault(rn, {"rolls": [], "first_seen": None})
                date_shift_room_map[shift_val][rn]["rolls"].extend(rolls)
                if date_shift_room_map[shift_val][rn]["first_seen"] is None:
                    shift_first_seen_rooms.setdefault(shift_val, [])
                    date_shift_room_map[shift_val][rn]["first_seen"] = len(shift_first_seen_rooms[shift_val])
                    if rn not in shift_first_seen_rooms[shift_val]:
                        shift_first_seen_rooms[shift_val].append(rn)

        for shift_val in sorted(date_shift_room_map.keys()):
            # build ordered_rooms for this shift:
            fill_order = shift_first_seen_rooms.get(shift_val, [])
            used_set = set(fill_order)

            all_rooms = [(rn, rooms_meta[rn]["capacity"], rooms_meta[rn]["block"]) for rn in rooms_meta.keys()]
            # remaining rooms sorted by capacity desc, then numeric tie-break
            remaining_rooms_sorted = [
                t for t in sorted(all_rooms, key=lambda x: (x[1], -_room_numeric_value(x[0])), reverse=True)
                if t[0] not in used_set
            ]
            ordered_rooms = [(rn, rooms_meta[rn]["capacity"], rooms_meta[rn]["block"]) for rn in fill_order if rn in rooms_meta] + remaining_rooms_sorted

            # prepare rows: allocated counts come only from this shift's mapping
            seats_rows = []
            for rn, cap_i, block_val in ordered_rooms:
                allocated = len(date_shift_room_map.get(shift_val, {}).get(rn, {}).get("rolls", []))
                vacant = max(0, cap_i - allocated)
                seats_rows.append({
                    "Room No.": rn,
                    "Exam Capacity": cap_i,
                    "Block": block_val,
                    "Alloted": allocated,
                    "Vacant (B-C)": vacant
                })

            seats_df = pd.DataFrame(seats_rows, columns=["Room No.", "Exam Capacity", "Block", "Alloted", "Vacant (B-C)"])

            # safe filename for shift: lowercase, replace spaces with underscore
            safe_shift = re.sub(r"\s+", "_", shift_val.strip().lower()) or "unknown"
            seats_out_name = f"op_seats_left_{safe_shift}.xlsx"
            try:
                seats_out_path = date_folder / seats_out_name
                with pd.ExcelWriter(seats_out_path, engine="openpyxl") as writer:
                    seats_df.to_excel(writer, sheet_name="op_seats_left", index=False)
                LOG.info("Wrote datewise seats summary for %s (%s) to %s", date_str, shift_val, seats_out_path)
            except Exception:
                LOG.exception("Failed to write datewise op_seats_left for %s %s", date_folder, shift_val)

    # -------------------- Build overall op_overall_seating_arrangement.xlsx --------------------
    try:
        overall_df = pd.DataFrame(overall_rows, columns=[
            "Date", "Day", "course_code", "Room", "Allocated_students_count", "Roll_list (semicolon separated_)"
        ]) if overall_rows else pd.DataFrame(columns=[
            "Date", "Day", "course_code", "Room", "Allocated_students_count", "Roll_list (semicolon separated_)"
        ])
        overall_out_path = out_dir / "op_overall_seating_arrangement.xlsx"
        with pd.ExcelWriter(overall_out_path, engine="openpyxl") as writer:
            overall_df.to_excel(writer, sheet_name="op_overall_seating_arrangement", index=False)
        LOG.info("Wrote overall seating arrangement (from mapping) to %s", overall_out_path)
    except Exception:
        LOG.exception("Failed to write op_overall_seating_arrangement.xlsx")

    LOG.info("Seating allocation completed; files written to %s", out_dir)
    return str(out_dir)

