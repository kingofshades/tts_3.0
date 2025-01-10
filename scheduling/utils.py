import math
import pandas as pd

def build_sections_for_semester(semester_num, num_students, section_size=50):
    """
    e.g. if num_students=120 => 3 sections => [S1A, S1B, S1C].
    We embed the semester in the section name for clarity (S<sem><chr>).
    """
    count = math.ceil(num_students / section_size)
    return [f"S{semester_num}{chr(65+i)}" for i in range(count)]

def build_section_dataframe(
    section, courses, schedule_map,
    DAYS, THEORY_TIMESLOTS, TIMESLOT_LABELS,
    LAB_SLOTS, LAB_SLOT_LABELS,
    theory_rooms, lab_rooms, special_lab_rooms
):
    """
    Builds a tabular view of one section's schedule.
    """
    combined_labs = lab_rooms[:]
    for slist in special_lab_rooms.values():
        for labn in slist:
            if labn not in combined_labs:
                combined_labs.append(labn)

    rows = []
    for (code, cname, is_lab, _tn) in courses:
        day_map = {d: [] for d in DAYS}
        for day in DAYS:
            if not is_lab:
                # Theory => (day, timeslot, room)
                for t in THEORY_TIMESLOTS:
                    for r in theory_rooms:
                        val = schedule_map.get((day, t, r))
                        if val == (section, code):
                            label = TIMESLOT_LABELS.get(t, f"Slot {t}")
                            day_map[day].append(f"{r} [{label}]")
            else:
                # Lab => (day, lab_slot, lab_room)
                for ls in LAB_SLOTS:
                    for lr in combined_labs:
                        val = schedule_map.get((day, ls, lr))
                        if val == (section, code):
                            label = LAB_SLOT_LABELS.get(ls, f"LabSlot {ls}")
                            day_map[day].append(f"{lr} [{label}]")

        row_data = {
            "Course Code": code,
            "Course Name": cname,
            "Section": section
        }
        for d in DAYS:
            row_data[d] = ", ".join(day_map[d]) if day_map[d] else ""
        rows.append(row_data)

    df = pd.DataFrame(rows, columns=[
        "Course Code","Course Name","Section", *DAYS
    ])
    return df

def build_room_usage_df(
    room, schedule_map, is_lab,
    DAYS, THEORY_TIMESLOTS, TIMESLOT_LABELS,
    LAB_SLOTS, LAB_SLOT_LABELS
):
    """
    Similar to original: shows how a single room is used.
    """
    data = []
    if not is_lab:
        # theory usage
        for day in DAYS:
            row = {}
            for t in THEORY_TIMESLOTS:
                val = schedule_map.get((day, t, room))
                if val is None:
                    row[t] = "Free"
                else:
                    (sec, code) = val
                    label = f"{sec}-{code}"
                    row[t] = label
            data.append(row)
        df = pd.DataFrame(data, index=DAYS)
        df.columns = [TIMESLOT_LABELS.get(t, f"Slot {t}") for t in THEORY_TIMESLOTS]
    else:
        # lab usage
        for day in DAYS:
            row = {}
            for ls in LAB_SLOTS:
                val = schedule_map.get((day, ls, room))
                if val is None:
                    row[ls] = "Free"
                else:
                    (sec, code) = val
                    row[ls] = f"{sec}-{code}"
            df_columns = [LAB_SLOT_LABELS.get(ls, f"LabSlot {ls}") for ls in LAB_SLOTS]
            data.append(row)
        df = pd.DataFrame(data, index=DAYS)
        df.columns = df_columns
    return df

def export_timetables_to_excel(
    schedule_map, semester_sections_map, semester_courses_map,
    filename, DAYS, THEORY_TIMESLOTS, TIMESLOT_LABELS,
    LAB_SLOTS, LAB_SLOT_LABELS,
    theory_rooms, lab_rooms, special_lab_rooms
):
    import pandas as pd
    from openpyxl import Workbook

    combined_frames = []
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        for sem, sections in semester_sections_map.items():
            sem_frames = []
            for sec in sections:
                courses = semester_courses_map[sem]
                df_sec = build_section_dataframe(
                    sec, courses, schedule_map,
                    DAYS, THEORY_TIMESLOTS, TIMESLOT_LABELS,
                    LAB_SLOTS, LAB_SLOT_LABELS,
                    theory_rooms, lab_rooms, special_lab_rooms
                )
                df_sec.insert(0, "Semester", sem)
                sem_frames.append(df_sec)
                combined_frames.append(df_sec)

            if sem_frames:
                df_s = pd.concat(sem_frames, ignore_index=True)
                df_s.to_excel(writer, sheet_name=f"Semester_{sem}", index=False)

        if combined_frames:
            df_combined = pd.concat(combined_frames, ignore_index=True)
            df_combined.to_excel(writer, sheet_name="All_Sections", index=False)

    print(f"Timetables saved to {filename}")
