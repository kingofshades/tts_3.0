# utils.py
import math
import pandas as pd

def build_sections_for_semester(semester_num, num_students, section_size=50, program_code="A"):
    """
    If num_students=120 and section_size=50 => returns e.g. 3 sections.
    For example, if program_code='A', and sem=1, you get [S1A1, S1A2, S1A3].
    """
    count = math.ceil(num_students / section_size)
    sections = []
    for i in range(count):
        # e.g. S1A1, S1A2, ...
        sections.append(f"S{semester_num}{program_code}{i+1}")
    return sections

def build_section_dataframe(
    section_name,
    courses,
    schedule_map,
    DAYS,
    THEORY_TIMESLOTS,
    TIMESLOT_LABELS,
    LAB_SLOTS,
    LAB_SLOT_LABELS,
    theory_rooms,
    lab_rooms,
    special_lab_rooms
):
    """
    Build a DataFrame for a single section's assigned timetable.
    Each row => (Course Code, Course Name, Section, Monday, Tuesday, ...)
    """
    # Combine normal + special labs so we can detect them in schedule_map
    combined_labs = list(lab_rooms)
    for slist in special_lab_rooms.values():
        for labn in slist:
            if labn not in combined_labs:
                combined_labs.append(labn)

    rows = []
    for (code, cname, is_lab, _tn) in courses:
        day_map = {d: [] for d in DAYS}
        occupant_id = (section_name, code)  # how occupant is stored in schedule_map

        # We'll scan schedule_map for day/slot/room => occupant
        for (day, slot, room), occupant in schedule_map.items():
            if occupant == occupant_id:
                if is_lab and room in combined_labs and slot in LAB_SLOTS:
                    label = LAB_SLOT_LABELS.get(slot, f"LabSlot {slot}")
                    day_map[day].append(f"{room} [{label}]")
                elif (not is_lab) and room in theory_rooms and slot in THEORY_TIMESLOTS:
                    label = TIMESLOT_LABELS.get(slot, f"Slot {slot}")
                    day_map[day].append(f"{room} [{label}]")

        row_data = {
            "Course Code": code,
            "Course Name": cname,
            "Section": section_name
        }
        for d in DAYS:
            row_data[d] = ", ".join(day_map[d]) if day_map[d] else ""
        rows.append(row_data)

    df = pd.DataFrame(rows, columns=["Course Code","Course Name","Section", *DAYS])
    return df

def build_room_usage_df(
    room,
    schedule_map,
    is_lab,
    DAYS,
    THEORY_TIMESLOTS,
    TIMESLOT_LABELS,
    LAB_SLOTS,
    LAB_SLOT_LABELS
):
    """
    For a given room, build a usage table by day vs timeslot (theory) or lab_slot (lab).
    Index => DAYS, Columns => Timeslots
    """
    data = []
    if not is_lab:
        # theory usage
        for day in DAYS:
            row = {}
            for t in THEORY_TIMESLOTS:
                occupant = schedule_map.get((day, t, room))
                if occupant is None:
                    row[t] = "Free"
                else:
                    (sec, code) = occupant
                    row[t] = f"{sec}-{code}"
            data.append(row)
        df = pd.DataFrame(data, index=DAYS)
        df.columns = [TIMESLOT_LABELS.get(t, f"Slot {t}") for t in THEORY_TIMESLOTS]
    else:
        # lab usage
        for day in DAYS:
            row = {}
            for ls in LAB_SLOTS:
                occupant = schedule_map.get((day, ls, room))
                if occupant is None:
                    row[ls] = "Free"
                else:
                    (sec, code) = occupant
                    row[ls] = f"{sec}-{code}"
            data.append(row)
        df = pd.DataFrame(data, index=DAYS)
        df.columns = [LAB_SLOT_LABELS.get(ls, f"LabSlot {ls}") for ls in LAB_SLOTS]

    return df

def export_timetables_to_excel(
    schedule_map,
    semester_sections_map,
    semester_courses_map,
    filename,
    DAYS,
    THEORY_TIMESLOTS,
    TIMESLOT_LABELS,
    LAB_SLOTS,
    LAB_SLOT_LABELS,
    theory_rooms,
    lab_rooms,
    special_lab_rooms
):
    """
    Write each semester's timetable to a separate sheet,
    plus an "All_Sections" sheet in the same Excel.
    """
    import pandas as pd
    from openpyxl import Workbook

    writer = pd.ExcelWriter(filename, engine="openpyxl")
    combined_frames = []

    for sem, sections in semester_sections_map.items():
        frames = []
        for sec in sections:
            courses = semester_courses_map[sem]
            df_sec = build_section_dataframe(
                section_name=sec,
                courses=courses,
                schedule_map=schedule_map,
                DAYS=DAYS,
                THEORY_TIMESLOTS=THEORY_TIMESLOTS,
                TIMESLOT_LABELS=TIMESLOT_LABELS,
                LAB_SLOTS=LAB_SLOTS,
                LAB_SLOT_LABELS=LAB_SLOT_LABELS,
                theory_rooms=theory_rooms,
                lab_rooms=lab_rooms,
                special_lab_rooms=special_lab_rooms
            )
            df_sec.insert(0, "Semester", sem)
            frames.append(df_sec)
            combined_frames.append(df_sec)

        if frames:
            df_sem = pd.concat(frames, ignore_index=True)
            df_sem.to_excel(writer, sheet_name=f"Semester_{sem}", index=False)

    if combined_frames:
        df_all = pd.concat(combined_frames, ignore_index=True)
        df_all.to_excel(writer, sheet_name="All_Sections", index=False)

    writer.close()
    print(f"Timetables exported to {filename}")

def build_full_room_usage_df(
    room: str,
    rtype: str,           # "theory" or "lab"
    usage_data: dict,     # from usage_data.json
    schedule_map: dict,   # new allocations
    DAYS: list,
    THEORY_TIMESLOTS: list,
    LAB_SLOTS: list,
    TIMESLOT_LABELS: dict,
    LAB_SLOT_LABELS: dict
):
    """
    Returns a DataFrame that merges 'old' usage from usage_data.json
    with occupant labels from schedule_map. For old usage, we only
    know the slot was occupied, not the occupant name.

    If a slot was previously occupied, we show '(Previously Occupied)'.
    If schedule_map has a new occupant, we show the occupant code (e.g. 'S4DS2-CC2141').
    If a slot is both old and new (shouldn't happen if usage_data is correct),
    we'll display the new occupant label, because the solver wouldn't re-use that slot.
    Otherwise, 'Free'.
    """
    # total days
    data = []

    if rtype == "theory":
        slot_list = THEORY_TIMESLOTS
        slot_labels = TIMESLOT_LABELS
        old_usage_dict = usage_data["theory"]
    else:
        slot_list = LAB_SLOTS
        slot_labels = LAB_SLOT_LABELS
        old_usage_dict = usage_data["lab"]

    # Prepare usage dictionary for old usage
    # old_usage_dict looks like: usage_data["theory"] = { "R4": {"Monday": [0,1], "Tuesday": [...]}, ... }
    # We'll see if `room` is in old_usage_dict, then find which days/slots.
    old_room_usage = old_usage_dict.get(room, {})

    # Now build a day-based table
    for day in DAYS:
        row = {}
        # for each slot in the relevant timeslots
        for slot in slot_list:
            # 1) Check old usage
            previously_used = False
            if day in old_room_usage and slot in old_room_usage[day]:
                previously_used = True

            # 2) Check new occupant (schedule_map)
            occupant = None
            # schedule_map for theory => key=(day, timeslot, room) => occupant=(sec,code)
            # schedule_map for lab => key=(day, labslot, room) => occupant=(sec,code)
            # occupant is a tuple, e.g. ("S4DS2", "CC2141")
            if rtype == "theory":
                occupant_tuple = schedule_map.get((day, slot, room), None)
            else:
                occupant_tuple = schedule_map.get((day, slot, room), None)

            if occupant_tuple:
                sec_name, course_code = occupant_tuple
                occupant = f"{sec_name}-{course_code}"  # e.g. "S4DS2-CC2141"

            # 3) Decide what to display
            if occupant is not None:
                # new occupant in this slot
                row[slot] = occupant
            else:
                # occupant is None => check old usage
                if previously_used:
                    row[slot] = "(Previously Occupied)"
                else:
                    row[slot] = "Free"

        data.append(row)

    # Build the DataFrame
    df = pd.DataFrame(data, index=DAYS)
    # rename columns from slot => label
    new_cols = {}
    for s in slot_list:
        if rtype == "theory":
            new_cols[s] = TIMESLOT_LABELS.get(s, f"Slot {s}")
        else:
            new_cols[s] = LAB_SLOT_LABELS.get(s, f"LabSlot {s}")

    df.rename(columns=new_cols, inplace=True)

    return df
