# data_io.py
import os
import json
import pandas as pd

def load_usage(json_path: str) -> dict:
    """
    Loads usage data from the specified JSON file.
    If file doesn't exist, return empty usage.
    """
    if not os.path.exists(json_path):
        return {"theory": {}, "lab": {}}

    with open(json_path, "r", encoding="utf-8") as f:
        return json.load(f)

def save_usage(json_path: str, usage_data: dict):
    """
    Saves the usage dictionary to JSON.
    """
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(usage_data, f, indent=2)

def reset_usage(json_path: str):
    """
    Resets usage data to an empty structure.
    """
    empty_data = {"theory": {}, "lab": {}}
    save_usage(json_path, empty_data)

def parse_single_excel(file):
    """
    Reads a single Excel with these sheets:
      - 'Roadmap' => columns: semester, course_code, course_name, is_lab, times_needed
      - 'Rooms' => columns: room_name, room_type
      - 'StudentCapacity' => columns: semester, student_count
      - 'Electives' (optional) => columns: 
          [elective_code, elective_name, elective_type, sections_count, can_use_theory, can_use_lab]

    Returns a tuple:
      semester_courses_map, theory_rooms, lab_rooms, student_capacities, electives_list
    """
    xls = pd.ExcelFile(file)

    # -------------------------------
    # 1) Roadmap
    # -------------------------------
    df_roadmap = pd.read_excel(xls, "Roadmap")
    semester_courses_map = {}
    for _, row in df_roadmap.iterrows():
        sem = int(row["semester"])
        code = str(row["course_code"]).strip()
        cname = str(row["course_name"]).strip()
        is_lab = (str(row["is_lab"]).lower() == "true")
        times_needed = int(row["times_needed"])
        if sem not in semester_courses_map:
            semester_courses_map[sem] = []
        semester_courses_map[sem].append((code, cname, is_lab, times_needed))

    # -------------------------------
    # 2) Rooms
    # -------------------------------
    df_rooms = pd.read_excel(xls, "Rooms")
    theory_rooms = []
    lab_rooms = []
    for _, row in df_rooms.iterrows():
        rname = str(row["room_name"]).strip()
        rtype = str(row["room_type"]).strip().lower()
        if rtype == "theory":
            theory_rooms.append(rname)
        else:
            lab_rooms.append(rname)

    # -------------------------------
    # 3) StudentCapacity
    # -------------------------------
    df_cap = pd.read_excel(xls, "StudentCapacity")
    df_cap.columns = df_cap.columns.str.lower().str.strip()
    if "semester" not in df_cap.columns or "student_count" not in df_cap.columns:
        raise ValueError("StudentCapacity sheet must have 'semester' and 'student_count' columns.")

    student_capacities = {}
    for _, row in df_cap.iterrows():
        sem = int(row["semester"])
        student_capacities[sem] = int(row["student_count"])

    # -------------------------------
    # 4) Electives (optional)
    # -------------------------------
    electives_list = []
    if "Electives" in xls.sheet_names:
        df_elec = pd.read_excel(xls, "Electives")
        # Clean up column names
        df_elec.columns = df_elec.columns.str.lower().str.strip()

        # For each row, collect the info
        for _, row in df_elec.iterrows():
            e_code = str(row["elective_code"]).strip()
            e_name = str(row["elective_name"]).strip()
            e_type = str(row["elective_type"]).strip()
            e_sections = int(row["sections_count"])

            # If columns are missing, default them to True
            can_theory = True
            can_lab = True
            if "can_use_theory" in df_elec.columns:
                can_theory = str(row["can_use_theory"]).lower() == "true"
            if "can_use_lab" in df_elec.columns:
                can_lab = str(row["can_use_lab"]).lower() == "true"

            # Build a tuple
            electives_list.append({
                "code": e_code,
                "name": e_name,
                "etype": e_type,
                "sections_count": e_sections,
                "can_theory": can_theory,
                "can_lab": can_lab
            })

    return semester_courses_map, theory_rooms, lab_rooms, student_capacities, electives_list
