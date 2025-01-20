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

def parse_single_excel(file) -> (dict, list, list, dict):
    """
    Reads a single Excel with these sheets:
      - 'Roadmap' => columns: semester, course_code, course_name, is_lab, times_needed
      - 'Rooms' => columns: room_name, room_type
      - 'StudentCapacity' => columns: semester, student_count

    Returns:
      - semester_courses_map (dict[sem -> list[(code, cname, is_lab, times_needed)]])
      - theory_rooms (list[str])
      - lab_rooms (list[str])
      - student_capacities (dict[sem -> int])
    """
    xls = pd.ExcelFile(file)

    # 1) Roadmap
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

    # 2) Rooms
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

    # 3) StudentCapacity
    df_cap = pd.read_excel(xls, "StudentCapacity")

    # handle minor variations in casing/spaces
    df_cap.columns = df_cap.columns.str.lower().str.strip()

    if "semester" not in df_cap.columns or "student_count" not in df_cap.columns:
        raise ValueError("StudentCapacity sheet must have 'semester' and 'student_count' columns.")

    student_capacities = {}
    for _, row in df_cap.iterrows():
        sem = int(row["semester"])
        student_capacities[sem] = int(row["student_count"])

    return semester_courses_map, theory_rooms, lab_rooms, student_capacities
