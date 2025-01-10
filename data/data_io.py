import csv
import json
import os

def load_rooms(csv_path: str) -> dict:
    """
    Reads 'rooms.csv' => returns a dict with 'theory_rooms' and 'lab_rooms'.
    CSV format:
      room_name,room_type
      R1,theory
      Lab1,lab
      ...
    """
    if not os.path.exists(csv_path):
        raise FileNotFoundError(f"Rooms CSV not found at {csv_path}")

    theory_rooms = []
    lab_rooms = []
    with open(csv_path, mode='r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            rname = row['room_name'].strip()
            rtype = row['room_type'].strip().lower()
            if rtype == "theory":
                theory_rooms.append(rname)
            elif rtype == "lab":
                lab_rooms.append(rname)
    return {
        "theory_rooms": theory_rooms,
        "lab_rooms": lab_rooms
    }

def load_usage(json_path: str) -> dict:
    """
    Returns usage_data with structure:
    {
      "theory": { room_name: { day: [timeslots], ... }, ...},
      "lab": { room_name: { day: [labslots], ... }, ... }
    }
    """
    if not os.path.exists(json_path):
        return {"theory": {}, "lab": {}}

    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return data

def save_usage(json_path: str, usage_data: dict):
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(usage_data, f, indent=2)

def parse_uploaded_semester_file(file) -> dict:
    """
    Expects a CSV with columns: semester,course_code,course_name,is_lab,times_needed
    Returns a dict: { semester -> [(code, cname, bool, times_needed), ...] }
    """
    import io
    import csv

    file_data = file.getvalue().decode('utf-8')
    reader = csv.DictReader(io.StringIO(file_data))

    semester_courses = {}
    for row in reader:
        sem = int(row['semester'])
        code = row['course_code'].strip()
        cname = row['course_name'].strip()
        is_lab = (row['is_lab'].lower() == 'true')
        times_needed = int(row['times_needed'])

        if sem not in semester_courses:
            semester_courses[sem] = []
        semester_courses[sem].append((code, cname, is_lab, times_needed))

    return semester_courses

def reset_usage(json_path: str):
    """
    Clears all usage (makes all rooms/timeslots free).
    """
    empty_data = {"theory": {}, "lab": {}}
    save_usage(json_path, empty_data)
