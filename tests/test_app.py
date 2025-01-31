# tests/test_app.py
import pytest
import os
from data.data_io import load_usage, reset_usage, save_usage
from scheduling.solver import schedule_timetable, schedule_electives

@pytest.fixture
def fresh_usage(tmp_path):
    # create a temp usage file
    path = os.path.join(tmp_path, "usage_data.json")
    reset_usage(path)
    return path

def test_schedule_timetable_basic(fresh_usage):
    usage_path = fresh_usage
    usage_data = load_usage(usage_path)

    # Minimal data
    selected_semesters = [1]
    semester_courses_map = {
        1: [
            ("CS101", "IntroCS", False, 1),  # non-lab, times_needed=1 => 1*2=2 theory slots needed
        ]
    }
    section_sizes = {1: 50}  # => 1 section
    DAYS = ["Mon","Tue"]
    THEORY_TIMESLOTS = [0,1]
    LAB_SLOTS = [0]
    theory_rooms = ["T1"]
    lab_rooms = ["L1"]

    # We won't define special labs
    special_lab = {}

    result = schedule_timetable(
        selected_semesters,
        semester_courses_map,
        section_sizes,
        usage_data,
        DAYS,
        THEORY_TIMESLOTS,
        {},
        LAB_SLOTS,
        {},
        {},
        theory_rooms,
        lab_rooms,
        special_lab_rooms=special_lab
    )
    assert result is not None, "Should schedule fine in minimal scenario."

def test_schedule_electives_basic(fresh_usage):
    usage_path = fresh_usage
    usage_data = load_usage(usage_path)
    # Mark T1 -> Monday slot 0 as used
    usage_data["theory"]["T1"] = {"Monday":[0]}
    save_usage(usage_path, usage_data)

    # Minimal electives
    electives_list = [
        {
            "elective_code":"EL001",
            "elective_name":"TestElec",
            "elective_type":"General",
            "sections":2,
            "prefer_lab":False,
            "times_needed_theory":2,
            "times_needed_lab":1
        }
    ]
    DAYS=["Monday","Tuesday"]
    THEORY_TIMESLOTS=[0,1]
    LAB_SLOTS=[0,1]
    theory_rooms=["T1","T2"]
    lab_rooms=["L1"]

    usage_data_loaded = load_usage(usage_path)
    updated, allocs = schedule_electives(
        electives_list,
        usage_data_loaded,
        DAYS,
        THEORY_TIMESLOTS,
        LAB_SLOTS,
        theory_rooms,
        lab_rooms
    )
    assert len(allocs) > 0, "Should have allocated some slots for electives."
