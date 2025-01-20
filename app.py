import streamlit as st
import os
import math
import pandas as pd
from openpyxl import Workbook

# Local imports
from data.data_io import (
    load_usage,
    save_usage,
    reset_usage,
    parse_single_excel
)
from scheduling.solver import schedule_timetable
from scheduling.utils import (
    build_full_room_usage_df,
    build_section_dataframe,
    build_room_usage_df,
    export_timetables_to_excel
)

def main():
    st.title("UMT Timetable Scheduler (Merging Usage Data + Excel Rooms)")

    # Basic config for your schedule
    DAYS = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"]
    THEORY_TIMESLOTS = [0,1,2,3,4,5,6]  # 7 theory slots each day => 6 days => 42 total per theory room
    LAB_SLOTS = [0,1,2,3]             # 4 lab slots each day => 6 days => 24 total per lab room

    TIMESLOT_LABELS = {
        0: "8:00-9:15",
        1: "9:30-10:45",
        2: "11:00-12:15",
        3: "12:30-1:45",
        4: "2:00-3:15",
        5: "3:30-4:45",
        6: "5:00-6:15"
    }

    LAB_SLOT_LABELS = {
        0: "8:00-10:30",
        1: "11:00-1:30",
        2: "2:00-4:30",
        3: "5:00-7:30"
    }

    # Overlap map for partial clash constraints (lab slot -> theory slots overlapping)
    LAB_OVERLAP_MAP = {
        0: [0, 1],
        1: [2, 3],
        2: [4, 5],
        3: [6]
    }

    # Helper function: how many free slots remain for a given room based on usage_data
    def get_free_slot_count(rtype: str, room_name: str) -> int:
        """Return how many free timeslots remain for this room_name of rtype (theory|lab) 
           given the usage_data. Theory rooms have 42 total, labs have 24 total."""
        if rtype == "theory":
            total_slots = len(DAYS) * len(THEORY_TIMESLOTS)  # 42
            used = 0
            if room_name in usage_data["theory"]:
                for day in usage_data["theory"][room_name]:
                    used += len(usage_data["theory"][room_name][day])
            return total_slots - used
        else:
            total_slots = len(DAYS) * len(LAB_SLOTS)  # 24
            used = 0
            if room_name in usage_data["lab"]:
                for day in usage_data["lab"][room_name]:
                    used += len(usage_data["lab"][room_name][day])
            return total_slots - used

    # 1) Sidebar: Reset usage
    st.sidebar.subheader("Usage Data")
    if st.sidebar.button("Reset Usage Data"):
        reset_usage(os.path.join("data", "usage_data.json"))
        st.warning("Usage data has been reset. Please reload or re-run.")
        return

    # 2) Upload Excel
    st.sidebar.subheader("Upload Excel")
    excel_file = st.sidebar.file_uploader("Excel with Roadmap, Rooms, StudentCapacity", type=["xlsx"])
    if not excel_file:
        st.info("Please upload an Excel file to continue.")
        return

    # Parse the Excel data
    with st.spinner("Reading Excel..."):
        semester_courses_map, excel_theory_rooms, excel_lab_rooms, student_capacities = parse_single_excel(excel_file)
    st.success("Excel loaded successfully!")

    # 3) Load usage data
    usage_data = load_usage(os.path.join("data", "usage_data.json"))

    # 3a) Merge rooms from usage_data + Excel
    #    (the usage_data might have rooms not in Excel, or vice versa)
    usage_theory_rooms = set(usage_data["theory"].keys())  # from JSON
    usage_lab_rooms = set(usage_data["lab"].keys())

    # unify sets
    all_theory_rooms = set(excel_theory_rooms) | usage_theory_rooms
    all_lab_rooms = set(excel_lab_rooms) | usage_lab_rooms

    # Now convert them back to lists
    all_theory_rooms = list(all_theory_rooms)
    all_lab_rooms = list(all_lab_rooms)

    # 4) Remove existing rooms if not needed
    st.sidebar.subheader("Remove Existing Rooms")
    theory_remove = st.sidebar.multiselect("Remove Theory Rooms", all_theory_rooms)
    lab_remove = st.sidebar.multiselect("Remove Lab Rooms", all_lab_rooms)
    if st.sidebar.button("Remove Selected Rooms"):
        for r in theory_remove:
            if r in all_theory_rooms:
                all_theory_rooms.remove(r)
                st.sidebar.success(f"Theory room '{r}' removed.")
                # Also remove from usage_data if present
                if r in usage_data["theory"]:
                    usage_data["theory"].pop(r, None)
        for r in lab_remove:
            if r in all_lab_rooms:
                all_lab_rooms.remove(r)
                st.sidebar.success(f"Lab room '{r}' removed.")
                if r in usage_data["lab"]:
                    usage_data["lab"].pop(r, None)
        # Save usage_data after removal
        save_usage(os.path.join("data", "usage_data.json"), usage_data)

    # 5) Add new rooms dynamically
    st.sidebar.subheader("Add a New Room")
    new_room_name = st.sidebar.text_input("Room Name", "")
    new_room_type = st.sidebar.selectbox("Room Type", ["theory","lab"])
    if st.sidebar.button("Add Room"):
        if new_room_name.strip():
            if new_room_type == "theory":
                if new_room_name not in all_theory_rooms:
                    all_theory_rooms.append(new_room_name)
                    st.sidebar.success(f"Theory room '{new_room_name}' added!")
                else:
                    st.sidebar.error(f"Theory room '{new_room_name}' already exists.")
            else:
                if new_room_name not in all_lab_rooms:
                    all_lab_rooms.append(new_room_name)
                    st.sidebar.success(f"Lab room '{new_room_name}' added!")
                else:
                    st.sidebar.error(f"Lab room '{new_room_name}' already exists.")
        else:
            st.sidebar.error("Please enter a valid room name.")

    # 6) Filter out rooms with 0 free slots
    #    so we only schedule with rooms that still have availability
    filtered_theory_rooms = []
    for r in all_theory_rooms:
        free = get_free_slot_count("theory", r)
        if free > 0:
            filtered_theory_rooms.append(r)

    filtered_lab_rooms = []
    for r in all_lab_rooms:
        free = get_free_slot_count("lab", r)
        if free > 0:
            filtered_lab_rooms.append(r)

    # 7) Show the "current rooms" that actually have free slots
    st.write("### Current Rooms (with > 0 free slots)")
    col1, col2 = st.columns(2)
    with col1:
        st.write("**Theory Rooms**")
        st.write(filtered_theory_rooms)
    with col2:
        st.write("**Lab Rooms**")
        st.write(filtered_lab_rooms)

    # 8) Select Semesters
    st.sidebar.subheader("Select Semesters")
    all_semesters = sorted(semester_courses_map.keys())
    selected_semesters = st.sidebar.multiselect("Semesters", all_semesters, default=all_semesters)
    if not selected_semesters:
        st.warning("No semester selected.")
        return

    # 9) Student capacity overrides
    st.sidebar.subheader("Student Capacities")
    final_capacities = {}
    for sem in selected_semesters:
        default_val = student_capacities.get(sem, 50)
        new_val = st.sidebar.number_input(f"Semester {sem} Students", min_value=1, value=default_val)
        final_capacities[sem] = new_val

    # 10) Program code
    st.sidebar.subheader("Program Code")
    program_code = st.sidebar.text_input("Enter Program Code", value="A")

    # 11) Output filenames
    st.sidebar.subheader("Output Filenames")
    out_file = st.sidebar.text_input("Excel Timetable Filename", "timetables.xlsx")
    remcap_file = st.sidebar.text_input("Remaining Capacity Workbook", "remaining_capacity.xlsx")

    # 12) Generate Timetable
    if st.sidebar.button("Generate Timetable"):
        # (A) Calculate how many total theory/lab slots are needed
        total_needed_theory_slots = 0
        total_needed_lab_slots = 0

        for sem in selected_semesters:
            num_students = final_capacities[sem]
            section_count = math.ceil(num_students / 50)  # each section up to 50 students
            for (course_code, cname, is_lab, times_needed) in semester_courses_map[sem]:
                if is_lab:
                    total_needed_lab_slots += (times_needed * section_count)
                else:
                    total_needed_theory_slots += (times_needed * section_count)

        # (B) Calculate how many free slots we have in the filtered rooms
        #     Because those are the only rooms we are scheduling with.
        #     Sum the free slots across all filtered theory rooms
        free_theory_cap = 0
        for r in filtered_theory_rooms:
            free_theory_cap += get_free_slot_count("theory", r)

        free_lab_cap = 0
        for r in filtered_lab_rooms:
            free_lab_cap += get_free_slot_count("lab", r)

        # Compare
        if total_needed_theory_slots > free_theory_cap:
            st.error(
                f"Not enough free THEORY slots.\n"
                f"Needed = {total_needed_theory_slots}, Have = {free_theory_cap} "
                f"(with {len(filtered_theory_rooms)} usable theory rooms)."
            )
            return
        elif total_needed_theory_slots == free_theory_cap:
            st.info(f"Exactly enough free THEORY slots: needed={total_needed_theory_slots}, available={free_theory_cap}")
        else:
            st.info(f"Enough free THEORY slots: needed={total_needed_theory_slots}, available={free_theory_cap}")

        if total_needed_lab_slots > free_lab_cap:
            st.error(
                f"Not enough free LAB slots.\n"
                f"Needed = {total_needed_lab_slots}, Have = {free_lab_cap} "
                f"(with {len(filtered_lab_rooms)} usable lab rooms)."
            )
            return
        elif total_needed_lab_slots == free_lab_cap:
            st.info(f"Exactly enough free LAB slots: needed={total_needed_lab_slots}, available={free_lab_cap}")
        else:
            st.info(f"Enough free LAB slots: needed={total_needed_lab_slots}, available={free_lab_cap}")

        # (C) If feasible, proceed to the solver
        with st.spinner("Scheduling..."):
            from scheduling.solver import schedule_timetable
            try:
                result = schedule_timetable(
                    selected_semesters=selected_semesters,
                    semester_courses_map=semester_courses_map,
                    section_sizes=final_capacities,
                    usage_data=load_usage(os.path.join("data", "usage_data.json")),  # fresh load
                    DAYS=DAYS,
                    THEORY_TIMESLOTS=THEORY_TIMESLOTS,
                    TIMESLOT_LABELS=TIMESLOT_LABELS,
                    LAB_SLOTS=LAB_SLOTS,
                    LAB_SLOT_LABELS=LAB_SLOT_LABELS,
                    LAB_OVERLAP_MAP=LAB_OVERLAP_MAP,
                    theory_rooms=filtered_theory_rooms,   # Only pass rooms that have free slots
                    lab_rooms=filtered_lab_rooms,
                    special_lab_rooms={
                        # Hard-coded example
                        "NS125L": ["PhysicsLab1", "PhysicsLab2"],
                        "CC121L": ["DLDLab1", "DLDLab2"]
                    },
                    section_size=50,
                    program_code=program_code
                )
            except ValueError as e:
                st.error(f"Scheduling Error: {e}")
                return

        if not result:
            st.error("No feasible solution found. Try adding more rooms or adjusting constraints.")
            return

        schedule_map, semester_sections_map, new_allocs = result
        st.success("Timetable generated successfully!")

        # (D) Update usage data
        usage_data_current = load_usage(os.path.join("data", "usage_data.json"))
        for (rtype, rname, day, slot, occupant) in new_allocs:
            usage_data_current.setdefault(rtype, {})
            usage_data_current[rtype].setdefault(rname, {})
            usage_data_current[rtype][rname].setdefault(day, [])
            usage_data_current[rtype][rname][day].append(slot)

        save_usage(os.path.join("data", "usage_data.json"), usage_data_current)
        st.info("Usage data updated in JSON.")

        # (E) Export timetables to Excel
        export_timetables_to_excel(
            schedule_map=schedule_map,
            semester_sections_map=semester_sections_map,
            semester_courses_map=semester_courses_map,
            filename=out_file,
            DAYS=DAYS,
            THEORY_TIMESLOTS=THEORY_TIMESLOTS,
            TIMESLOT_LABELS=TIMESLOT_LABELS,
            LAB_SLOTS=LAB_SLOTS,
            LAB_SLOT_LABELS=LAB_SLOT_LABELS,
            theory_rooms=filtered_theory_rooms,
            lab_rooms=filtered_lab_rooms,
            special_lab_rooms={
                "NS125L": ["PhysicsLab1", "PhysicsLab2"],
                "CC121L": ["DLDLab1", "DLDLab2"]
            }
        )
        st.success(f"Timetables exported to '{out_file}'")

        # (F) Display Timetables by Section
        st.header("Generated Timetables by Section")
        for sem, sec_list in semester_sections_map.items():
            st.subheader(f"Semester {sem}")
            courses = semester_courses_map[sem]
            for sec_name in sec_list:
                st.write(f"**Section {sec_name}**")
                from scheduling.utils import build_section_dataframe
                df_sec = build_section_dataframe(
                    section_name=sec_name,
                    courses=courses,
                    schedule_map=schedule_map,
                    DAYS=DAYS,
                    THEORY_TIMESLOTS=THEORY_TIMESLOTS,
                    TIMESLOT_LABELS=TIMESLOT_LABELS,
                    LAB_SLOTS=LAB_SLOTS,
                    LAB_SLOT_LABELS=LAB_SLOT_LABELS,
                    theory_rooms=filtered_theory_rooms,
                    lab_rooms=filtered_lab_rooms,
                    special_lab_rooms={
                        "NS125L": ["PhysicsLab1", "PhysicsLab2"],
                        "CC121L": ["DLDLab1", "DLDLab2"]
                    }
                )
                st.table(df_sec)

        # (G) Room usage & remaining capacity
        st.header("Room Usage & Remaining Capacity")
        # Summaries
        summary_rows = []
        # re-load final usage
        final_usage = load_usage(os.path.join("data", "usage_data.json"))

        # For theory rooms (filtered or not?)
        # We'll focus on those we used in scheduling => filtered_theory_rooms
        for r in (filtered_theory_rooms):
            used_count = 0
            if r in final_usage["theory"]:
                for day in final_usage["theory"][r]:
                    used_count += len(final_usage["theory"][r][day])
            free_count = (len(DAYS) * len(THEORY_TIMESLOTS)) - used_count
            summary_rows.append([r, "Theory", used_count, free_count, (len(DAYS) * len(THEORY_TIMESLOTS))])

        # For lab
        # You might also unify special labs here if you want to display them too
        for r in (filtered_lab_rooms):
            used_count = 0
            if r in final_usage["lab"]:
                for day in final_usage["lab"][r]:
                    used_count += len(final_usage["lab"][r][day])
            free_count = (len(DAYS) * len(LAB_SLOTS)) - used_count
            summary_rows.append([r, "Lab", used_count, free_count, (len(DAYS) * len(LAB_SLOTS))])

        df_summary = pd.DataFrame(summary_rows, columns=["Room","Type","Used Slots","Free Slots","Total Slots"])
        st.write("### Summary of Room Usage")
        st.table(df_summary)

        # Save usage + usage tables to Excel
        writer_usage = pd.ExcelWriter(remcap_file, engine="openpyxl")
        df_summary.to_excel(writer_usage, sheet_name="Summary", index=False)

        st.subheader("Detailed Room Usage")
        from scheduling.utils import build_room_usage_df
        # Theory usage
        for rr in filtered_theory_rooms:
            st.write(f"**Room: {rr} (Theory)**")
            df_tr = build_full_room_usage_df(
                room=rr,
                rtype="theory",
                usage_data=final_usage,  # or load_usage(...) to be sure
                schedule_map=schedule_map,
                DAYS=DAYS,
                THEORY_TIMESLOTS=THEORY_TIMESLOTS,
                LAB_SLOTS=LAB_SLOTS,
                TIMESLOT_LABELS=TIMESLOT_LABELS,
                LAB_SLOT_LABELS=LAB_SLOT_LABELS
            )
            st.table(df_tr)
            df_tr.to_excel(writer_usage, sheet_name=f"{rr[:20]}_Usage", index=True)

        # Lab usage
        for lb in filtered_lab_rooms:
            st.write(f"**Lab: {lb}**")
            df_lb = build_full_room_usage_df(
                room=lb,
                rtype="lab",
                usage_data=final_usage,
                schedule_map=schedule_map,
                DAYS=DAYS,
                THEORY_TIMESLOTS=THEORY_TIMESLOTS,
                LAB_SLOTS=LAB_SLOTS,
                TIMESLOT_LABELS=TIMESLOT_LABELS,
                LAB_SLOT_LABELS=LAB_SLOT_LABELS
            )

            st.table(df_lb)
            df_lb.to_excel(writer_usage, sheet_name=f"{lb[:20]}_Usage", index=True)

        writer_usage.close()
        st.success(f"Room usage details and remaining capacity saved to '{remcap_file}'")

        # (H) Note for multiple programs
        st.markdown("""
        **Note**: To schedule for multiple programs while using leftover slots, 
        do **not** reset the usage data between runs.  
        1) Upload the new Excel (or the same if it covers multiple programs).  
        2) Choose your new semesters or the same ones with a new Program Code.  
        3) Generate Timetable. The solver will see leftover free slots from the previous schedule.
        """)

if __name__ == "__main__":
    main()
