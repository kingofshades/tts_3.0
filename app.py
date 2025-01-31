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
from scheduling.electives_solver import schedule_electives
from scheduling.utils import (
    build_full_room_usage_df,
    build_section_dataframe,
    build_room_usage_df,
    export_timetables_to_excel
)

def main():
    st.title("UMT Timetable Scheduler")

    # --------------------------------
    # BASIC CONFIG
    # --------------------------------
    DAYS = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"]
    THEORY_TIMESLOTS = [0,1,2,3,4,5,6]
    LAB_SLOTS = [0,1,2,3]

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

    LAB_OVERLAP_MAP = {
        0: [0, 1],
        1: [2, 3],
        2: [4, 5],
        3: [6]
    }

    # --------------------------------
    # SIDEBAR: RESET & NEW TIMETABLE
    # --------------------------------
    st.sidebar.subheader("Actions")
    if st.sidebar.button("Reset Usage Data"):
        reset_usage(os.path.join("data", "usage_data.json"))
        st.write("""<script>location.reload()</script>""", unsafe_allow_html=True)
        st.rerun(scope="app")

    if st.sidebar.button("Schedule a New Timetable"):
        st.write("""<script>location.reload()</script>""", unsafe_allow_html=True)
        st.rerun(scope="app")

    # --------------------------------
    # UPLOAD EXCEL
    # --------------------------------
    st.sidebar.subheader("Upload Excel")
    excel_file = st.sidebar.file_uploader("Excel with Roadmap, Rooms, StudentCapacity, Electives", type=["xlsx"])
    if not excel_file:
        st.info("Please upload an Excel file to continue.")
        return

    with st.spinner("Reading Excel..."):
        (semester_courses_map,
         excel_theory_rooms,
         excel_lab_rooms,
         student_capacities,
         electives_list) = parse_single_excel(excel_file)
    st.success("Excel loaded successfully!")

    # Load usage data
    usage_data = load_usage(os.path.join("data", "usage_data.json"))

    # --------------------------------
    # INIT SESSION STATE (ROOMS)
    # --------------------------------
    if "theory_rooms_current" not in st.session_state:
        usage_theory_rooms = set(usage_data["theory"].keys())
        st.session_state["theory_rooms_current"] = list(set(excel_theory_rooms) | usage_theory_rooms)

    if "lab_rooms_current" not in st.session_state:
        usage_lab_rooms = set(usage_data["lab"].keys())
        st.session_state["lab_rooms_current"] = list(set(excel_lab_rooms) | usage_lab_rooms)

    if "special_lab_rooms" not in st.session_state:
        st.session_state["special_lab_rooms"] = {
            "NS125L": ["PhysicsLab1", "PhysicsLab2"],
            "CC121L": ["DLDLab1", "DLDLab2"]
        }

    current_theory_rooms = st.session_state["theory_rooms_current"]
    current_lab_rooms = st.session_state["lab_rooms_current"]
    special_lab_rooms = st.session_state["special_lab_rooms"]

    # --------------------------------
    # REMOVE ROOMS
    # --------------------------------
    st.sidebar.subheader("Remove Existing Rooms")
    theory_remove = st.sidebar.multiselect("Remove Theory Rooms", current_theory_rooms)
    lab_remove = st.sidebar.multiselect("Remove Lab Rooms", current_lab_rooms)

    if st.sidebar.button("Remove Selected Rooms"):
        for r in theory_remove:
            if r in current_theory_rooms:
                current_theory_rooms.remove(r)
                st.sidebar.success(f"Theory room '{r}' removed for this run.")
        for r in lab_remove:
            if r in current_lab_rooms:
                current_lab_rooms.remove(r)
                st.sidebar.success(f"Lab room '{r}' removed for this run.")

    # --------------------------------
    # ADD NEW ROOM
    # --------------------------------
    st.sidebar.subheader("Add a New Room")
    new_room_name = st.sidebar.text_input("Room Name", "")
    new_room_type = st.sidebar.selectbox("Room Type", ["theory","lab"])
    if st.sidebar.button("Add Room"):
        if new_room_name.strip():
            if new_room_type == "theory":
                if new_room_name not in current_theory_rooms:
                    current_theory_rooms.append(new_room_name)
                    st.sidebar.success(f"Theory room '{new_room_name}' added!")
                    if new_room_name not in usage_data["theory"]:
                        usage_data["theory"][new_room_name] = {}
                        save_usage(os.path.join("data", "usage_data.json"), usage_data)
                else:
                    st.sidebar.error(f"Theory room '{new_room_name}' already exists.")
            else:
                if new_room_name not in current_lab_rooms:
                    current_lab_rooms.append(new_room_name)
                    st.sidebar.success(f"Lab room '{new_room_name}' added!")
                    if new_room_name not in usage_data["lab"]:
                        usage_data["lab"][new_room_name] = {}
                        save_usage(os.path.join("data", "usage_data.json"), usage_data)
                else:
                    st.sidebar.error(f"Lab room '{new_room_name}' already exists.")
        else:
            st.sidebar.error("Please enter a valid room name.")

    # --------------------------------
    # MANAGE SPECIAL LABS
    # --------------------------------
    st.sidebar.subheader("Manage Special Labs")
    for code, labs in special_lab_rooms.items():
        st.sidebar.write(f"**{code}** => {labs}")

    spec_rm_key = st.sidebar.text_input("Remove SpecialLab code", "")
    if st.sidebar.button("Remove SpecialLab Key"):
        if spec_rm_key in special_lab_rooms:
            del special_lab_rooms[spec_rm_key]
            st.sidebar.success(f"Removed special labs for '{spec_rm_key}'")

    new_spec_key = st.sidebar.text_input("New SpecialLab code", "")
    new_spec_val = st.sidebar.text_input("New SpecialLab rooms (comma separated)", "")
    if st.sidebar.button("Add/Update SpecialLab"):
        room_list = [x.strip() for x in new_spec_val.split(",") if x.strip()]
        if new_spec_key.strip() and room_list:
            special_lab_rooms[new_spec_key.strip()] = room_list
            st.sidebar.success(f"Set special labs for {new_spec_key} = {room_list}")

    # --------------------------------
    # FREE SLOTS HELPER
    # --------------------------------
    def get_free_slot_count(rtype: str, room_name: str) -> int:
        if rtype=="theory":
            total_slots = len(DAYS)*len(THEORY_TIMESLOTS)
            used=0
            if room_name in usage_data["theory"]:
                for dd in usage_data["theory"][room_name]:
                    used += len(usage_data["theory"][room_name][dd])
            return total_slots - used
        else:
            total_slots = len(DAYS)*len(LAB_SLOTS)
            used=0
            if room_name in usage_data["lab"]:
                for dd in usage_data["lab"][room_name]:
                    used += len(usage_data["lab"][room_name][dd])
            return total_slots - used

    # Filter
    filtered_theory_rooms = [r for r in current_theory_rooms if get_free_slot_count("theory", r)>0]
    filtered_lab_rooms = [r for r in current_lab_rooms if get_free_slot_count("lab", r)>0]

    st.write("### Current Rooms (with > 0 free slots)")
    col1, col2 = st.columns(2)
    with col1:
        st.write("**Theory Rooms**")
        st.write(filtered_theory_rooms)
    with col2:
        st.write("**Lab Rooms**")
        st.write(filtered_lab_rooms)

    # --------------------------------
    # SELECT SEMESTERS
    # --------------------------------
    st.sidebar.subheader("Select Semesters")
    all_semesters = sorted(semester_courses_map.keys())
    selected_semesters = st.sidebar.multiselect("Semesters", all_semesters, default=all_semesters)
    if not selected_semesters:
        st.warning("No semester selected.")
        return

    # Student capacities
    st.sidebar.subheader("Student Capacities")
    final_capacities={}
    for sem in selected_semesters:
        def_val = student_capacities.get(sem, 50)
        val = st.sidebar.number_input(f"Semester {sem} Students", min_value=1, value=def_val)
        final_capacities[sem] = val

    # Program code
    st.sidebar.subheader("Program Code")
    program_code = st.sidebar.text_input("Enter Program Code", value="A")

    # Output filenames
    st.sidebar.subheader("Output Filenames")
    out_file = st.sidebar.text_input("Main Timetable Filename", "timetables.xlsx")
    remcap_file = st.sidebar.text_input("Remaining Capacity Workbook", "remaining_capacity.xlsx")
    elec_out_file = st.sidebar.text_input("Electives Timetable Filename", "electives_timetable.xlsx")

    # --------------------------------
    # PRE-CALC FOR MAIN
    # --------------------------------
    total_needed_theory_slots=0
    total_needed_lab_slots=0
    for sem in selected_semesters:
        num_students = final_capacities[sem]
        sec_count = math.ceil(num_students/50)
        for (code, cname, is_lab, times_needed) in semester_courses_map[sem]:
            if is_lab:
                total_needed_lab_slots += times_needed*sec_count
            else:
                total_needed_theory_slots += times_needed*sec_count

    free_theory_cap = sum(get_free_slot_count("theory", r) for r in filtered_theory_rooms)
    free_lab_cap = sum(get_free_slot_count("lab", r) for r in filtered_lab_rooms)

    st.write("### Required vs. Available Slots (Main Courses)")
    st.write(f"**Theory** needed={total_needed_theory_slots}, available={free_theory_cap}")
    st.write(f"**Lab** needed={total_needed_lab_slots}, available={free_lab_cap}")
    can_generate_main=True
    if total_needed_theory_slots>free_theory_cap or total_needed_lab_slots>free_lab_cap:
        can_generate_main=False

    # --------------------------------
    # BUTTON: Generate Main Timetable
    # --------------------------------
    if st.button("Generate Timetable (Main)", disabled=(not can_generate_main)):
        if not can_generate_main:
            st.error("Not enough slots for main courses. Add more rooms or reduce load.")
            st.stop()

        with st.spinner("Scheduling main timetable..."):
            usage_now = load_usage(os.path.join("data","usage_data.json"))
            try:
                result = schedule_timetable(
                    selected_semesters=selected_semesters,
                    semester_courses_map=semester_courses_map,
                    section_sizes=final_capacities,
                    usage_data=usage_now,
                    DAYS=DAYS,
                    THEORY_TIMESLOTS=THEORY_TIMESLOTS,
                    TIMESLOT_LABELS=TIMESLOT_LABELS,
                    LAB_SLOTS=LAB_SLOTS,
                    LAB_SLOT_LABELS=LAB_SLOT_LABELS,
                    LAB_OVERLAP_MAP=LAB_OVERLAP_MAP,
                    theory_rooms=filtered_theory_rooms,
                    lab_rooms=filtered_lab_rooms,
                    special_lab_rooms=special_lab_rooms,
                    section_size=50,
                    program_code=program_code
                )
            except ValueError as e:
                st.error(f"Scheduling error: {e}")
                st.stop()

        if not result:
            st.error("No feasible solution found for main timetable.")
            st.stop()

        schedule_map, sem_sections_map, new_allocs = result
        st.success("Main Timetable generated successfully!")

        # Update usage
        usage_data_current = load_usage(os.path.join("data","usage_data.json"))
        for (rtype, rname, d, slot, occupant) in new_allocs:
            usage_data_current.setdefault(rtype, {})
            usage_data_current[rtype].setdefault(rname, {})
            usage_data_current[rtype][rname].setdefault(d, [])
            usage_data_current[rtype][rname][d].append(slot)
        save_usage(os.path.join("data","usage_data.json"), usage_data_current)

        # Export
        export_timetables_to_excel(
            schedule_map=schedule_map,
            semester_sections_map=sem_sections_map,
            semester_courses_map=semester_courses_map,
            filename=out_file,
            DAYS=DAYS,
            THEORY_TIMESLOTS=THEORY_TIMESLOTS,
            TIMESLOT_LABELS=TIMESLOT_LABELS,
            LAB_SLOTS=LAB_SLOTS,
            LAB_SLOT_LABELS=LAB_SLOT_LABELS,
            theory_rooms=filtered_theory_rooms,
            lab_rooms=filtered_lab_rooms,
            special_lab_rooms=special_lab_rooms
        )
        st.success(f"Main timetables exported to '{out_file}'")

        # Show Timetables by Section
        st.header("Generated Timetables by Section (Main)")
        for sem, sec_list in sem_sections_map.items():
            st.subheader(f"Semester {sem}")
            courses = semester_courses_map[sem]
            for sec_name in sec_list:
                st.write(f"**Section {sec_name}**")
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
                    special_lab_rooms=special_lab_rooms
                )
                st.table(df_sec)

        # Summaries
        st.header("Room Usage & Remaining Capacity (Main)")
        usage_after_main = load_usage(os.path.join("data","usage_data.json"))
        summary_rows=[]
        for r in filtered_theory_rooms:
            used_count=0
            if r in usage_after_main["theory"]:
                for d in usage_after_main["theory"][r]:
                    used_count+=len(usage_after_main["theory"][r][d])
            total_slots = len(DAYS)*len(THEORY_TIMESLOTS)
            free_count = total_slots - used_count
            summary_rows.append([r,"Theory", used_count, free_count, total_slots])

        for r in filtered_lab_rooms:
            used_count=0
            if r in usage_after_main["lab"]:
                for d in usage_after_main["lab"][r]:
                    used_count+=len(usage_after_main["lab"][r][d])
            total_slots = len(DAYS)*len(LAB_SLOTS)
            free_count = total_slots - used_count
            summary_rows.append([r,"Lab", used_count, free_count, total_slots])

        df_main_sum = pd.DataFrame(summary_rows, columns=["Room","Type","Used Slots","Free Slots","Total Slots"])
        st.write("### Summary of Room Usage (Main)")
        st.table(df_main_sum)

        writer_usage = pd.ExcelWriter(remcap_file, engine="openpyxl")
        df_main_sum.to_excel(writer_usage, sheet_name="Summary", index=False)
        # Detailed usage
        for rr in filtered_theory_rooms:
            df_tr = build_full_room_usage_df(
                room=rr,
                rtype="theory",
                usage_data=usage_after_main,
                schedule_map=schedule_map,
                DAYS=DAYS,
                THEORY_TIMESLOTS=THEORY_TIMESLOTS,
                LAB_SLOTS=LAB_SLOTS,
                TIMESLOT_LABELS=TIMESLOT_LABELS,
                LAB_SLOT_LABELS=LAB_SLOT_LABELS
            )
            df_tr.to_excel(writer_usage, sheet_name=(rr[:20]+"_Usage"), index=True)

        for lb in filtered_lab_rooms:
            df_lb = build_full_room_usage_df(
                room=lb,
                rtype="lab",
                usage_data=usage_after_main,
                schedule_map=schedule_map,
                DAYS=DAYS,
                THEORY_TIMESLOTS=THEORY_TIMESLOTS,
                LAB_SLOTS=LAB_SLOTS,
                TIMESLOT_LABELS=TIMESLOT_LABELS,
                LAB_SLOT_LABELS=LAB_SLOT_LABELS
            )
            df_lb.to_excel(writer_usage, sheet_name=(lb[:20]+"_Usage"), index=True)

        writer_usage.close()
        st.success(f"Room usage details + capacity saved to '{remcap_file}'")

    # --------------------------------
    # ELECTIVES
    # --------------------------------
    leftover_theory_cap = sum(get_free_slot_count("theory", r) for r in filtered_theory_rooms)
    leftover_lab_cap = sum(get_free_slot_count("lab", r) for r in filtered_lab_rooms)

    if electives_list:
        st.header("Schedule Electives")
        # rough worst-case
        max_theory_needed=0
        max_lab_needed=0
        for elec in electives_list:
            if elec["can_theory"]:
                max_theory_needed += 2*elec["sections_count"]
            if elec["can_lab"]:
                max_lab_needed += 1*elec["sections_count"]

        st.write("**Potential Electives Requirements**:")
        st.write(f"- Up to {max_theory_needed} theory slots (worst case).")
        st.write(f"- Up to {max_lab_needed} lab slots (worst case).")
        st.write(f"- Leftover theory capacity: {leftover_theory_cap}, leftover lab capacity: {leftover_lab_cap}")

        if st.button("Generate Electives Timetable"):
            with st.spinner("Scheduling electives..."):
                usage_latest = load_usage(os.path.join("data","usage_data.json"))
                result_elec = schedule_electives(
                    electives_list=electives_list,
                    usage_data=usage_latest,
                    DAYS=DAYS,
                    THEORY_TIMESLOTS=THEORY_TIMESLOTS,
                    LAB_SLOTS=LAB_SLOTS,
                    theory_rooms=filtered_theory_rooms,
                    lab_rooms=filtered_lab_rooms,
                    timeslot_labels=TIMESLOT_LABELS,
                    lab_slot_labels=LAB_SLOT_LABELS,
                    theory_needed=2,  # 2 distinct timeslots for theory
                    lab_needed=1      # 1 slot for lab
                )

            if result_elec is None:
                st.error("No feasible solution found for electives. Add more rooms or remove electives.")
            else:
                schedule_map_elec, new_allocs_elec = result_elec
                st.success("Electives scheduled successfully!")

                # Update usage
                usage_data_latest2 = load_usage(os.path.join("data","usage_data.json"))
                for (rtype, rname, d, slot, occupant) in new_allocs_elec:
                    usage_data_latest2.setdefault(rtype, {})
                    usage_data_latest2[rtype].setdefault(rname, {})
                    usage_data_latest2[rtype][rname].setdefault(d, [])
                    usage_data_latest2[rtype][rname][d].append(slot)
                save_usage(os.path.join("data","usage_data.json"), usage_data_latest2)

                # Build DataFrame
                rows=[]
                for (e_code, idx), asgn in schedule_map_elec.items():
                    for (rtype, room, d, slot) in asgn:
                        sec_label = f"A{idx+1}"
                        if rtype=="theory":
                            slot_label = TIMESLOT_LABELS.get(slot, str(slot))
                        else:
                            slot_label = LAB_SLOT_LABELS.get(slot, str(slot))
                        rows.append([e_code, sec_label, rtype.capitalize(), room, d, slot_label])
                df_elec = pd.DataFrame(rows, columns=["ElectiveCode","Section","RoomType","Room","Day","Slot"])
                st.write("### Electives Timetable")
                st.table(df_elec)

                # Save to excel
                with pd.ExcelWriter(elec_out_file, engine="openpyxl") as writer:
                    df_elec.to_excel(writer, sheet_name="ElectivesSchedule", index=False)
                st.success(f"Electives timetable saved to '{elec_out_file}'.")


if __name__=="__main__":
    main()
