import streamlit as st
import os

from data.data_io import (
    load_rooms,
    load_usage,
    save_usage,
    parse_uploaded_semester_file,
    reset_usage
)
from scheduling.solver import schedule_timetable
from scheduling.utils import (
    build_room_usage_df,
    export_timetables_to_excel,
    build_section_dataframe
)


def main():
    st.title("UMT Timetable Scheduler")

    # Basic config
    DAYS = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"]
    THEORY_TIMESLOTS = [0,1,2,3,4,5,6]
    TIMESLOT_LABELS = {
        0: "8:00-9:15",
        1: "9:30-10:45",
        2: "11:00-12:15",
        3: "12:30-1:45",
        4: "2:00-3:15",
        5: "3:30-4:45",
        6: "5:00-6:15"
    }
    LAB_SLOTS = [0,1,2,3]
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

    data_folder = "data"
    usage_path = os.path.join(data_folder, "usage_data.json")
    rooms_csv_path = os.path.join(data_folder, "rooms.csv")

    # 1) Load Rooms
    try:
        room_data = load_rooms(rooms_csv_path)
    except FileNotFoundError:
        st.error("rooms.csv not found in 'data' folder.")
        return

    theory_rooms = room_data["theory_rooms"]
    lab_rooms = room_data["lab_rooms"]

    # Update special labs to match what's in rooms.csv
    special_lab_rooms = {
        "NS125L": ["PhysicsLab1", "PhysicsLab2"],
        "CC121L": ["DLDLab1", "DLDLab2"]
    }

    # 2) Load usage data
    usage_data = load_usage(usage_path)

    # 3) Upload CSV
    st.sidebar.markdown("## Upload Program Road-Map File")
    file = st.sidebar.file_uploader(
        "Upload CSV (semester, course_code, course_name, is_lab, times_needed)",
        type=["csv"]
    )
    if not file:
        st.info("Please upload a CSV to proceed.")
        return

    semester_courses_map = parse_uploaded_semester_file(file)
    all_semesters = sorted(semester_courses_map.keys())
    if not all_semesters:
        st.warning("No valid rows found in the uploaded CSV.")
        return

    selected_semesters = st.sidebar.multiselect("Select Semesters", all_semesters)
    if not selected_semesters:
        st.warning("No semesters selected.")
        return

    # 4) Section sizes
    section_sizes = {}
    for sem in selected_semesters:
        n = st.sidebar.number_input(f"Number of Students (Semester {sem})", min_value=1, value=50)
        section_sizes[sem] = n

    # 5) Option to reset usage
    if st.sidebar.button("Reset Usage Data"):
        reset_usage(usage_path)
        # NOTE: st.experimental_rerun() might not exist in older versions of Streamlit
        # st.experimental_rerun()  # If you have a newer Streamlit; otherwise comment out.
        st.warning("Usage reset! Reload or re-run to see changes.")
        return

    # 6) Generate Timetable
    if st.sidebar.button("Generate Timetable"):
        with st.spinner("Scheduling..."):
            try:
                result = schedule_timetable(
                    selected_semesters=selected_semesters,
                    semester_courses_map=semester_courses_map,
                    section_sizes=section_sizes,
                    usage_data=usage_data,
                    DAYS=DAYS,
                    THEORY_TIMESLOTS=THEORY_TIMESLOTS,
                    TIMESLOT_LABELS=TIMESLOT_LABELS,
                    LAB_SLOTS=LAB_SLOTS,
                    LAB_SLOT_LABELS=LAB_SLOT_LABELS,
                    LAB_OVERLAP_MAP=LAB_OVERLAP_MAP,
                    theory_rooms=theory_rooms,
                    lab_rooms=lab_rooms,
                    special_lab_rooms=special_lab_rooms,
                    section_size=50
                )
            except ValueError as e:
                st.error(f"Scheduling Error: {e}")
                return

        if not result:
            st.error("No feasible solution found. Try adjusting constraints or adding rooms.")
        else:
            schedule_map, sem_sections_map, new_allocs = result
            st.success("Timetable generated successfully!")

            # 7) Update usage_data so future runs know rooms are used
            for (rtype, rname, day, slot, occupant) in new_allocs:
                usage_data[rtype].setdefault(rname, {})
                usage_data[rtype][rname].setdefault(day, [])
                usage_data[rtype][rname][day].append(slot)

            save_usage(usage_path, usage_data)
            st.info("Usage data updated for future scheduling runs.")

            # 8) Export to Excel
            out_file = "timetables.xlsx"
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
                theory_rooms=theory_rooms,
                lab_rooms=lab_rooms,
                special_lab_rooms=special_lab_rooms
            )
            st.success(f"Exported timetables to {out_file}")

            # 9) Display each section’s schedule
            st.header("Generated Timetables")
            for sem, sections in sem_sections_map.items():
                st.subheader(f"Semester {sem}")
                for sec in sections:
                    st.write(f"**Section {sec}**")
                    df_sec = build_section_dataframe(
                        section=sec,
                        courses=semester_courses_map[sem],
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
                    st.table(df_sec)

            # 10) Show resource usage
            st.header("Resource Usage (Final)")
            st.subheader("Theory Rooms")
            for rr in theory_rooms:
                st.write(f"**Room: {rr}**")
                df_tr = build_room_usage_df(
                    room=rr,
                    schedule_map=schedule_map,
                    is_lab=False,
                    DAYS=DAYS,
                    THEORY_TIMESLOTS=THEORY_TIMESLOTS,
                    TIMESLOT_LABELS=TIMESLOT_LABELS,
                    LAB_SLOTS=LAB_SLOTS,
                    LAB_SLOT_LABELS=LAB_SLOT_LABELS
                )
                st.table(df_tr)

            st.subheader("Lab Rooms")
            # Combine labs + special labs if you want to see them in one block
            combined_labs = list(lab_rooms)
            for slist in special_lab_rooms.values():
                for labn in slist:
                    if labn not in combined_labs:
                        combined_labs.append(labn)

            for lb in combined_labs:
                st.write(f"**Lab: {lb}**")
                df_lr = build_room_usage_df(
                    room=lb,
                    schedule_map=schedule_map,
                    is_lab=True,
                    DAYS=DAYS,
                    THEORY_TIMESLOTS=THEORY_TIMESLOTS,
                    TIMESLOT_LABELS=TIMESLOT_LABELS,
                    LAB_SLOTS=LAB_SLOTS,
                    LAB_SLOT_LABELS=LAB_SLOT_LABELS
                )
                st.table(df_lr)

            # Free slots
            st.header("Free Slots Summary")
            total_theory = len(DAYS)*len(THEORY_TIMESLOTS)
            total_lab = len(DAYS)*len(LAB_SLOTS)

            usage_theory = {r: 0 for r in theory_rooms}
            usage_lab = {l: 0 for l in combined_labs}

            for (day, slot, room), occupant in schedule_map.items():
                if room in theory_rooms and slot in THEORY_TIMESLOTS:
                    usage_theory[room] += 1
                elif room in combined_labs and slot in LAB_SLOTS:
                    usage_lab[room] += 1

            st.subheader("Theory Rooms")
            for r in theory_rooms:
                used = usage_theory[r]
                st.write(f"{r}: {total_theory - used} free / {total_theory} total")

            st.subheader("Lab Rooms")
            for l in combined_labs:
                used = usage_lab[l]
                st.write(f"{l}: {total_lab - used} free / {total_lab} total")


if __name__ == "__main__":
    main()
