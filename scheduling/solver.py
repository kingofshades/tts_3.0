# solver.py
from ortools.sat.python import cp_model
from .utils import build_sections_for_semester

def schedule_timetable(
    selected_semesters,
    semester_courses_map,
    section_sizes,
    usage_data,
    DAYS,
    THEORY_TIMESLOTS,
    TIMESLOT_LABELS,
    LAB_SLOTS,
    LAB_SLOT_LABELS,
    LAB_OVERLAP_MAP,
    theory_rooms,
    lab_rooms,           # all labs from CSV (including special)
    special_lab_rooms,   # e.g. { "NS125L": ["PhysicsLab1","PhysicsLab2"], "CC121L": ["DLDLab1","DLDLab2"] }
    section_size=50
):
    """
    Returns (schedule_map, semester_sections_map, new_allocations) or None if infeasible.
    new_allocations: list of (rtype, room_name, day, slot, occupant_label)
    """

    # ------------------------------------------------------------
    # 1) Distinguish normal vs. special labs
    # ------------------------------------------------------------
    # Gather all labs from 'special_lab_rooms'
    all_special_labs = set()
    for slist in special_lab_rooms.values():
        for labn in slist:
            # Make sure to strip() the name in case CSV has trailing spaces
            all_special_labs.add(labn.strip())

    # Convert the input lab_rooms to a stripped list (remove trailing spaces)
    lab_rooms = [lr.strip() for lr in lab_rooms]

    # normal_labs = any lab in lab_rooms that isn't in all_special_labs
    normal_labs = [lr for lr in lab_rooms if lr not in all_special_labs]

    # We'll combine them for capacity checks, but never for assignment of normal-lab courses
    combined_labs = set(normal_labs)
    for sl in all_special_labs:
        combined_labs.add(sl)
    combined_labs = list(combined_labs)

    # ------------------------------------------------------------
    # 2) Basic usage count for capacity
    # ------------------------------------------------------------
    theory_used = 0
    for r in usage_data["theory"]:
        for day in usage_data["theory"][r]:
            theory_used += len(usage_data["theory"][r][day])

    lab_used = 0
    for r in usage_data["lab"]:
        for day in usage_data["lab"][r]:
            lab_used += len(usage_data["lab"][r][day])

    total_theory_capacity = len(DAYS) * len(THEORY_TIMESLOTS) * len(theory_rooms)
    total_lab_capacity = len(DAYS) * len(LAB_SLOTS) * len(combined_labs)

    available_theory = total_theory_capacity - theory_used
    available_lab = total_lab_capacity - lab_used

    needed_theory = 0
    needed_lab = 0
    for sem in selected_semesters:
        for (code, cname, is_lab, times_needed) in semester_courses_map[sem]:
            if is_lab:
                needed_lab += times_needed
            else:
                needed_theory += times_needed

    if needed_theory > available_theory:
        raise ValueError(
            f"Not enough free theory slots. Need={needed_theory}, Have={available_theory}."
        )
    if needed_lab > available_lab:
        raise ValueError(
            f"Not enough free lab slots. Need={needed_lab}, Have={available_lab}."
        )

    # ------------------------------------------------------------
    # 3) Build sections
    # ------------------------------------------------------------
    semester_sections_map2 = {}
    all_sections = []
    for sem in selected_semesters:
        n_students = section_sizes[sem]
        secs = build_sections_for_semester(sem, n_students, section_size)
        semester_sections_map2[sem] = secs
        all_sections.extend(secs)

    # ------------------------------------------------------------
    # 4) Create model & variables
    # ------------------------------------------------------------
    model = cp_model.CpModel()
    assignments = {}
    day_assigned = {}

    for sem in selected_semesters:
        courses = semester_courses_map[sem]
        for section in semester_sections_map2[sem]:
            for (code, cname, is_lab, times_needed) in courses:
                # Distinct-day markers for theory
                for day in DAYS:
                    day_assigned[(section, code, day)] = model.NewBoolVar(
                        f"day_{section}_{code}_{day}"
                    )

                if is_lab:
                    # If code is "special-lab" => only special labs from dictionary
                    if code in special_lab_rooms:
                        # e.g. ["PhysicsLab1","PhysicsLab2"]
                        valid_labs = [x.strip() for x in special_lab_rooms[code]]
                    else:
                        # normal-lab code => only normal_labs (excludes PhysicsLab1,2, DLDLab1,2, etc.)
                        valid_labs = normal_labs

                    # Create bool vars only if usage not preoccupied
                    for day in DAYS:
                        for lslot in LAB_SLOTS:
                            for labr in valid_labs:
                                preused = usage_data["lab"].get(labr, {}).get(day, [])
                                if lslot in preused:
                                    continue
                                key = (section, code, day, lslot, labr)
                                assignments[key] = model.NewBoolVar(
                                    f"LabVar_{section}_{code}_{day}_{lslot}_{labr}"
                                )

                else:
                    # theory
                    for day in DAYS:
                        for t in THEORY_TIMESLOTS:
                            for r in theory_rooms:
                                preused = usage_data["theory"].get(r, {}).get(day, [])
                                if t in preused:
                                    continue
                                key = (section, code, day, t, r)
                                assignments[key] = model.NewBoolVar(
                                    f"TheoryVar_{section}_{code}_{day}_{t}_{r}"
                                )

    # ------------------------------------------------------------
    # 5) times_needed constraints
    # ------------------------------------------------------------
    for sem in selected_semesters:
        for section in semester_sections_map2[sem]:
            for (code, cname, is_lab, times_needed) in semester_courses_map[sem]:
                if is_lab:
                    if code in special_lab_rooms:
                        valid_labs = [x.strip() for x in special_lab_rooms[code]]
                    else:
                        valid_labs = normal_labs

                    lab_vars = []
                    for day in DAYS:
                        for lslot in LAB_SLOTS:
                            for labr in valid_labs:
                                key = (section, code, day, lslot, labr)
                                if key in assignments:
                                    lab_vars.append(assignments[key])
                    model.Add(sum(lab_vars) == times_needed)

                else:
                    # theory
                    th_vars = []
                    for day in DAYS:
                        for t in THEORY_TIMESLOTS:
                            for r in theory_rooms:
                                key = (section, code, day, t, r)
                                if key in assignments:
                                    th_vars.append(assignments[key])
                    model.Add(sum(th_vars) == times_needed)

    # ------------------------------------------------------------
    # 6) Distinct-day for theory (unchanged)
    # ------------------------------------------------------------
    for sem in selected_semesters:
        for section in semester_sections_map2[sem]:
            for (code, cname, is_lab, times_needed) in semester_courses_map[sem]:
                if is_lab:
                    continue
                for day in DAYS:
                    relevant_vars = []
                    for t in THEORY_TIMESLOTS:
                        for r in theory_rooms:
                            key = (section, code, day, t, r)
                            if key in assignments:
                                relevant_vars.append(assignments[key])
                    model.Add(sum(relevant_vars) >= day_assigned[(section, code, day)])
                    model.Add(sum(relevant_vars) <= len(THEORY_TIMESLOTS)*day_assigned[(section, code, day)])

                model.Add(
                    sum(day_assigned[(section, code, d)] for d in DAYS) == times_needed
                )

    # ------------------------------------------------------------
    # 7) No double-booking theory room
    # ------------------------------------------------------------
    for day in DAYS:
        for t in THEORY_TIMESLOTS:
            for r in theory_rooms:
                model.Add(
                    sum(
                        assignments.get((sec, c, day, t, r), 0)
                        for sec in all_sections
                        for sem2 in selected_semesters
                        for (c, cname, lab, tn) in semester_courses_map[sem2]
                        if not lab
                    ) <= 1
                )

    # ------------------------------------------------------------
    # 8) No double-booking lab room
    # ------------------------------------------------------------
    for day in DAYS:
        for ls in LAB_SLOTS:
            for labr in combined_labs:  # includes normal + special labs
                model.Add(
                    sum(
                        assignments.get((sec, c, day, ls, labr), 0)
                        for sec in all_sections
                        for sem2 in selected_semesters
                        for (c, cname, lab, tn) in semester_courses_map[sem2]
                        if lab
                    ) <= 1
                )

    # ------------------------------------------------------------
    # 9) Prevent same-section clashes
    # ------------------------------------------------------------
    for sec in all_sections:
        for day in DAYS:
            # (a) only one theory timeslot
            for t in THEORY_TIMESLOTS:
                model.Add(
                    sum(
                        assignments.get((sec, c, day, t, rr), 0)
                        for sem2 in selected_semesters
                        for (c, cname, lab, tn) in semester_courses_map[sem2]
                        for rr in theory_rooms
                        if not lab
                    ) <= 1
                )

            # (b) only one lab slot
            for ls in LAB_SLOTS:
                model.Add(
                    sum(
                        assignments.get((sec, c, day, ls, labr), 0)
                        for sem2 in selected_semesters
                        for (c, cname, lab, tn) in semester_courses_map[sem2]
                        for labr in combined_labs
                        if lab
                    ) <= 1
                )

            # (c) partial overlap (lab vs. theory)
            for sem2 in selected_semesters:
                for (c, cname, lab, tn) in semester_courses_map[sem2]:
                    if not lab:
                        continue
                    for labr in combined_labs:
                        for ls in LAB_SLOTS:
                            lab_var = assignments.get((sec, c, day, ls, labr), None)
                            if lab_var is None:
                                continue
                            overlap_list = LAB_OVERLAP_MAP[ls]
                            for sem3 in selected_semesters:
                                for (c2, _cn2, lb2, _tn2) in semester_courses_map[sem3]:
                                    if lb2:
                                        continue
                                    for rr2 in theory_rooms:
                                        for t2 in overlap_list:
                                            tvar = assignments.get((sec, c2, day, t2, rr2), None)
                                            if tvar is not None:
                                                model.Add(lab_var + tvar <= 1)

    # ------------------------------------------------------------
    # 10) Solve
    # ------------------------------------------------------------
    solver = cp_model.CpSolver()
    status = solver.Solve(model)
    if status not in [cp_model.FEASIBLE, cp_model.OPTIMAL]:
        return None

    # ------------------------------------------------------------
    # 11) Build final schedule
    # ------------------------------------------------------------
    schedule_map = {}
    new_allocations = []
    for sem in selected_semesters:
        for sec in semester_sections_map2[sem]:
            for (code, cname, lab, tn) in semester_courses_map[sem]:
                occupant = f"{sec}-{code}"
                if not lab:
                    # theory
                    for day in DAYS:
                        for t in THEORY_TIMESLOTS:
                            for r in theory_rooms:
                                key = (sec, code, day, t, r)
                                var = assignments.get(key)
                                if var is not None and solver.Value(var) == 1:
                                    schedule_map[(day, t, r)] = (sec, code)
                                    new_allocations.append(("theory", r, day, t, occupant))
                else:
                    # special-lab or normal-lab
                    if code in special_lab_rooms:
                        valid_labs = [x.strip() for x in special_lab_rooms[code]]
                    else:
                        valid_labs = normal_labs

                    for day in DAYS:
                        for ls in LAB_SLOTS:
                            for lr in valid_labs:
                                var = assignments.get((sec, code, day, ls, lr))
                                if var is not None and solver.Value(var) == 1:
                                    schedule_map[(day, ls, lr)] = (sec, code)
                                    new_allocations.append(("lab", lr, day, ls, occupant))

    return schedule_map, semester_sections_map2, new_allocations
