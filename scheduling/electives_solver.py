# scheduling/electives_solver.py

from ortools.sat.python import cp_model

def schedule_electives(
    electives_list,
    usage_data,
    DAYS,
    THEORY_TIMESLOTS,
    LAB_SLOTS,
    theory_rooms,
    lab_rooms,
    timeslot_labels=None,
    lab_slot_labels=None,
    theory_needed=2,  # 2 timeslots if theory
    lab_needed=1      # 1 timeslot if lab
):
    """
    Schedules electives with distinct-day logic *only* for theory courses.
    - If choose_theory=1 => sum(theory_vars)=2, sum(day_assigned)=2
    - If choose_theory=0 => sum(lab_vars)=1, and no day_assigned (0)

    For purely lab-based electives (can_theory=False, can_lab=True),
      choose_theory=0 => no day-distinct constraint.

    Returns: (schedule_map, new_allocations) or None if infeasible
    """

    model = cp_model.CpModel()

    # ----------------------------------------------------------------
    # 1) Build sets of "used" combos from usage_data
    # ----------------------------------------------------------------
    used_theory = set()
    for rname in usage_data["theory"]:
        for day in usage_data["theory"][rname]:
            for ts in usage_data["theory"][rname][day]:
                used_theory.add((rname, day, ts))

    used_lab = set()
    for rname in usage_data["lab"]:
        for day in usage_data["lab"][rname]:
            for ls in usage_data["lab"][rname][day]:
                used_lab.add((rname, day, ls))

    # ----------------------------------------------------------------
    # 2) Build leftover theory/lab combos
    # ----------------------------------------------------------------
    theory_combos = []
    for r in theory_rooms:
        for d in DAYS:
            for ts in THEORY_TIMESLOTS:
                if (r, d, ts) not in used_theory:
                    theory_combos.append((r, d, ts))

    lab_combos = []
    for r in lab_rooms:
        for d in DAYS:
            for ls in LAB_SLOTS:
                if (r, d, ls) not in used_lab:
                    lab_combos.append((r, d, ls))

    # Optional: debug prints to see leftover distribution
    print("[DEBUG] leftover THEORY combos by day:")
    leftover_day_theory = {day: 0 for day in DAYS}
    for (r, d, ts) in theory_combos:
        leftover_day_theory[d] += 1
    for d in DAYS:
        print(f"  {d} => {leftover_day_theory[d]}")

    print("[DEBUG] leftover LAB combos by day:")
    leftover_day_lab = {day: 0 for day in DAYS}
    for (r, d, ls) in lab_combos:
        leftover_day_lab[d] += 1
    for d in DAYS:
        print(f"  {d} => {leftover_day_lab[d]}")

    # ----------------------------------------------------------------
    # 3) Define CP-SAT variables
    # ----------------------------------------------------------------
    assignments = {}    # (code, idx, rtype, room, d, slot) -> Bool
    choose_theory = {}  # (code, idx) => Bool (1 => theory, 0 => lab)
    day_assigned = {}   # (code, idx, day) => Bool, but used *only* if theory

    for elec in electives_list:
        e_code = elec["code"]
        sec_count = elec["sections_count"]
        can_th = elec["can_theory"]
        can_lb = elec["can_lab"]

        for idx in range(sec_count):
            # Decide: 1 => theory, 0 => lab
            choose_var = model.NewBoolVar(f"chooseTheory_{e_code}_{idx}")
            choose_theory[(e_code, idx)] = choose_var

            if not can_th:
                # must do lab
                model.Add(choose_var == 0)
            if not can_lb:
                # must do theory
                model.Add(choose_var == 1)

            # day_assigned used ONLY for theory
            # We'll create the variable for each day, but we'll see how it's forced to 0 if choose_theory=0
            for d in DAYS:
                day_assigned[(e_code, idx, d)] = model.NewBoolVar(f"dayAsg_{e_code}_{idx}_{d}")

            # Create combos
            for (room, dd, ts) in theory_combos:
                var = model.NewBoolVar(f"T_{e_code}_{idx}_{room}_{dd}_{ts}")
                assignments[(e_code, idx, "theory", room, dd, ts)] = var
                # If we are not using theory => var=0
                model.Add(var <= choose_var)

            for (room, dd, ls) in lab_combos:
                var = model.NewBoolVar(f"L_{e_code}_{idx}_{room}_{dd}_{ls}")
                assignments[(e_code, idx, "lab", room, dd, ls)] = var
                # If we are using theory => var=0
                model.Add(var <= (1 - choose_var))

    # ----------------------------------------------------------------
    # 4) Distinct-day logic (theory only)
    # ----------------------------------------------------------------
    # For a theory-based elective => sum(day_assigned)=theory_needed(=2)
    # For a lab-based elective => sum(day_assigned)=0
    #   (since day_assigned is not relevant for labs)
    for elec in electives_list:
        e_code = elec["code"]
        sec_count = elec["sections_count"]
        for idx in range(sec_count):
            # sum(day_assigned) => 2 if theory, else 0 if lab
            model.Add(
                sum(day_assigned[(e_code, idx, d)] for d in DAYS)
                == theory_needed * choose_theory[(e_code, idx)]
            )

            # link day_assigned to theory combos
            for d in DAYS:
                # relevant theory combos that day
                relevant_th = []
                for (room, dd, ts) in theory_combos:
                    if dd == d:
                        relevant_th.append(assignments[(e_code, idx, "theory", room, dd, ts)])
                # If day_assigned=1 => must have >=1 theory combos that day
                # (and <= len(THEORY_TIMESLOTS) combos)
                model.Add(sum(relevant_th) >= day_assigned[(e_code, idx, d)])
                model.Add(sum(relevant_th) <= len(THEORY_TIMESLOTS)*day_assigned[(e_code, idx, d)])

    # ----------------------------------------------------------------
    # 5) Timeslot constraints
    # ----------------------------------------------------------------
    # If theory => sum(theory combos)=2
    # If lab => sum(lab combos)=1
    for elec in electives_list:
        e_code = elec["code"]
        sec_count = elec["sections_count"]
        for idx in range(sec_count):
            theory_list = []
            lab_list = []
            for (room, d, ts) in theory_combos:
                theory_list.append(assignments[(e_code, idx, "theory", room, d, ts)])
            for (room, d, ls) in lab_combos:
                lab_list.append(assignments[(e_code, idx, "lab", room, d, ls)])

            model.Add(sum(theory_list) == theory_needed * choose_theory[(e_code, idx)])
            model.Add(sum(lab_list) == lab_needed * (1 - choose_theory[(e_code, idx)]))

    # ----------------------------------------------------------------
    # 6) No double-booking
    # ----------------------------------------------------------------
    # theory
    for (room, dd, ts) in theory_combos:
        model.Add(
            sum(
                assignments[(elec["code"], i, "theory", room, dd, ts)]
                for elec in electives_list
                for i in range(elec["sections_count"])
            ) <= 1
        )
    # lab
    for (room, dd, ls) in lab_combos:
        model.Add(
            sum(
                assignments[(elec["code"], i, "lab", room, dd, ls)]
                for elec in electives_list
                for i in range(elec["sections_count"])
            ) <= 1
        )

    # ----------------------------------------------------------------
    # 7) Solve
    # ----------------------------------------------------------------
    solver = cp_model.CpSolver()
    status = solver.Solve(model)
    if status not in [cp_model.FEASIBLE, cp_model.OPTIMAL]:
        print("[DEBUG] => No feasible solution found for electives.")
        return None

    print("[DEBUG] => Found a feasible solution for electives!")

    # ----------------------------------------------------------------
    # 8) Build solution
    # ----------------------------------------------------------------
    schedule_map = {}
    new_allocations = []

    for elec in electives_list:
        e_code = elec["code"]
        sec_count = elec["sections_count"]
        for idx in range(sec_count):
            assigned_slots = []
            occupant_label = f"Elective-{e_code}-A{idx+1}"

            # gather theory
            for (room, d, ts) in theory_combos:
                if solver.Value(assignments[(e_code, idx, "theory", room, d, ts)]) == 1:
                    assigned_slots.append(("theory", room, d, ts))
            # gather lab
            for (room, d, ls) in lab_combos:
                if solver.Value(assignments[(e_code, idx, "lab", room, d, ls)]) == 1:
                    assigned_slots.append(("lab", room, d, ls))

            schedule_map[(e_code, idx)] = assigned_slots
            for (rtype, rname, day, slot) in assigned_slots:
                new_allocations.append((rtype, rname, day, slot, occupant_label))

    return schedule_map, new_allocations
