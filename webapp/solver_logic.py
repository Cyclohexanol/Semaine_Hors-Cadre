# solver_logic.py
import pandas as pd
import pulp
from collections import defaultdict
import math
import os
import time
import traceback

# --- Parameters and Config (Keep as is) ---
PREF_REWARD = 10
VETO_PENALTY = 1000
DEVIATION_WEIGHT = 50
EXTRA_NEUTRAL_PENALTY = 25
ATELIERS_SHEET = "Ateliers"
PREFERENCES_SHEET = "Preferences"
EXPECTED_SESSIONS = ["Lundi matin", "Lundi après-midi", "Mardi matin", "Mardi après-midi", "Mercredi matin"]
TOTAL_SESSIONS = len(EXPECTED_SESSIONS)
SESSION_COLUMNS = [f"Session {i}" for i in range(1, TOTAL_SESSIONS + 1)]
session_indices = {session: i for i, session in enumerate(EXPECTED_SESSIONS)}
# --- End Parameters ---


# --- Main optimization function ---
def run_optimization(input_excel_path, output_excel_path, category_diversity_weight=0, progress_callback=None):
    """
    Runs the planning optimization.

    Args:
        input_excel_path (str): Path to the uploaded Excel template.
        output_excel_path (str): Path where the resulting Excel should be saved.
        category_diversity_weight (float): Weight for category diversity penalty (0 = disabled).

    Returns:
        tuple: (success: bool, message: str, stats: dict | None)
               stats dictionary contains key metrics on success, otherwise None.
    """
    def _progress(step, pct):
        if progress_callback:
            progress_callback(step, pct)

    t_start = time.time()
    print(f"Starting optimization for input: {input_excel_path}")
    _progress("Démarrage...", 5)
    stats_summary = None
    try:
        # --- LECTURE DES DONNÉES (Keep as is) ---
        print("Lecture du fichier Excel...")
        _progress("Lecture des données...", 10)
        try:
            activities_df = pd.read_excel(input_excel_path, sheet_name=ATELIERS_SHEET)
            prefs_df = pd.read_excel(input_excel_path, sheet_name=PREFERENCES_SHEET)
        except FileNotFoundError: return False, f"ERREUR: Fichier d'entrée introuvable.", None
        except ValueError as e: sheet_name = str(e).split("'")[1] if "'" in str(e) else "[inconnu]"; return False, f"ERREUR: Onglet '{sheet_name}' introuvable.", None
        except Exception as e: return False, f"Erreur lecture Excel: {e}", None
        missing_session_cols = [col for col in SESSION_COLUMNS if col not in activities_df.columns];
        if missing_session_cols: return False, f"ERREUR: Colonnes manquantes '{ATELIERS_SHEET}': {', '.join(missing_session_cols)}", None

        # --- PRÉPARATION DES ATELIERS (Keep as is) ---
        print("Préparation des instances d'ateliers...")
        _progress("Préparation des ateliers...", 20)
        activity_instances = []
        activity_dict = {}
        for idx, row in activities_df.iterrows():
            instance_sessions = []; start_session = None; start_session_index = float('inf'); activity_code = row["Code"]
            for col_name in SESSION_COLUMNS:
                session_name = row[col_name]
                if pd.notna(session_name) and isinstance(session_name, str) and session_name.strip():
                    session_name = session_name.strip()
                    if session_name in session_indices:
                        if session_name not in instance_sessions:
                            instance_sessions.append(session_name); current_index = session_indices[session_name]
                            if current_index < start_session_index: start_session_index = current_index; start_session = session_name
                    else: print(f"Att L{idx+2} ({activity_code}): Sess '{session_name}' non reconnue.")
            if not instance_sessions: print(f"Att L{idx+2} ({activity_code}): Aucune session valide. Ligne ignorée."); continue
            duration = len(instance_sessions); instance_sessions.sort(key=lambda s: session_indices[s])
            category = str(row.get("Catégorie", "")).strip() if pd.notna(row.get("Catégorie", None)) else ""
            inst = { "instance_id": idx, "code": activity_code, "Description": row["Description"], "Enseignant": row["Enseignant"], "Salle": row["Salle"], "max": int(row["Nombre d'élèves max par session"]), "ideal": int(row["Nombre idéal d'élèves par session"]), "session_start": start_session, "duration": duration, "sessions_covered": instance_sessions, "category": category }
            activity_instances.append(inst)
        activity_dict = {inst["instance_id"]: inst for inst in activity_instances}
        if not activity_instances: return False, "ERREUR: Aucune instance d'atelier valide chargée.", None
        print(f"{len(activity_instances)} instances d'ateliers chargées.")

        # Build category mapping for diversity constraint
        category_workshops = defaultdict(list)
        for a_id, inst in activity_dict.items():
            if inst["category"]:
                category_workshops[inst["category"]].append(a_id)
        # Only keep categories that have at least one workshop
        categories = sorted(c for c in category_workshops if len(category_workshops[c]) > 0)
        use_category_diversity = category_diversity_weight > 0 and len(categories) >= 2
        if use_category_diversity:
            print(f"{len(categories)} catégories détectées: {', '.join(categories)}")
        elif category_diversity_weight > 0 and len(categories) < 2:
            print("Attention: Moins de 2 catégories trouvées, diversité catégorielle désactivée.")
            use_category_diversity = False

        # --- PRÉPARATION DES PRÉFÉRENCES (Keep as is) ---
        print("Préparation des préférences élèves...")
        _progress("Chargement des préférences...", 30)
        students = []
        student_dict = {}
        activity_codes_in_prefs = set(prefs_df.columns) - {"Nom", "Prénom", "Classe", "# Préférences"}
        for idx, row in prefs_df.iterrows():
            student_id = f"{row['Nom']}_{row['Prénom']}_{row['Classe']}_{idx}"
            prefs = defaultdict(int)
            for code in activity_codes_in_prefs:
                val = row[code]
                if pd.notna(val):
                    try: pref_val = int(float(val)); prefs[code] = pref_val if pref_val in [1, -1] else 0
                    except (ValueError, TypeError): pass
            students.append({"id": student_id, "nom": row["Nom"], "prenom": row["Prénom"], "classe": row["Classe"], "prefs": prefs})
        student_ids = [s["id"] for s in students]; student_dict = {s["id"]: s for s in students}
        if not student_ids: return False, "ERREUR: Aucun élève chargé.", None
        print(f"{len(student_ids)} élèves chargés.")


        # --- MODÈLE D'OPTIMISATION (Keep as is) ---
        t_model = time.time()
        print(f"Mise en place du modèle... (lecture: {t_model - t_start:.1f}s)")
        _progress("Construction du modèle...", 40)
        def _safe(name): return name.replace('/', '_').replace('\\', '_').replace(':', '_').replace(' ', '_').replace('-', '_')
        prob = pulp.LpProblem("PlanningAteliersWeb", pulp.LpMinimize)
        x = {s: {a_id: pulp.LpVariable(f"x_{s}_{a_id}", cat="Binary") for a_id in activity_dict} for s in student_ids}
        dev = {a_id: pulp.LpVariable(f"dev_{a_id}", lowBound=0, cat="Continuous") for a_id in activity_dict}
        neutral_indicator = {s: {a_id: 1 if student_dict[s]["prefs"][activity_dict[a_id]["code"]] == 0 else 0 for a_id in activity_dict} for s in student_ids}
        z = {s: pulp.LpVariable(f"z_{s}", lowBound=0, cat="Integer") for s in student_ids}
        w = {s: pulp.LpVariable(f"w_{s}", lowBound=0, cat="Continuous") for s in student_ids}
        assignment_costs = pulp.lpSum((-PREF_REWARD if student_dict[s]["prefs"][activity_dict[a_id]["code"]] == 1 else VETO_PENALTY if student_dict[s]["prefs"][activity_dict[a_id]["code"]] == -1 else 0) * x[s][a_id] for s in student_ids for a_id in activity_dict)
        deviation_costs = pulp.lpSum(DEVIATION_WEIGHT * dev[a_id] for a_id in dev)
        extra_neutral_penalty_term = pulp.lpSum(EXTRA_NEUTRAL_PENALTY * w[s] for s in student_ids)

        # Category diversity variables and penalty
        category_penalty_term = 0
        if use_category_diversity:
            safe_categories = {c: _safe(c) for c in categories}
            y_cat = {s: {c: pulp.LpVariable(f"ycat_{_safe(s)}_{safe_categories[c]}", lowBound=0, upBound=1, cat="Continuous") for c in categories} for s in student_ids}
            # Pre-fix y_cat=0 for impossible student-category pairs (all workshops vetoed)
            for s in student_ids:
                for c in categories:
                    if all(student_dict[s]["prefs"][activity_dict[a]["code"]] == -1 for a in category_workshops[c]):
                        y_cat[s][c].upBound = 0
            category_penalty_term = category_diversity_weight * pulp.lpSum(
                len(categories) - pulp.lpSum(y_cat[s][c] for c in categories)
                for s in student_ids
            )

        prob += assignment_costs + deviation_costs + extra_neutral_penalty_term + category_penalty_term, "TotalCost"

        # --- CONTRAINTES ---
        # Pre-compute lookups for constraint generation
        safe_student_ids = {s: _safe(s) for s in student_ids}
        code_to_instances = defaultdict(list)
        for a_id, inst in activity_dict.items():
            code_to_instances[inst["code"]].append(a_id)
        instances_covering_session = defaultdict(list)
        for a, inst_data in activity_dict.items():
            for sess in inst_data['sessions_covered']:
                if sess in session_indices: instances_covering_session[sess].append(a)
        safe_sessions = {sess: _safe(sess) for sess in EXPECTED_SESSIONS}

        t_constraints = time.time()
        print(f"Ajout des contraintes... (variables: {t_constraints - t_model:.1f}s)")
        _progress("Ajout des contraintes...", 55)
        for s in student_ids: prob += pulp.lpSum(activity_dict[a]['duration'] * x[s][a] for a in activity_dict) == TOTAL_SESSIONS, f"TotalDuration_{s}"
        for a in activity_dict: prob += pulp.lpSum(x[s][a] for s in student_ids) <= activity_dict[a]["max"], f"CapacitéMax_{a}"
        for a in activity_dict: n_a = pulp.lpSum(x[s][a] for s in student_ids); ideal = activity_dict[a]["ideal"]; prob += n_a - ideal <= dev[a], f"DevPos_{a}"; prob += ideal - n_a <= dev[a], f"DevNeg_{a}"
        for s in student_ids:
            ss = safe_student_ids[s]
            for code, rel_inst in code_to_instances.items():
                if len(rel_inst) > 1: prob += pulp.lpSum(x[s][a] for a in rel_inst) <= 1, f"UniqueCode_{ss}_{_safe(code)}"
        for s in student_ids:
            ss = safe_student_ids[s]
            for sess in EXPECTED_SESSIONS:
                rel_inst = instances_covering_session[sess]
                if rel_inst: prob += pulp.lpSum(x[s][a] for a in rel_inst) <= 1, f"Overlap_{ss}_{safe_sessions[sess]}"
        for s in student_ids: prob += z[s] == pulp.lpSum(neutral_indicator[s][a] * x[s][a] for a in activity_dict), f"ZDef_{s}"; prob += w[s] >= z[s] - 1, f"WDef_{s}"

        # Category diversity constraints
        # Only CatMax needed: y_cat[s][c] <= sum(x[s][a]) forces y=0 when no workshop assigned.
        # The objective maximizes y_cat, so the solver sets y=1 whenever allowed — no >= needed.
        if use_category_diversity:
            print("Ajout des contraintes de diversité catégorielle...")
            for s in student_ids:
                ss = safe_student_ids[s]
                for c in categories:
                    cat_inst = category_workshops[c]
                    prob += y_cat[s][c] <= pulp.lpSum(x[s][a] for a in cat_inst), f"CatMax_{ss}_{safe_categories[c]}"
                # Session-count upper bound: can't cover more categories than sessions
                prob += pulp.lpSum(y_cat[s][c] for c in categories) <= TOTAL_SESSIONS, f"CatSessionBound_{ss}"

        # --- SOLVE THE MODEL ---
        t_solve = time.time()
        cbc_threads = os.cpu_count() or 1

        # Warm-start: solve without category penalty first, then use as starting point
        if use_category_diversity:
            print(f"Phase 1: résolution sans diversité catégorielle (warm-start)...")
            # Temporarily replace objective with category-free version
            prob_warmup = pulp.LpProblem("WarmStart", pulp.LpMinimize)
            # Re-use the same x, dev, w variables in a lightweight problem
            x_warmup = {s: {a_id: pulp.LpVariable(f"xw_{s}_{a_id}", cat="Binary") for a_id in activity_dict} for s in student_ids}
            dev_warmup = {a_id: pulp.LpVariable(f"devw_{a_id}", lowBound=0, cat="Continuous") for a_id in activity_dict}
            z_warmup = {s: pulp.LpVariable(f"zw_{s}", lowBound=0, cat="Integer") for s in student_ids}
            w_warmup = {s: pulp.LpVariable(f"ww_{s}", lowBound=0, cat="Continuous") for s in student_ids}
            # Objective without category penalty
            warmup_assignment = pulp.lpSum((-PREF_REWARD if student_dict[s]["prefs"][activity_dict[a_id]["code"]] == 1 else VETO_PENALTY if student_dict[s]["prefs"][activity_dict[a_id]["code"]] == -1 else 0) * x_warmup[s][a_id] for s in student_ids for a_id in activity_dict)
            warmup_deviation = pulp.lpSum(DEVIATION_WEIGHT * dev_warmup[a_id] for a_id in dev_warmup)
            warmup_neutral = pulp.lpSum(EXTRA_NEUTRAL_PENALTY * w_warmup[s] for s in student_ids)
            prob_warmup += warmup_assignment + warmup_deviation + warmup_neutral, "WarmupCost"
            # Same structural constraints
            for s in student_ids: prob_warmup += pulp.lpSum(activity_dict[a]['duration'] * x_warmup[s][a] for a in activity_dict) == TOTAL_SESSIONS, f"TD_{s}"
            for a in activity_dict: prob_warmup += pulp.lpSum(x_warmup[s][a] for s in student_ids) <= activity_dict[a]["max"], f"CM_{a}"
            for a in activity_dict: n_a = pulp.lpSum(x_warmup[s][a] for s in student_ids); ideal = activity_dict[a]["ideal"]; prob_warmup += n_a - ideal <= dev_warmup[a], f"DP_{a}"; prob_warmup += ideal - n_a <= dev_warmup[a], f"DN_{a}"
            for s in student_ids:
                ss = safe_student_ids[s]
                for code, rel_inst in code_to_instances.items():
                    if len(rel_inst) > 1: prob_warmup += pulp.lpSum(x_warmup[s][a] for a in rel_inst) <= 1, f"UC_{ss}_{_safe(code)}"
            for s in student_ids:
                ss = safe_student_ids[s]
                for sess in EXPECTED_SESSIONS:
                    rel_inst = instances_covering_session[sess]
                    if rel_inst: prob_warmup += pulp.lpSum(x_warmup[s][a] for a in rel_inst) <= 1, f"OV_{ss}_{safe_sessions[sess]}"
            for s in student_ids: prob_warmup += z_warmup[s] == pulp.lpSum(neutral_indicator[s][a] * x_warmup[s][a] for a in activity_dict), f"ZD_{s}"; prob_warmup += w_warmup[s] >= z_warmup[s] - 1, f"WD_{s}"
            # Solve warm-up (fast)
            solver_warmup = pulp.PULP_CBC_CMD(msg=False, timeLimit=60, gapRel=0.01, threads=cbc_threads)
            prob_warmup.solve(solver_warmup)
            t_warmup = time.time()
            print(f"Phase 1 terminée: {pulp.LpStatus[prob_warmup.status]} ({t_warmup - t_solve:.1f}s)")
            # Transfer solution as warm-start
            if prob_warmup.status == pulp.LpStatusOptimal:
                for s in student_ids:
                    for a_id in activity_dict:
                        val = pulp.value(x_warmup[s][a_id])
                        if val is not None:
                            x[s][a_id].setInitialValue(val)
            print(f"Phase 2: résolution complète avec diversité catégorielle...")

        print(f"Résolution du modèle... (contraintes: {time.time() - t_constraints:.1f}s)")
        _progress("Résolution en cours...", 65)
        solver = pulp.PULP_CBC_CMD(
            msg=False,
            timeLimit=300,
            gapRel=0.03 if use_category_diversity else 0.01,
            threads=cbc_threads,
            warmStart=True if use_category_diversity else False,
        )
        result_status = prob.solve(solver)
        t_solved = time.time()
        print(f"Statut du solveur : {pulp.LpStatus[prob.status]} (résolution: {t_solved - t_solve:.1f}s, total: {t_solved - t_start:.1f}s)")
        optimal_solution_found = (prob.status == pulp.LpStatusOptimal)
        if not optimal_solution_found:
            status_text = pulp.LpStatus[prob.status]; msg = f"ERREUR: Solution optimale non trouvée (statut: {status_text})."
            if prob.status == pulp.LpStatusInfeasible: msg = f"ERREUR: Modèle infaisable (statut: {status_text}). Vérifiez capacités, vetos, structure."
            return False, msg, None

        # --- PREPARE OUTPUT ---
        print("Préparation des fichiers de sortie...")
        _progress("Préparation des résultats...", 85)

        # Cache all assignments once (avoids repeated pulp.value() calls)
        assignments = defaultdict(list)  # student -> list of assigned activity IDs
        for s in student_ids:
            for a in activity_dict:
                if x[s][a] is not None and pulp.value(x[s][a]) == 1:
                    assignments[s].append(a)

        # 1. Create student schedule
        output_rows_student = []
        for s in student_ids:
            stud = student_dict[s]; row = {"Nom": stud["nom"], "Prénom": stud["prenom"], "Classe": stud["classe"]}
            for session in EXPECTED_SESSIONS: row[session] = ""
            assigned_sessions_check = set()
            for a in assignments[s]:
                instance = activity_dict[a]; code = instance["code"]
                for session_covered in instance["sessions_covered"]:
                    if session_covered in assigned_sessions_check: print(f"ERREUR LOGIQUE (inattendu): Chevauchement {s} session {session_covered}")
                    if session_covered in row: row[session_covered] = code; assigned_sessions_check.add(session_covered)
                    else: print(f"Attention: Session {session_covered} ({code}) pas dans colonnes attendues.")
            output_rows_student.append(row)
        student_schedule_df = pd.DataFrame(output_rows_student, columns=["Nom", "Prénom", "Classe"] + EXPECTED_SESSIONS)

        # 2. Create workshop schedule
        instance_student_lists = defaultdict(list); max_students = 0
        for s in student_ids:
            stud = student_dict[s]; label = f"{stud['nom']} {stud['prenom']} ({stud['classe']})"
            for a in assignments[s]:
                instance_student_lists[a].append(label)
        for a in activity_dict:
            instance_student_lists[a] = sorted(instance_student_lists[a]); max_students = max(max_students, len(instance_student_lists[a]))
        row_labels = ["Code", "Description", "Enseignant", "Salle", "Session Début", "Sessions Couvertes", "Durée", "Max Élèves", "Idéal Élèves", "Nb Affectés", "--- Élèves ---"] + [f"Élève {i+1}" for i in range(max_students)]
        activity_data_for_df = {}
        for a in activity_dict:
            inst = activity_dict[a]; students_list = instance_student_lists[a]; sessions_str = ", ".join(inst["sessions_covered"])
            padded_students = students_list + [""] * (max_students - len(students_list))
            instance_column_data = [inst["code"], inst["Description"], inst["Enseignant"], inst["Salle"], inst["session_start"], sessions_str, inst["duration"], inst["max"], inst["ideal"], len(students_list), ""] + padded_students
            activity_data_for_df[f"Atelier_{inst['code']}_Inst{a}"] = instance_column_data
        activity_schedule_df = pd.DataFrame(activity_data_for_df)
        if not activity_schedule_df.empty: activity_schedule_df.index = row_labels

        # --- Calculate statistics AND Preference Distribution (single pass) ---
        # KPI is session-based: a 2-session preferred workshop = 2 sessions satisfied out of TOTAL_SESSIONS
        stats_rows = [("Statut Final Solveur", pulp.LpStatus[prob.status])]
        total_pref_sessions = 0; total_veto_sessions = 0; total_neutral_sessions = 0
        total_deviation = sum(pulp.value(dev[a]) or 0 for a in dev)
        student_neutral_counts = defaultdict(int)
        prefs_distribution = defaultdict(int)
        students_by_pref_sessions = defaultdict(list)
        track_categories = bool(categories)
        cat_div_dist = defaultdict(int)

        for s in student_ids:
            neutral_sessions_for_student = 0
            pref_sessions_for_student = 0
            cats_covered = set() if track_categories else None
            for a in assignments[s]:
                instance = activity_dict[a]; act_code = instance["code"]
                vote = student_dict[s]["prefs"][act_code]
                duration = instance["duration"]

                if vote == 1:
                    total_pref_sessions += duration
                    pref_sessions_for_student += duration
                elif vote == -1:
                    total_veto_sessions += duration
                else:
                    total_neutral_sessions += duration
                    neutral_sessions_for_student += duration

                if track_categories and instance["category"]:
                    cats_covered.add(instance["category"])

            student_neutral_counts[s] = neutral_sessions_for_student
            prefs_distribution[pref_sessions_for_student] += 1
            students_by_pref_sessions[pref_sessions_for_student].append(student_dict[s])
            if track_categories:
                cat_div_dist[f"{len(cats_covered)}/{len(categories)}"] += 1

        # Continue with other stats
        average_deviation = total_deviation / len(activity_instances) if activity_instances else 0
        obj_value = pulp.value(prob.objective) if prob.objective is not None else "N/A"
        total_student_sessions = TOTAL_SESSIONS * len(student_ids)
        pref_rate = f"{round((total_pref_sessions / total_student_sessions) * 100, 1)}%" if total_student_sessions else "N/A"

        category_diversity_distribution = dict(cat_div_dist) if track_categories else {}

        # Compute new KPIs
        num_students = len(student_ids)
        fully_satisfied = prefs_distribution.get(TOTAL_SESSIONS, 0)
        mostly_satisfied = sum(prefs_distribution.get(k, 0) for k in range(3, TOTAL_SESSIONS + 1))
        pref_rate_float = round((total_pref_sessions / total_student_sessions) * 100, 1) if total_student_sessions else 0

        # Create the stats dictionary to return
        stats_summary = {
            "objective_value": f"{obj_value:.2f}" if isinstance(obj_value, (int, float)) else "N/A",
            "pref_count": total_pref_sessions,
            "veto_count": total_veto_sessions,
            "neutral_count": total_neutral_sessions,
            "pref_rate": pref_rate,
            "pref_rate_float": pref_rate_float,
            "total_deviation": f"{total_deviation:.2f}",
            "avg_deviation": f"{average_deviation:.2f}",
            "total_assignments": total_student_sessions,
            "students_processed": num_students,
            "workshops_processed": len(activity_instances),
            "prefs_distribution": dict(prefs_distribution),
            "max_prefs_possible": TOTAL_SESSIONS,
            "fully_satisfied_count": fully_satisfied,
            "fully_satisfied_pct": f"{100 * fully_satisfied / num_students:.1f}%" if num_students else "N/A",
            "mostly_satisfied_count": mostly_satisfied,
            "mostly_satisfied_pct": f"{100 * mostly_satisfied / num_students:.1f}%" if num_students else "N/A",
            "category_diversity_distribution": category_diversity_distribution,
            "categories": categories,
            "category_diversity_weight": category_diversity_weight
        }

        # Add stats to full stats DataFrame (for Excel)
        stats_rows.extend([ ("Valeur Objectif Calculée", stats_summary["objective_value"]), ("Nb sessions Préférence", total_pref_sessions), ("Nb sessions Veto", total_veto_sessions), ("Nb sessions Neutre", total_neutral_sessions), ("Taux Préférence (sessions)", pref_rate), ("Déviation totale", stats_summary["total_deviation"]), ("Déviation moyenne/instance", stats_summary["avg_deviation"]) ])
        neutral_distribution = defaultdict(int); max_neutral_observed = 0
        for s in student_ids: count = student_neutral_counts[s]; neutral_distribution[count] += 1; max_neutral_observed = max(max_neutral_observed, count)
        stats_rows.append(("--- Analyse Choix Neutres / Élève ---", ""))
        for i in range(max_neutral_observed + 1): stats_rows.append((f"Nb élèves avec {i} neutres", neutral_distribution[i]))
        # Preference distribution (sessions satisfied out of TOTAL_SESSIONS)
        stats_rows.append(("--- Analyse Préférences / Élève (sessions) ---", ""))
        for i in sorted(prefs_distribution.keys(), reverse=True):
            stats_rows.append((f"Nb élèves avec {i}/{TOTAL_SESSIONS} sessions préférées", prefs_distribution[i]))

        # Category diversity distribution
        if category_diversity_distribution:
            stats_rows.append(("--- Analyse Diversité Catégorielle ---", ""))
            stats_rows.append((f"Poids diversité catégorielle", category_diversity_weight))
            sorted_cat_keys = sorted(category_diversity_distribution.keys(), key=lambda k: int(k.split('/')[0]), reverse=True)
            for key in sorted_cat_keys:
                stats_rows.append((f"Nb élèves couvrant {key} catégories", category_diversity_distribution[key]))

        stats_df = pd.DataFrame(stats_rows, columns=["Statistique", "Valeur"])


        # Prepare DataFrame for students by preference sessions
        pref_count_student_data = []
        for sessions_key in sorted(students_by_pref_sessions.keys()):
            for student_detail in students_by_pref_sessions[sessions_key]:
                taux = round((sessions_key / TOTAL_SESSIONS) * 100, 1)
                pref_count_student_data.append({
                    "Sessions Préférées": f"{sessions_key}/{TOTAL_SESSIONS}",
                    "Taux": f"{taux}%",
                    "Nom": student_detail["nom"],
                    "Prénom": student_detail["prenom"],
                    "Classe": student_detail["classe"]
                })
        students_by_pref_df = pd.DataFrame(pref_count_student_data)


        # --- WRITE OUTPUT (Keep as is) ---
        print(f"Écriture du fichier de sortie '{output_excel_path}'...")
        _progress("Écriture du fichier...", 95)
        try:
            with pd.ExcelWriter(output_excel_path, engine="xlsxwriter") as writer:
                student_schedule_df.to_excel(writer, sheet_name="Planning par élève", index=False)
                if not activity_schedule_df.empty:
                    activity_schedule_df.to_excel(writer, sheet_name="Planning par Atelier", index=True, header=True)
                stats_df.to_excel(writer, sheet_name="Résumé et Stats", index=False)
                # ADDED: Write the new sheet
                if not students_by_pref_df.empty:
                    students_by_pref_df.to_excel(writer, sheet_name="Élèves par Nb Préférences", index=False)

            print("Écriture réussie.")
            success_msg = "Optimisation terminée avec succès."
            return True, success_msg, stats_summary

        except Exception as e:
            print(f"ERREUR lors de l'écriture du fichier Excel: {e}")
            return False, f"ERREUR lors de l'écriture du fichier Excel: {e}", None

    except Exception as e:
        print(f"ERREUR inattendue dans run_optimization: {e}")
        print(traceback.format_exc())
        return False, f"Une erreur serveur inattendue est survenue.", None