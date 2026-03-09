# solver_logic.py
import pandas as pd
import pulp
from collections import defaultdict
import math
import os
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
def run_optimization(input_excel_path, output_excel_path, category_diversity_weight=0):
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
    print(f"Starting optimization for input: {input_excel_path}")
    stats_summary = None
    try:
        # --- LECTURE DES DONNÉES (Keep as is) ---
        print("Lecture du fichier Excel...")
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
        categories = sorted(set(inst["category"] for inst in activity_instances if inst["category"]))
        category_workshops = defaultdict(list)
        for a_id, inst in activity_dict.items():
            if inst["category"]:
                category_workshops[inst["category"]].append(a_id)
        use_category_diversity = category_diversity_weight > 0 and len(categories) >= 2
        if use_category_diversity:
            print(f"{len(categories)} catégories détectées: {', '.join(categories)}")
        elif category_diversity_weight > 0 and len(categories) < 2:
            print("Attention: Moins de 2 catégories trouvées, diversité catégorielle désactivée.")
            use_category_diversity = False

        # --- PRÉPARATION DES PRÉFÉRENCES (Keep as is) ---
        print("Préparation des préférences élèves...")
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
        print("Mise en place du modèle...")
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
            y_cat = {s: {c: pulp.LpVariable(f"ycat_{s}_{c}".replace(' ','_').replace('/','_'), cat="Binary") for c in categories} for s in student_ids}
            category_penalty_term = category_diversity_weight * pulp.lpSum(
                len(categories) - pulp.lpSum(y_cat[s][c] for c in categories)
                for s in student_ids
            )

        prob += assignment_costs + deviation_costs + extra_neutral_penalty_term + category_penalty_term, "TotalCost"

        # --- CONTRAINTES (Keep as is) ---
        print("Ajout des contraintes...")
        for s in student_ids: prob += pulp.lpSum(activity_dict[a]['duration'] * x[s][a] for a in activity_dict) == TOTAL_SESSIONS, f"TotalDuration_{s}"
        for a in activity_dict: prob += pulp.lpSum(x[s][a] for s in student_ids) <= activity_dict[a]["max"], f"CapacitéMax_{a}"
        for a in activity_dict: n_a = pulp.lpSum(x[s][a] for s in student_ids); ideal = activity_dict[a]["ideal"]; prob += n_a - ideal <= dev[a], f"DevPos_{a}"; prob += ideal - n_a <= dev[a], f"DevNeg_{a}"
        unique_codes = set(inst["code"] for inst in activity_instances);
        for s in student_ids:
            for code in unique_codes:
                rel_inst = [a_id for a_id, inst in activity_dict.items() if inst["code"] == code];
                if len(rel_inst) > 1: prob += pulp.lpSum(x[s][a] for a in rel_inst) <= 1, f"UniqueCode_{s}_{code}"
        instances_covering_session = defaultdict(list);
        for a, inst_data in activity_dict.items():
            for sess in inst_data['sessions_covered']:
                if sess in session_indices: instances_covering_session[sess].append(a)
        for s in student_ids:
            safe_s_id = s.replace('/', '_').replace('\\', '_').replace(':', '_').replace(' ', '_')
            for sess in EXPECTED_SESSIONS:
                rel_inst = instances_covering_session[sess]
                if rel_inst: safe_sess = sess.replace(' ','_').replace('-','_').replace('/','_').replace(':','_'); prob += pulp.lpSum(x[s][a] for a in rel_inst) <= 1, f"Overlap_{safe_s_id}_{safe_sess}"
        for s in student_ids: prob += z[s] == pulp.lpSum(neutral_indicator[s][a] * x[s][a] for a in activity_dict), f"ZDef_{s}"; prob += w[s] >= z[s] - 1, f"WDef_{s}"

        # Category diversity constraints
        if use_category_diversity:
            print("Ajout des contraintes de diversité catégorielle...")
            for s in student_ids:
                safe_s = s.replace('/', '_').replace('\\', '_').replace(':', '_').replace(' ', '_')
                for c in categories:
                    safe_c = c.replace(' ','_').replace('/','_').replace(':','_').replace('-','_')
                    cat_inst = category_workshops[c]
                    # y_cat[s][c] = 1 if student has at least one workshop in category c
                    prob += y_cat[s][c] <= pulp.lpSum(x[s][a] for a in cat_inst), f"CatMax_{safe_s}_{safe_c}"
                    for a in cat_inst:
                        prob += y_cat[s][c] >= x[s][a], f"CatMin_{safe_s}_{safe_c}_{a}"

        # --- SOLVE THE MODEL (Keep as is) ---
        print("Résolution du modèle...")
        solver = pulp.PULP_CBC_CMD(msg=False, timeLimit=300)
        result_status = prob.solve(solver)
        print("Statut du solveur :", pulp.LpStatus[prob.status])
        optimal_solution_found = (prob.status == pulp.LpStatusOptimal)
        if not optimal_solution_found:
            status_text = pulp.LpStatus[prob.status]; msg = f"ERREUR: Solution optimale non trouvée (statut: {status_text})."
            if prob.status == pulp.LpStatusInfeasible: msg = f"ERREUR: Modèle infaisable (statut: {status_text}). Vérifiez capacités, vetos, structure."
            return False, msg, None

        # --- PREPARE OUTPUT ---
        print("Préparation des fichiers de sortie...")
        student_schedule_df = pd.DataFrame(columns=["Nom", "Prénom", "Classe"] + EXPECTED_SESSIONS)
        activity_schedule_df = pd.DataFrame()
        stats_df = pd.DataFrame(columns=["Statistique", "Valeur"])

        # 1. Create student schedule (Keep as is)
        output_rows_student = []
        for s in student_ids:
            stud = student_dict[s]; row = {"Nom": stud["nom"], "Prénom": stud["prenom"], "Classe": stud["classe"]}
            for session in EXPECTED_SESSIONS: row[session] = ""
            assigned_sessions_check = set(); assigned_duration_check = 0
            for a in activity_dict:
                if x[s][a] is not None and pulp.value(x[s][a]) == 1:
                    instance = activity_dict[a]; code = instance["code"]; covered_sessions = instance["sessions_covered"]
                    assigned_duration_check += instance['duration']
                    for session_covered in covered_sessions:
                        if session_covered in assigned_sessions_check: print(f"ERREUR LOGIQUE (inattendu): Chevauchement {s} session {session_covered}")
                        if session_covered in row: row[session_covered] = code; assigned_sessions_check.add(session_covered)
                        else: print(f"Attention: Session {session_covered} ({code}) pas dans colonnes attendues.")
            output_rows_student.append(row)
        student_schedule_df = pd.DataFrame(output_rows_student, columns=["Nom", "Prénom", "Classe"] + EXPECTED_SESSIONS)

        # 2. Create workshop schedule (Keep as is)
        instance_student_lists = defaultdict(list); max_students = 0
        for a in activity_dict:
            student_list = []
            for s in student_ids:
                if x[s][a] is not None and pulp.value(x[s][a]) == 1:
                    stud = student_dict[s]; student_list.append(f"{stud['nom']} {stud['prenom']} ({stud['classe']})")
            instance_student_lists[a] = sorted(student_list); max_students = max(max_students, len(student_list))
        row_labels = ["Code", "Description", "Enseignant", "Salle", "Session Début", "Sessions Couvertes", "Durée", "Max Élèves", "Idéal Élèves", "Nb Affectés", "--- Élèves ---"] + [f"Élève {i+1}" for i in range(max_students)]
        activity_data_for_df = {}
        for a in activity_dict:
            inst = activity_dict[a]; students = instance_student_lists[a]; sessions_str = ", ".join(inst["sessions_covered"])
            padded_students = students + [""] * (max_students - len(students))
            instance_column_data = [inst["code"], inst["Description"], inst["Enseignant"], inst["Salle"], inst["session_start"], sessions_str, inst["duration"], inst["max"], inst["ideal"], len(students), ""] + padded_students
            activity_data_for_df[f"Atelier_{inst['code']}_Inst{a}"] = instance_column_data
        activity_schedule_df = pd.DataFrame(activity_data_for_df)
        if not activity_schedule_df.empty: activity_schedule_df.index = row_labels

        # --- Calculate statistics AND Preference Distribution ---
        # KPI is session-based: a 2-session preferred workshop = 2 sessions satisfied out of TOTAL_SESSIONS
        stats_rows = [("Statut Final Solveur", pulp.LpStatus[prob.status])]
        total_pref_sessions = 0; total_veto_sessions = 0; total_neutral_sessions = 0
        total_deviation = sum(pulp.value(dev[a]) or 0 for a in dev)
        student_neutral_counts = defaultdict(int)
        student_pref_sessions = defaultdict(int)  # Sessions satisfied per student

        students_by_pref_sessions = defaultdict(list)

        for s in student_ids:
            neutral_sessions_for_student = 0
            pref_sessions_for_student = 0
            for a in activity_dict:
                if x[s][a] is not None and pulp.value(x[s][a]) == 1:
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

            student_neutral_counts[s] = neutral_sessions_for_student
            student_pref_sessions[s] = pref_sessions_for_student
            students_by_pref_sessions[pref_sessions_for_student].append(student_dict[s])

        # Distribution: how many students have X sessions satisfied (out of TOTAL_SESSIONS)
        prefs_distribution = defaultdict(int)
        for s_id in student_ids:
            prefs_distribution[student_pref_sessions[s_id]] += 1

        # Continue with other stats
        average_deviation = total_deviation / len(activity_instances) if activity_instances else 0
        obj_value = pulp.value(prob.objective) if prob.objective is not None else "N/A"
        total_student_sessions = TOTAL_SESSIONS * len(student_ids)
        pref_rate = f"{round((total_pref_sessions / total_student_sessions) * 100, 1)}%" if total_student_sessions else "N/A"

        # Category diversity stats
        category_diversity_distribution = {}
        if categories:
            student_category_counts = {}
            for s in student_ids:
                cats_covered = set()
                for a in activity_dict:
                    if x[s][a] is not None and pulp.value(x[s][a]) == 1:
                        cat = activity_dict[a]["category"]
                        if cat:
                            cats_covered.add(cat)
                student_category_counts[s] = len(cats_covered)
            cat_div_dist = defaultdict(int)
            for s in student_ids:
                key = f"{student_category_counts[s]}/{len(categories)}"
                cat_div_dist[key] += 1
            category_diversity_distribution = dict(cat_div_dist)

        # Create the stats dictionary to return
        stats_summary = {
            "objective_value": f"{obj_value:.2f}" if isinstance(obj_value, (int, float)) else "N/A",
            "pref_count": total_pref_sessions,
            "veto_count": total_veto_sessions,
            "neutral_count": total_neutral_sessions,
            "pref_rate": pref_rate,
            "total_deviation": f"{total_deviation:.2f}",
            "avg_deviation": f"{average_deviation:.2f}",
            "total_assignments": total_student_sessions,
            "students_processed": len(student_ids),
            "workshops_processed": len(activity_instances),
            "prefs_distribution": dict(prefs_distribution),
            "max_prefs_possible": TOTAL_SESSIONS,
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