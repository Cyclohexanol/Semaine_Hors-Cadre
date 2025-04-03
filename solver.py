#!/usr/bin/env python3
"""
Schedule Optimizer for a School Activity Week (French Template)
Handles multi-session workshops defined by explicit columns in the template.
Includes anti-overlap constraint, correct status checking.
Outputs simplified student plan and transposed workshop plan with labels in the first column.

Input file: Template.xlsx
Tabs:
  1. "Ateliers" with columns:
       "Code", "Description", "Enseignant", "Salle", "Nombre de périodes",
       "Nombre d'élèves max par session", "Nombre idéal d'élèves par session",
       "Session 1", "Session 2", ..., "Session N" (for max N sessions)
  2. "Preferences" with columns:
       "Nom", "Prénom", "Classe", "# Préférences", and then one column per activity code

Chaque élève doit avoir son planning rempli pour toutes les sessions disponibles, sans chevauchement.
Un atelier peut durer plusieurs sessions, définies explicitement dans les colonnes Session X.
Si un élève est affecté à un atelier multi-sessions, il est occupé pour toutes ces sessions listées.
Aucun atelier ne doit dépasser son "Nombre d'élèves max par session".
Un même atelier (même code) ne doit pas être affecté plus d'une fois à un élève.
L'objectif est de valoriser les préférences (1) et pénaliser les veto (-1)
tout en s'approchant du nombre idéal d'élèves par session pour chaque instance.
Une pénalité non linéaire est ajoutée pour les affectations neutres :
chaque affectation neutre au-delà de la première pour un élève est pénalisée.
"""

import pandas as pd
import pulp
from collections import defaultdict
import math # For isnan

# ------------------------------
# PARAMETERS (ajustez si nécessaire)
# ------------------------------

PREF_REWARD = 10          # Récompense (coût négatif) pour une préférence
VETO_PENALTY = 1000       # Pénalité pour un veto
DEVIATION_WEIGHT = 1      # Pondération sur la déviation (écart par rapport à l'idéal)
EXTRA_NEUTRAL_PENALTY = 25  # Pénalité additionnelle pour chaque neutral en excès (au-delà de 1) par élève

# ------------------------------
# CHEMINS ET NOMS DES FEUILLES
# ------------------------------
EXCEL_FILE = "Template.xlsx"
ATELIERS_SHEET = "Ateliers"
PREFERENCES_SHEET = "Preferences"

# ------------------------------
# CONFIGURATION DES SESSIONS (IMPORTANT)
# ------------------------------
# Define the complete, ordered list of session names used throughout the week.
EXPECTED_SESSIONS = ["Lundi matin", "Lundi après-midi", "Mardi matin", "Mardi après-midi", "Mercredi matin"]
TOTAL_SESSIONS = len(EXPECTED_SESSIONS)
# Define the names of the columns in the Excel sheet that list the sessions for each activity.
SESSION_COLUMNS = [f"Session {i}" for i in range(1, TOTAL_SESSIONS + 1)]
# Internal mapping for quick lookup and order validation
session_indices = {session: i for i, session in enumerate(EXPECTED_SESSIONS)}


# ------------------------------
# LECTURE DES DONNÉES DANS L'EXCEL
# ------------------------------

print("Lecture du fichier Excel...")
try:
    activities_df = pd.read_excel(EXCEL_FILE, sheet_name=ATELIERS_SHEET)
    prefs_df = pd.read_excel(EXCEL_FILE, sheet_name=PREFERENCES_SHEET)
except FileNotFoundError:
    print(f"ERREUR: Le fichier '{EXCEL_FILE}' est introuvable.")
    exit()
except ValueError as e:
    sheet_name = str(e).split("'")[1] if "'" in str(e) else "[inconnu]"
    print(f"ERREUR: L'onglet '{sheet_name}' est introuvable dans '{EXCEL_FILE}'. Vérifiez les noms des feuilles.")
    exit()

# Vérification des colonnes Session X
missing_session_cols = [col for col in SESSION_COLUMNS if col not in activities_df.columns]
if missing_session_cols:
    print(f"ERREUR: Colonnes manquantes dans '{ATELIERS_SHEET}': {', '.join(missing_session_cols)}")
    exit()

# ------------------------------
# PRÉPARATION DES ATELIERS (INSTANCES)
# ------------------------------

print("Préparation des instances d'ateliers...")
activity_instances = []
activities_starting_in_session = defaultdict(list)

for idx, row in activities_df.iterrows():
    instance_sessions = []
    start_session = None
    start_session_index = float('inf')
    activity_code = row["Code"]

    for col_name in SESSION_COLUMNS:
        session_name = row[col_name]
        if pd.notna(session_name) and isinstance(session_name, str) and session_name.strip():
            session_name = session_name.strip()
            if session_name in session_indices:
                if session_name not in instance_sessions:
                    instance_sessions.append(session_name)
                    current_index = session_indices[session_name]
                    if current_index < start_session_index:
                        start_session_index = current_index
                        start_session = session_name
                # else: ignore duplicate silently now
            else:
                print(f"Attention ligne {idx+2} ({activity_code}): Session '{session_name}' non reconnue.")

    if not instance_sessions:
        print(f"Attention ligne {idx+2} ({activity_code}): Aucun nom de session valide trouvé. Ligne ignorée.")
        continue

    duration = len(instance_sessions)
    instance_sessions.sort(key=lambda s: session_indices[s])

    inst = {
        "instance_id": idx, "code": activity_code, "Description": row["Description"],
        "Enseignant": row["Enseignant"], "Salle": row["Salle"],
        "max": int(row["Nombre d'élèves max par session"]),
        "ideal": int(row["Nombre idéal d'élèves par session"]),
        "session_start": start_session, "duration": duration, "sessions_covered": instance_sessions
    }
    activity_instances.append(inst)
    if start_session:
        activities_starting_in_session[start_session].append(inst["instance_id"])

activity_dict = {inst["instance_id"]: inst for inst in activity_instances}

print(f"{len(activity_instances)} instances d'ateliers chargées.")
if not activity_instances: print("ERREUR: Aucune instance d'atelier valide chargée."); exit()
print("Ordre des sessions utilisé pour le planning:", EXPECTED_SESSIONS)


# ------------------------------
# PRÉPARATION DES PRÉFÉRENCES DES ÉLÈVES
# ------------------------------

print("Préparation des préférences élèves...")
students = []
activity_codes_in_prefs = set(prefs_df.columns) - {"Nom", "Prénom", "Classe", "# Préférences"}

for idx, row in prefs_df.iterrows():
    student_id = f"{row['Nom']}_{row['Prénom']}_{row['Classe']}_{idx}"
    prefs = defaultdict(int)
    for code in activity_codes_in_prefs:
        val = row[code]
        if pd.notna(val):
            try:
                pref_val = int(float(val))
                if pref_val in [1, -1]: prefs[code] = pref_val
            except (ValueError, TypeError): pass
    students.append({"id": student_id, "nom": row["Nom"], "prenom": row["Prénom"], "classe": row["Classe"], "prefs": prefs})

student_ids = [s["id"] for s in students]
student_dict = {s["id"]: s for s in students}

print(f"{len(student_ids)} élèves chargés.")
if not student_ids: print("ERREUR: Aucun élève chargé."); exit()

# ------------------------------
# MODÈLE D'OPTIMISATION
# ------------------------------

print("Mise en place du modèle d'optimisation...")
prob = pulp.LpProblem("PlanningAteliersMultiSessionColonnes", pulp.LpMinimize)

x = {s: {inst["instance_id"]: pulp.LpVariable(f"x_{s}_{inst['instance_id']}", cat="Binary")
         for inst in activity_instances} for s in student_ids}
dev = {inst["instance_id"]: pulp.LpVariable(f"dev_{inst['instance_id']}", lowBound=0, cat="Continuous")
       for inst in activity_instances}
neutral_indicator = {s: {inst["instance_id"]: 1 if student_dict[s]["prefs"][inst["code"]] == 0 else 0
                         for inst in activity_instances} for s in student_ids}

assignment_costs = [(-PREF_REWARD if student_dict[s]["prefs"][inst["code"]] == 1 else
                     VETO_PENALTY if student_dict[s]["prefs"][inst["code"]] == -1 else 0) * x[s][inst["instance_id"]]
                    for s in student_ids for inst in activity_instances]
deviation_costs = [DEVIATION_WEIGHT * dev[a] for a in dev]
base_objective = pulp.lpSum(assignment_costs + deviation_costs)

# ------------------------------
# CONTRAINTES
# ------------------------------

print("Ajout des contraintes...")
# (1) Total Duration
for s in student_ids:
    prob += pulp.lpSum(activity_dict[a]['duration'] * x[s][a] for a in activity_dict) == TOTAL_SESSIONS, f"TotalDuration_{s}"
# (2) Max Capacity
for inst in activity_instances:
    a = inst["instance_id"]
    prob += pulp.lpSum(x[s][a] for s in student_ids) <= inst["max"], f"CapacitéMax_{a}"
# (3) Deviation
for inst in activity_instances:
    a = inst["instance_id"]; n_a = pulp.lpSum(x[s][a] for s in student_ids); ideal = inst["ideal"]
    prob += n_a - ideal <= dev[a], f"DevPos_{a}"; prob += ideal - n_a <= dev[a], f"DevNeg_{a}"
# (4) Unique Code per Student
unique_codes = set(inst["code"] for inst in activity_instances)
for s in student_ids:
    for code in unique_codes:
        rel_inst = [inst["instance_id"] for inst in activity_instances if inst["code"] == code]
        if len(rel_inst) > 1: prob += pulp.lpSum(x[s][a] for a in rel_inst) <= 1, f"UniqueCode_{s}_{code}"
# (5) Anti-Overlap per Session per Student
print("Ajout de la contrainte anti-chevauchement par session...")
instances_covering_session = defaultdict(list)
for a, inst_data in activity_dict.items():
    for sess in inst_data['sessions_covered']:
        if sess in session_indices: instances_covering_session[sess].append(a)
for s in student_ids:
    safe_s_id = s.replace('/', '_').replace('\\', '_').replace(':', '_').replace(' ', '_')
    for sess in EXPECTED_SESSIONS:
        rel_inst = instances_covering_session[sess]
        if rel_inst:
            safe_sess = sess.replace(' ','_').replace('-','_').replace('/','_').replace(':','_')
            prob += pulp.lpSum(x[s][a] for a in rel_inst) <= 1, f"Overlap_{safe_s_id}_{safe_sess}"

# ------------------------------
# PENALTY FOR CONCENTRATED NEUTRAL CHOICES
# ------------------------------
print("Ajout de la pénalité pour choix neutres...")
z = {}; w = {}
for s in student_ids:
    z[s] = pulp.LpVariable(f"z_{s}", lowBound=0, cat="Integer"); w[s] = pulp.LpVariable(f"w_{s}", lowBound=0, cat="Continuous")
    prob += z[s] == pulp.lpSum(neutral_indicator[s][a] * x[s][a] for a in activity_dict), f"ZDef_{s}"
    prob += w[s] >= z[s] - 1, f"WDef_{s}"
extra_neutral_penalty_term = EXTRA_NEUTRAL_PENALTY * pulp.lpSum(w[s] for s in student_ids)
prob += base_objective + extra_neutral_penalty_term, "TotalCost"

# ------------------------------
# SOLVE THE MODEL
# ------------------------------
print("Résolution du modèle...")
solver = pulp.PULP_CBC_CMD(msg=True, timeLimit=300)
result_status = prob.solve(solver)
print("Statut du solveur :", pulp.LpStatus[prob.status])
optimal_solution_found = (prob.status == pulp.LpStatusOptimal)
if not optimal_solution_found: print_solver_warning(prob.status) # Define helper below if needed

# ------------------------------
# PREPARE OUTPUT
# ------------------------------
print("Préparation des fichiers de sortie...")
student_schedule_df = pd.DataFrame(columns=["Nom", "Prénom", "Classe"] + EXPECTED_SESSIONS)
activity_schedule_df = pd.DataFrame() # Will be built differently
stats_df = pd.DataFrame(columns=["Statistique", "Valeur"])

if optimal_solution_found:
    # 1. Create student schedule (simplified code only)
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
        if assigned_duration_check != TOTAL_SESSIONS: print(f"ATTENTION: {s} durée {assigned_duration_check}/{TOTAL_SESSIONS}")
        if len(assigned_sessions_check) != TOTAL_SESSIONS: print(f"ATTENTION: {s} sessions uniques {len(assigned_sessions_check)}/{TOTAL_SESSIONS}")
        output_rows_student.append(row)
    student_schedule_df = pd.DataFrame(output_rows_student, columns=["Nom", "Prénom", "Classe"] + EXPECTED_SESSIONS)

    # 2. Create workshop/activity schedule (transposed with labels)
    print("Préparation du planning par atelier (format transposé)...")
    instance_student_lists = defaultdict(list)
    max_students = 0
    for inst in activity_instances: # First pass: collect students and find max count
        a = inst["instance_id"]; student_list = []
        for s in student_ids:
            if x[s][a] is not None and pulp.value(x[s][a]) == 1:
                stud = student_dict[s]; student_list.append(f"{stud['nom']} {stud['prenom']} ({stud['classe']})")
        instance_student_lists[a] = sorted(student_list)
        if len(student_list) > max_students: max_students = len(student_list)

    # Define the exact order and labels for rows in the output
    row_labels = ["Code", "Description", "Enseignant", "Salle", "Session Début",
                  "Sessions Couvertes", "Durée", "Max Élèves", "Idéal Élèves", "Nb Affectés",
                  "--- Élèves ---"] + [f"Élève {i+1}" for i in range(max_students)]

    # Second pass: build the dictionary for the DataFrame (instances as columns)
    activity_data_for_df = {}
    for inst in activity_instances:
        a = inst["instance_id"]; students = instance_student_lists[a]
        sessions_str = ", ".join(inst["sessions_covered"])
        padded_students = students + [""] * (max_students - len(students))
        instance_column_data = [
            inst["code"], inst["Description"], inst["Enseignant"], inst["Salle"], inst["session_start"],
            sessions_str, inst["duration"], inst["max"], inst["ideal"], len(students),
            "" # Placeholder for '--- Élèves ---' separator row
        ] + padded_students
        activity_data_for_df[f"Atelier_{inst['code']}_Inst{a}"] = instance_column_data # Use more descriptive column header

    # Create DataFrame and set the index explicitly
    activity_schedule_df = pd.DataFrame(activity_data_for_df)
    activity_schedule_df.index = row_labels

    # 3. Calculate statistics
    stats = [("Statut Final Solveur", pulp.LpStatus[prob.status])]
    total_assignments_made = 0; pref_count = 0; veto_count = 0; neutral_assign_count = 0
    total_deviation = sum(pulp.value(dev[a]) or 0 for a in dev)
    total_ideal = sum(inst["ideal"] for inst in activity_instances)
    student_neutral_counts = defaultdict(int)
    for s in student_ids:
        num_neutral_for_student = 0
        for a in activity_dict:
            if x[s][a] is not None and pulp.value(x[s][a]) == 1:
                total_assignments_made += 1; instance = activity_dict[a]; act_code = instance["code"]
                vote = student_dict[s]["prefs"][act_code]
                if vote == 1: pref_count += 1
                elif vote == -1: veto_count += 1
                else: neutral_assign_count += 1; num_neutral_for_student += 1
        student_neutral_counts[s] = num_neutral_for_student
    average_deviation = total_deviation / len(activity_instances) if activity_instances else 0
    obj_value = pulp.value(prob.objective) if prob.objective is not None else "N/A"
    stats.extend([
        ("Valeur Objectif Calculée", obj_value),
        ("  Coût Préférences (-)", -PREF_REWARD * pref_count),
        ("  Coût Veto (+)", VETO_PENALTY * veto_count),
        ("  Coût Déviation (+)", DEVIATION_WEIGHT * total_deviation),
        ("  Coût Neutres Excédentaires (+)", EXTRA_NEUTRAL_PENALTY * sum(pulp.value(w[s]) or 0 for s in student_ids)),
        ("--- Stats Affectations ---", ""),
        ("Total affectations (x[s][a]=1)", total_assignments_made),
        ("Total slots élève-session", len(student_ids) * TOTAL_SESSIONS),
        ("Nb affectations Préférence", pref_count), ("Nb affectations Veto", veto_count), ("Nb affectations Neutre", neutral_assign_count),
        ("Taux Préférence", f"{round((pref_count / total_assignments_made) * 100, 1)}%" if total_assignments_made else "N/A"),
        ("Taux Veto", f"{round((veto_count / total_assignments_made) * 100, 1)}%" if total_assignments_made else "N/A"),
        ("Taux Neutre", f"{round((neutral_assign_count / total_assignments_made) * 100, 1)}%" if total_assignments_made else "N/A"),
        ("--- Stats Ateliers (Instances) ---", ""),
        ("Déviation totale", round(total_deviation, 2)), ("Déviation moyenne/instance", round(average_deviation, 2)),
        ("Total idéal élèves", total_ideal),
        ("--- Analyse Choix Neutres / Élève ---", ""),
    ])
    neutral_distribution = defaultdict(int); max_neutral_observed = 0
    for s in student_ids: count = student_neutral_counts[s]; neutral_distribution[count] += 1; max_neutral_observed = max(max_neutral_observed, count)
    for i in range(max_neutral_observed + 1): stats.append((f"Nb élèves avec {i} neutres", neutral_distribution[i]))
    stats_df = pd.DataFrame(stats, columns=["Statistique", "Valeur"])

else: # Handle case where no optimal solution was found
     stats = [("Statut Final Solveur", pulp.LpStatus[prob.status])]
     stats.append(("Aucune solution optimale trouvée.", "Stats détaillées non disponibles."))
     stats_df = pd.DataFrame(stats, columns=["Statistique", "Valeur"])

# ------------------------------
# WRITE OUTPUT
# ------------------------------
output_filename = "planning_final.xlsx" # Changed filename
print(f"Écriture du fichier de sortie '{output_filename}'...")
try:
    with pd.ExcelWriter(output_filename, engine="xlsxwriter") as writer:
        student_schedule_df.to_excel(writer, sheet_name="Planning par élève", index=False)
        # Write activity schedule: index=True (row labels), header=True (instance IDs)
        if not activity_schedule_df.empty:
            activity_schedule_df.to_excel(writer, sheet_name="Planning par Atelier", index=True, header=True)
        else:
             pd.DataFrame().to_excel(writer, sheet_name="Planning par Atelier") # Empty placeholder
        stats_df.to_excel(writer, sheet_name="Résumé et Stats", index=False)

    print(f"\nLe planning final a été écrit dans le fichier '{output_filename}'.")
    if not optimal_solution_found: print_solver_warning(prob.status, is_final_msg=True) # Use helper again
    else: print("Onglets créés : 'Planning par élève', 'Planning par Atelier', 'Résumé et Stats'.")

except Exception as e:
    print(f"\nERREUR lors de l'écriture du fichier Excel '{output_filename}': {e}")
    print("Vérifiez si le fichier est déjà ouvert ou si vous avez les permissions d'écriture.")

# Helper function for warnings (optional)
def print_solver_warning(status, is_final_msg=False):
    prefix = "ATTENTION:" if not is_final_msg else "ATTENTION: Le fichier de sortie a été généré, mais il peut être vide ou incomplet car "
    if status == pulp.LpStatusInfeasible:
        print(f"{prefix}Le modèle est infaisable.")
    elif status == pulp.LpStatusNotSolved:
         print(f"{prefix}Le solveur n'a pas démarré ou a été interrompu.")
    elif status == pulp.LpStatusUndefined:
         print(f"{prefix}La solution optimale n'a pu être trouvée (non bornée ou infaisable).")
    else: # Other non-optimal statuses
        print(f"{prefix}Le solveur n'a pas trouvé de solution optimale (Statut: {pulp.LpStatus[status]}).")