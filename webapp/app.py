# app.py
import os
import uuid
import time
import json
import threading
import queue
import datetime
from flask import Flask, request, render_template, send_from_directory, redirect, url_for, flash, Response, jsonify
from werkzeug.utils import secure_filename
import traceback

# Import the solver function
from solver_logic import run_optimization

# --- Configuration ---
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
RESULT_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'results')
STATIC_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'static')
ALLOWED_EXTENSIONS = {'xlsx'}
TEMPLATE_FILENAME = 'Template.xlsx'

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('FLASK_SECRET_KEY', 'change_this_in_production_to_something_secret')
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['RESULT_FOLDER'] = RESULT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

# --- Job tracking for SSE ---
jobs = {}

# --- Helper Functions ---
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

def _friendly_error(message):
    """Map solver error messages to user-friendly French text."""
    if "Onglet" in message and "introuvable" in message:
        return message + " Assurez-vous d'utiliser le modèle fourni."
    if "infaisable" in message.lower() or "infeasible" in message.lower():
        return ("Impossible de trouver un planning valide. "
                "Causes possibles : trop de vetos, capacités insuffisantes, ou sessions manquantes dans les ateliers.")
    if "Colonnes manquantes" in message:
        return message + " Vérifiez que votre fichier suit le format du modèle."
    if "Aucun élève" in message:
        return "Aucun élève trouvé dans l'onglet Preferences. Vérifiez votre fichier."
    if "Aucune instance d'atelier" in message:
        return "Aucun atelier valide trouvé. Vérifiez l'onglet Ateliers et les colonnes Session."
    return message

def _run_solver_job(job_id, input_path, output_path, category_weight):
    """Run the solver in a background thread, pushing progress events to a queue."""
    job = jobs[job_id]
    q = job["queue"]

    def progress_callback(step, pct):
        q.put({"type": "progress", "step": step, "pct": pct})

    try:
        t_start = time.time()
        success, status_message, stats_summary = run_optimization(
            input_path, output_path,
            category_diversity_weight=category_weight,
            progress_callback=progress_callback
        )
        solve_time = round(time.time() - t_start, 1)

        # Clean up input file
        try:
            os.remove(input_path)
        except OSError:
            pass

        if success:
            q.put({
                "type": "complete",
                "success": True,
                "message": status_message,
                "filename": job["output_filename"],
                "stats": stats_summary,
                "solve_time": solve_time
            })
        else:
            # Clean up output file on failure
            if os.path.exists(output_path):
                try:
                    os.remove(output_path)
                except OSError:
                    pass
            q.put({
                "type": "complete",
                "success": False,
                "message": _friendly_error(status_message),
                "solve_time": solve_time
            })

    except Exception as e:
        print(f"Erreur inattendue dans le job {job_id}: {e}")
        print(traceback.format_exc())
        if os.path.exists(input_path):
            try:
                os.remove(input_path)
            except OSError:
                pass
        if os.path.exists(output_path):
            try:
                os.remove(output_path)
            except OSError:
                pass
        q.put({
            "type": "complete",
            "success": False,
            "message": "Une erreur serveur inattendue est survenue lors du traitement."
        })


# --- Routes ---
@app.route('/')
def index():
    """Renders the main upload page."""
    return render_template('index.html')

@app.route('/download_template')
def download_template():
    """Serves the template file."""
    try:
        return send_from_directory(STATIC_FOLDER, TEMPLATE_FILENAME, as_attachment=True)
    except FileNotFoundError:
        flash("ERREUR: Fichier template non trouvé sur le serveur.", "danger")
        return redirect(url_for('index'))

@app.route('/optimize', methods=['POST'])
def optimize():
    """Accepts file upload, starts solver in background thread, returns job_id."""
    if 'file' not in request.files:
        return jsonify({"error": "Aucun fichier sélectionné."}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "Aucun fichier sélectionné."}), 400

    if not (file and allowed_file(file.filename)):
        return jsonify({"error": "Type de fichier non autorisé. Utilisez un fichier .xlsx."}), 400

    unique_id = uuid.uuid4().hex
    job_id = unique_id[:12]
    input_filename = f"{unique_id}_input.xlsx"
    input_path = os.path.join(app.config['UPLOAD_FOLDER'], input_filename)

    now = datetime.datetime.now()
    timestamp = now.strftime("%Y-%m-%d_%H-%M-%S")
    human_readable_output_filename = f"Planning_{timestamp}.xlsx"
    output_path = os.path.join(app.config['RESULT_FOLDER'], human_readable_output_filename)

    try:
        file.save(input_path)
    except Exception as e:
        return jsonify({"error": f"Erreur lors de la sauvegarde du fichier: {e}"}), 500

    category_weight = request.form.get('category_weight', 0, type=float)

    # Create job entry
    jobs[job_id] = {
        "queue": queue.Queue(),
        "output_filename": human_readable_output_filename,
        "output_path": output_path,
        "created": time.time()
    }

    # Start solver in background thread
    thread = threading.Thread(
        target=_run_solver_job,
        args=(job_id, input_path, output_path, category_weight),
        daemon=True
    )
    thread.start()

    return jsonify({"job_id": job_id})

@app.route('/progress/<job_id>')
def progress(job_id):
    """SSE endpoint that streams solver progress events."""
    if job_id not in jobs:
        return jsonify({"error": "Job introuvable."}), 404

    def event_stream():
        q = jobs[job_id]["queue"]
        while True:
            try:
                event = q.get(timeout=300)
                event_type = event.get("type", "progress")
                data = json.dumps(event)
                yield f"event: {event_type}\ndata: {data}\n\n"
                if event_type == "complete":
                    # Clean up job after a delay (let client process the event)
                    def cleanup():
                        time.sleep(60)
                        jobs.pop(job_id, None)
                    threading.Thread(target=cleanup, daemon=True).start()
                    break
            except queue.Empty:
                # Send keepalive
                yield f": keepalive\n\n"

    return Response(event_stream(), mimetype='text/event-stream',
                    headers={'Cache-Control': 'no-cache', 'X-Accel-Buffering': 'no'})

@app.route('/download_result/<path:filename>')
def download_result(filename):
    """Serves the generated result file."""
    safe_filename = secure_filename(filename)
    if safe_filename != filename:
         flash("Nom de fichier invalide.", "danger")
         return redirect(url_for('index'))

    try:
        print(f"Tentative de téléchargement du résultat: {safe_filename}")
        return send_from_directory(app.config['RESULT_FOLDER'], safe_filename, as_attachment=True)
    except FileNotFoundError:
        print(f"ERREUR: Fichier résultat non trouvé pour téléchargement: {safe_filename}")
        flash("ERREUR: Fichier résultat introuvable. Il a peut-être expiré.", "danger")
        return redirect(url_for('index'))

# --- Main execution ---
if __name__ == '__main__':
    # Make sure debug=False for any production deployment!
    app.run(debug=True, host='0.0.0.0', port=5005)
