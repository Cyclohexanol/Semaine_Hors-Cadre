# app.py
import os
import uuid
import datetime # Import datetime module
from flask import Flask, request, render_template, send_from_directory, redirect, url_for, flash
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

# --- Helper Functions ---
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

# --- Routes ---
@app.route('/')
def index():
    """Renders the main upload page."""
    return render_template('index.html', message=None, success=False, filename=None, stats=None)

@app.route('/download_template')
def download_template():
    """Serves the template file."""
    try:
        return send_from_directory(STATIC_FOLDER, TEMPLATE_FILENAME, as_attachment=True)
    except FileNotFoundError:
        flash("ERREUR: Fichier template non trouvé sur le serveur.", "danger")
        return redirect(url_for('index'))

@app.route('/upload', methods=['POST'])
def upload_file():
    """Handles file upload, runs optimization, returns results."""
    if 'file' not in request.files:
        flash('Aucun fichier sélectionné.', 'warning')
        return redirect(request.url)

    file = request.files['file']
    if file.filename == '':
        flash('Aucun fichier sélectionné.', 'warning')
        return redirect(request.url)

    if file and allowed_file(file.filename):
        original_filename = secure_filename(file.filename) # Keep original name just for reference if needed
        # Use a unique ID for the temporary input file to avoid clashes
        unique_id = uuid.uuid4().hex
        input_filename = f"{unique_id}_input.xlsx"
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], input_filename)

        # Define output filename generation here, BEFORE calling the solver
        now = datetime.datetime.now()
        timestamp = now.strftime("%Y-%m-%d_%H-%M-%S") # Format: YYYY-MM-DD_HH-MM-SS
        # --- CHANGE: Human-readable output filename ---
        human_readable_output_filename = f"Planning_{timestamp}.xlsx"
        output_filename_for_storage = f"{unique_id}_result.xlsx" # Still use UUID internally if needed for cleanup logic later
        # For simplicity now, we'll just use the human-readable one for saving and serving
        output_path = os.path.join(app.config['RESULT_FOLDER'], human_readable_output_filename)
        # ------------------------------------------------

        try:
            file.save(input_path)
            print(f"Fichier uploadé sauvegardé: {input_path}")

            # --- Run the optimization ---
            success, status_message, stats_summary = run_optimization(input_path, output_path)
            # -----------------------------

            # Always try to clean up the uploaded input file
            try:
                os.remove(input_path)
                print(f"Fichier uploadé supprimé: {input_path}")
            except OSError as e:
                print(f"Erreur non bloquante lors de la suppression du fichier uploadé {input_path}: {e}")

            if success:
                 # If successful, pass the human-readable filename for the download link
                 return render_template('index.html',
                                       message=status_message,
                                       success=True,
                                       filename=human_readable_output_filename, # Pass the correct name
                                       stats=stats_summary)
            else:
                # If failed, try to clean up the output file path (which might not have been created)
                if os.path.exists(output_path):
                    try:
                        os.remove(output_path)
                        print(f"Fichier résultat (échoué) supprimé: {output_path}")
                    except OSError as e:
                         print(f"Erreur non bloquante lors de la suppression du fichier résultat (échoué) {output_path}: {e}")
                return render_template('index.html',
                                       message=status_message,
                                       success=False,
                                       filename=None,
                                       stats=None)

        except Exception as e:
            print(f"Erreur inattendue dans la route /upload: {e}")
            print(traceback.format_exc())
            if os.path.exists(input_path): os.remove(input_path)
            if os.path.exists(output_path): os.remove(output_path)
            flash(f"Une erreur serveur critique est survenue lors du traitement.", "danger")
            return redirect(url_for('index'))

    else:
        flash('Type de fichier non autorisé. Utilisez un fichier .xlsx.', 'warning')
        return redirect(request.url)

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
    app.run(debug=True, host='0.0.0.0', port=5000)