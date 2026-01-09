import os
import threading
import uuid
from flask import Flask, render_template, request, send_file, redirect, url_for, flash, jsonify
from werkzeug.utils import secure_filename

UPLOAD_FOLDER = 'uploads'
RESULT_FOLDER = 'results'
LOG_FOLDER = 'logs'
ALLOWED_EXTENSIONS = {'ifc'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['RESULT_FOLDER'] = RESULT_FOLDER
app.config['LOG_FOLDER'] = LOG_FOLDER
app.secret_key = 'your_secret_key'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)
os.makedirs(LOG_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def run_script(ifc_path, result_path, log_path):
    import subprocess
    print(f"DEBUG: Running script with ifc_path={ifc_path}, result_path={result_path}, log_path={log_path}")
    print(f"DEBUG: Current working directory: {os.getcwd()}")
    with open(log_path, 'w', encoding='utf-8') as log_file:
        try:
            process = subprocess.Popen(
                ['python3', 'nen2580_inhoud_excel.py', ifc_path, '-o', result_path],
                stdout=log_file,
                stderr=subprocess.STDOUT
            )
            process.wait()
            print("DEBUG: Subprocess finished.")
        except Exception as e:
            print(f"DEBUG: Subprocess failed with error: {e}")

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files.get('ifcfile')
        if not file or file.filename == '':
            flash('No file selected')
            return redirect(request.url)
        if not allowed_file(file.filename):
            flash('Invalid file type')
            return redirect(request.url)
        filename = secure_filename(file.filename)
        unique_id = str(uuid.uuid4())
        upload_path = os.path.join(app.config['UPLOAD_FOLDER'], unique_id + "_" + filename)
        result_filename = unique_id + "_result.xlsx"
        result_path = os.path.join(app.config['RESULT_FOLDER'], result_filename)
        log_path = os.path.join(app.config['LOG_FOLDER'], unique_id + ".log")
        file.save(upload_path)

        # Start processing in a background thread
        thread = threading.Thread(target=run_script, args=(upload_path, result_path, log_path))
        thread.start()

        return redirect(url_for('progress', task_id=unique_id, result_filename=result_filename))
    return render_template('index.html')

@app.route('/progress/<task_id>/<result_filename>')
def progress(task_id, result_filename):
    return render_template('progress.html', task_id=task_id, result_filename=result_filename)

@app.route('/logs/<task_id>')
def logs(task_id):
    log_path = os.path.join(app.config['LOG_FOLDER'], task_id + ".log")
    result_path = None
    # Check if result file exists
    for fname in os.listdir(app.config['RESULT_FOLDER']):
        if fname.startswith(task_id) and fname.endswith('.xlsx'):
            result_path = fname
            break
    logs = ""
    if os.path.exists(log_path):
        with open(log_path, 'r', encoding='utf-8') as f:
            logs = f.read()
    done = bool(result_path)
    return jsonify({'logs': logs, 'done': done, 'result_filename': result_path})

@app.route('/test-subprocess')
def test_subprocess():
    test_file = os.path.join(app.config['RESULT_FOLDER'], 'test_subprocess.txt')
    import subprocess
    with open(test_file, 'w', encoding='utf-8') as f:
        f.write('Subprocess test file.\n')
    return f"Test file written to {test_file}"

@app.route('/test-subprocess-run')
def test_subprocess_run():
    test_file = os.path.join(app.config['RESULT_FOLDER'], 'test_subprocess_run.txt')
    import subprocess
    with open(test_file, 'w', encoding='utf-8') as log_file:
        process = subprocess.Popen(
            ['echo', 'Hello from subprocess!'],
            stdout=log_file,
            stderr=subprocess.STDOUT
        )
        process.wait()
    return f"Subprocess wrote to {test_file}"

@app.route('/download/<filename>')
def download_file(filename):
    result_path = os.path.join(app.config['RESULT_FOLDER'], filename)
    if os.path.exists(result_path):
        return send_file(result_path, as_attachment=True)
    else:
        flash('Result file not found.')
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)

# The /view/<filename> route is removed because it used pandas to read Excel files.