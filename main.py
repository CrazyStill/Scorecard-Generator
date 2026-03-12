import os, sys, csv, json, shutil, tempfile, threading, queue, urllib.request
from flask import (
    Flask, render_template, request,
    redirect, url_for, flash,
    send_file, make_response, Response, stream_with_context
)
from werkzeug.utils import secure_filename
from io import BytesIO

from generate import (
    generate_scorecard,
    convert_docx_to_pdf,
    merge_two_pdfs
)

app = Flask(__name__)
app.secret_key = 'your_secret_key'
BASE_DIR = os.path.abspath(os.path.dirname(__file__))


# ── Auto-update check ─────────────────────────────────────────────────────────
# On startup a background thread fetches version.json from the GitHub Gist.
# If the remote version is newer than APP_VERSION the index page shows a modal.

APP_VERSION = "1.0.0"
UPDATE_CHECK_URL = (
    "https://gist.githubusercontent.com/CrazyStill/"
    "4bfa488d53a8322ccaad43bff389f876/raw/version.json"
)

_update_info = None   # None = still checking | False = up to date | dict = update available


def _parse_version(v):
    """Convert a version string like '1.2.3' to a comparable tuple."""
    try:
        return tuple(int(x) for x in str(v).split('.'))
    except Exception:
        return (0, 0, 0)


def _fetch_update_info():
    """Background thread: fetch version.json and populate _update_info."""
    global _update_info
    try:
        with urllib.request.urlopen(UPDATE_CHECK_URL, timeout=5) as resp:
            data = json.loads(resp.read().decode())
        if _parse_version(data.get("version", "0")) > _parse_version(APP_VERSION):
            _update_info = data       # dict with at least {"version": "x.y.z", "url": "..."}
        else:
            _update_info = False      # already up to date
    except Exception:
        _update_info = False          # network unavailable or Gist unreachable — fail silently


threading.Thread(target=_fetch_update_info, daemon=True).start()


@app.context_processor
def inject_globals():
    """Make APP_VERSION available in every template without explicit passing."""
    return {'APP_VERSION': APP_VERSION}


# ── Data directory ────────────────────────────────────────────────────────────

def get_sctemp_dir():
    """Return the path to the SCTEMP templates folder.

    Packaged .exe:  %APPDATA%\\ScorecardCreator\\SCTEMP\\
    Development:    <project root>\\SCTEMP\\
    """
    if getattr(sys, 'frozen', False):
        appdata = os.environ.get('APPDATA', os.path.expanduser('~'))
        base = os.path.join(appdata, 'ScorecardCreator')
    else:
        base = BASE_DIR
    sctemp = os.path.join(base, 'SCTEMP')
    os.makedirs(sctemp, exist_ok=True)
    return sctemp


SCTEMP_DIR = get_sctemp_dir()


def migrate_existing_templates():
    """Copy bundled templates into AppData on the first run of a packaged build.

    The installer bundles an SCTEMP folder next to the .exe. If AppData/SCTEMP
    is empty we copy those defaults in so users have templates ready to go.
    Runs only once — subsequent launches find AppData/SCTEMP non-empty.
    """
    if not getattr(sys, 'frozen', False):
        return
    install_dir = os.path.dirname(sys.executable)
    bundled_sctemp = os.path.join(install_dir, 'SCTEMP')
    if os.path.exists(bundled_sctemp) and not os.listdir(SCTEMP_DIR):
        for item in os.listdir(bundled_sctemp):
            src = os.path.join(bundled_sctemp, item)
            dst = os.path.join(SCTEMP_DIR, item)
            if os.path.isdir(src):
                shutil.copytree(src, dst)
            else:
                shutil.copy2(src, dst)


migrate_existing_templates()


# ── Downloads helpers ─────────────────────────────────────────────────────────

def get_downloads_dir():
    """Return the user's Downloads folder path.

    Reads the path from the Windows registry Shell Folders key so it works
    correctly when the user has redirected Downloads to another drive or
    OneDrive. Falls back to ~/Downloads if the registry lookup fails.
    """
    try:
        import winreg
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
        ) as key:
            # GUID {374DE290...} is the well-known Downloads folder identifier
            return winreg.QueryValueEx(key, '{374DE290-123F-4565-9164-39C4925E467B}')[0]
    except Exception:
        return os.path.join(os.path.expanduser('~'), 'Downloads')


def save_to_downloads(src_path, filename):
    """Copy src_path into the Downloads folder as filename.

    Uses shutil.copy (not copy2) so the destination gets the current
    timestamp — files then sort to the top in a date-ordered Downloads view.

    Returns the full destination path.
    """
    downloads = get_downloads_dir()
    os.makedirs(downloads, exist_ok=True)
    dst = os.path.join(downloads, filename)
    shutil.copy(src_path, dst)
    return dst


def allowed_file(filename, allowed_extensions):
    """Return True if filename has one of the allowed extensions."""
    return (
        '.' in filename
        and filename.rsplit('.', 1)[1].lower() in allowed_extensions
    )


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    """Home page: list all sport/template cards found in SCTEMP."""
    templates_list = []
    for sport in os.listdir(SCTEMP_DIR):
        sport_dir = os.path.join(SCTEMP_DIR, sport)
        if os.path.isdir(sport_dir):
            for template_name in os.listdir(sport_dir):
                template_dir = os.path.join(sport_dir, template_name)
                if os.path.isdir(template_dir):
                    templates_list.append({
                        'sport': sport,
                        'template': template_name
                    })
    update_info = _update_info if isinstance(_update_info, dict) else None
    return render_template('index.html', templates=templates_list, update_info=update_info)


@app.route('/upload', methods=['GET', 'POST'])
def upload():
    """Upload a new template: DOCX front, CSV data, and optional PDF back."""
    if request.method == 'POST':
        sport         = request.form.get('sport')
        template_name = request.form.get('template_name')
        if not sport or not template_name:
            flash("Sport and Template Name are required.", "danger")
            return redirect(request.url)

        sport          = secure_filename(sport)
        template_name  = secure_filename(template_name)
        template_dir   = os.path.join(SCTEMP_DIR, sport, template_name)
        os.makedirs(template_dir, exist_ok=True)

        front_file = request.files.get('front_file')
        if not front_file or not allowed_file(front_file.filename, {'docx'}):
            flash("A valid Word template file (.docx) is required.", "danger")
            return redirect(request.url)
        front_file.save(os.path.join(template_dir, 'template_front.docx'))

        csv_file = request.files.get('csv_file')
        if not csv_file or not allowed_file(csv_file.filename, {'csv'}):
            flash("A valid CSV file is required.", "danger")
            return redirect(request.url)
        csv_file.save(os.path.join(template_dir, 'template_data.csv'))

        if request.form.get('back_option') == 'yes':
            back_file = request.files.get('back_file')
            if back_file and allowed_file(back_file.filename, {'pdf'}):
                back_file.save(os.path.join(template_dir, 'template_back.pdf'))
            else:
                flash("Back design selected but no valid PDF uploaded.", "danger")
                return redirect(request.url)

        return redirect(url_for(
            'mapping',
            sport=sport,
            template_name=template_name
        ))

    return render_template('upload.html')


@app.route('/mapping/<sport>/<template_name>', methods=['GET','POST'])
def mapping(sport, template_name):
    """Configure the CSV-column → placeholder mapping for a template."""
    template_dir = os.path.join(
        SCTEMP_DIR,
        secure_filename(sport),
        secure_filename(template_name)
    )
    csv_path     = os.path.join(template_dir, 'template_data.csv')
    mapping_file = os.path.join(template_dir, 'mapping.json')

    if not os.path.exists(csv_path):
        flash("CSV file not found for this template.", "danger")
        return redirect(url_for('index'))

    # Load any previously saved mapping so the form pre-fills on re-visit
    existing_cards_per_page = 4
    existing_mapping = {}
    if os.path.exists(mapping_file):
        with open(mapping_file) as f:
            data = json.load(f)
        existing_cards_per_page = data.get("cards_per_page", 4)
        existing_mapping = data.get("mapping", {})

    if request.method == 'POST':
        # Optional CSV replacement: user can swap the data file at any time
        new_csv = request.files.get('new_csv')
        if new_csv and allowed_file(new_csv.filename, {'csv'}):
            new_csv.save(csv_path)
            flash("CSV file updated successfully.", "success")

        with open(csv_path, newline='', encoding='latin-1') as f:
            sample = f.read(1024); f.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample)
            except csv.Error:
                dialect = csv.excel
            reader = csv.reader(f, dialect=dialect)
            headers = next(reader)

        try:
            cards_per_page = max(1, min(4, int(request.form.get('cards_per_page', 4))))
        except ValueError:
            cards_per_page = 4

        new_mapping = {}
        for h in headers:
            new_mapping[h] = request.form.get(f"mapping_{h}", h)

        with open(mapping_file, 'w') as f:
            json.dump({
                "cards_per_page": cards_per_page,
                "mapping": new_mapping
            }, f)

        flash("Mapping saved successfully.", "success")
        return redirect(url_for(
            'mapping',
            sport=sport,
            template_name=template_name
        ))

    with open(csv_path, newline='', encoding='latin-1') as f:
        sample = f.read(1024); f.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample)
        except csv.Error:
            dialect = csv.excel
        reader = csv.reader(f, dialect=dialect)
        headers = next(reader)

    instructions = (
        "For each phrase to be replaced on the template, "
        "append an underscore and a number corresponding to "
        "the scorecard position on the page (e.g. _1, _2, _3...)."
    )
    return render_template(
        'mapping.html',
        sport=sport,
        template_name=template_name,
        instructions=instructions,
        headers=headers,
        existing_cards_per_page=existing_cards_per_page,
        existing_mapping=existing_mapping
    )


@app.route('/preview_pdf/<sport>/<template_name>')
def preview_pdf(sport, template_name):
    """Convert the template DOCX to PDF and serve it inline for the preview embed."""
    template_dir  = os.path.join(
        SCTEMP_DIR,
        secure_filename(sport),
        secure_filename(template_name)
    )
    front_docx    = os.path.join(template_dir, 'template_front.docx')
    if not os.path.exists(front_docx):
        flash("DOCX template not found.", "danger")
        return redirect(url_for('index'))

    temp_dir       = tempfile.mkdtemp()
    temp_front_pdf = os.path.join(temp_dir, "temp_front.pdf")
    try:
        convert_docx_to_pdf(front_docx, temp_front_pdf)
    except Exception as e:
        shutil.rmtree(temp_dir)
        return f"Error converting DOCX to PDF: {e}", 500

    back_pdf = os.path.join(template_dir, 'template_back.pdf')
    if os.path.exists(back_pdf):
        # Merge front + back so the preview shows both sides
        preview_pdf_path = os.path.join(temp_dir, "preview.pdf")
        merge_two_pdfs(temp_front_pdf, back_pdf, preview_pdf_path)
        pdf_to_send = preview_pdf_path
    else:
        pdf_to_send = temp_front_pdf

    # Read into memory before cleaning up temp files
    with open(pdf_to_send, 'rb') as f:
        pdf_bytes = f.read()
    shutil.rmtree(temp_dir)

    return send_file(
        BytesIO(pdf_bytes),
        mimetype='application/pdf',
        download_name='preview.pdf',
        as_attachment=False
    )


@app.route('/preview/<sport>/<template_name>', methods=['GET','POST'])
def preview(sport, template_name):
    """Preview page for a template. POST replaces the DOCX with a new upload."""
    template_dir = os.path.join(
        SCTEMP_DIR,
        secure_filename(sport),
        secure_filename(template_name)
    )
    docx_path    = os.path.join(template_dir, 'template_front.docx')

    if request.method == 'POST':
        new_docx = request.files.get('new_docx')
        if new_docx and allowed_file(new_docx.filename, {'docx'}):
            new_docx.save(docx_path)
            flash("Template updated successfully.", "success")
            return redirect(url_for('preview', sport=sport, template_name=template_name))
        flash("Please upload a valid DOCX file.", "danger")
        return redirect(url_for('preview', sport=sport, template_name=template_name))

    return render_template('preview.html', sport=sport, template_name=template_name)


@app.route('/download_template/<sport>/<template_name>')
def download_template(sport, template_name):
    """Save the DOCX template to the user's Downloads folder."""
    template_dir = os.path.join(
        SCTEMP_DIR,
        secure_filename(sport),
        secure_filename(template_name)
    )
    docx_path    = os.path.join(template_dir, 'template_front.docx')
    if not os.path.exists(docx_path):
        flash("DOCX template not found.", "danger")
        return redirect(url_for('index'))

    filename = f"{template_name}_template.docx"
    try:
        saved_path = save_to_downloads(docx_path, filename)
        flash(f"Saved to: {saved_path}", "success")
    except Exception as e:
        flash(f"Could not save file: {e}", "danger")
    return redirect(url_for('preview', sport=sport, template_name=template_name))


@app.route('/generate/<sport>/<template_name>', methods=['GET','POST'])
def generate(sport, template_name):
    """Generate scorecard PDFs from a filled CSV.

    GET:  Render the generate page with upload form.
    POST: Accept the CSV, run generation in a background thread, and stream
          progress back to the browser as Server-Sent Events (SSE). The
          frontend reads the stream and updates a Bootstrap progress bar.

    SSE event format (newline-delimited JSON):
        data: {"step": 2, "total": 5, "msg": "Converting page 2 of 5…"}
        data: {"done": true, "path": "C:\\Users\\...\\Downloads\\...pdf"}
        data: {"error": "Something went wrong"}
    """
    template_dir  = os.path.join(
        SCTEMP_DIR,
        secure_filename(sport),
        secure_filename(template_name)
    )
    front_path    = os.path.join(template_dir, 'template_front.docx')
    mapping_file  = os.path.join(template_dir, 'mapping.json')
    back_path     = os.path.join(template_dir, 'template_back.pdf')

    if request.method == 'POST':
        filled_csv_file = request.files.get('filled_csv')
        if not filled_csv_file or not allowed_file(filled_csv_file.filename, {'csv'}):
            # Return an immediate SSE error event so the JS can display it
            return Response(
                'data: ' + json.dumps({'error': 'Please upload a valid CSV file.'}) + '\n\n',
                mimetype='text/event-stream'
            )

        # Save the uploaded CSV before the streaming response starts — the file
        # object is only valid during this request, but generation runs in a
        # separate thread that outlives the upload handler.
        tmp_dir = tempfile.mkdtemp()
        filled_csv_path = os.path.join(tmp_dir, 'filled_data.csv')
        filled_csv_file.save(filled_csv_path)

        with open(mapping_file) as f:
            mapping_json = json.load(f)
        cards_per_page = max(1, min(4, int(mapping_json.get('cards_per_page', 4))))
        mapping_data   = mapping_json.get('mapping', {})

        # Queue used to pass progress events from the generation thread back to
        # the SSE streaming generator running on the main request thread.
        progress_q = queue.Queue()

        def run_generation():
            """Background thread: run the full generation pipeline."""
            try:
                def cb(step, total, msg):
                    progress_q.put(('progress', step, total, msg))

                output_pdf = generate_scorecard(
                    front_path,
                    filled_csv_path,
                    mapping_data,
                    cards_per_page=cards_per_page,
                    back_pdf_path=back_path if os.path.exists(back_path) else None,
                    temp_dir=tmp_dir,
                    progress_callback=cb,
                )
                # Signal the save step before writing to Downloads
                progress_q.put(('progress', None, None, 'Saving to Downloads…'))
                filename = f'{template_name}_Scorecards.pdf'
                saved_path = save_to_downloads(output_pdf, filename)
                progress_q.put(('done', saved_path))
            except ValueError as e:
                progress_q.put(('error', str(e)))
            except Exception as e:
                progress_q.put(('error', f'Generation failed: {e}'))
            finally:
                # Always clean up temp files regardless of success or failure
                shutil.rmtree(tmp_dir, ignore_errors=True)

        threading.Thread(target=run_generation, daemon=True).start()

        def stream():
            """Generator that yields SSE events until generation finishes."""
            while True:
                item = progress_q.get()   # blocks until the thread posts something
                if item[0] == 'progress':
                    yield 'data: ' + json.dumps({
                        'step': item[1], 'total': item[2], 'msg': item[3]
                    }) + '\n\n'
                elif item[0] == 'done':
                    yield 'data: ' + json.dumps({'done': True, 'path': item[1]}) + '\n\n'
                    break
                elif item[0] == 'error':
                    yield 'data: ' + json.dumps({'error': item[1]}) + '\n\n'
                    break

        return Response(stream_with_context(stream()), mimetype='text/event-stream')

    return render_template('generate.html', sport=sport, template_name=template_name)


@app.route('/download_csv/<sport>/<template_name>')
def download_csv(sport, template_name):
    """Save the blank CSV data template to the user's Downloads folder."""
    template_dir     = os.path.join(
        SCTEMP_DIR,
        secure_filename(sport),
        secure_filename(template_name)
    )
    csv_template_path = os.path.join(template_dir, 'template_data.csv')

    if not os.path.exists(csv_template_path):
        flash("CSV template not found.", "danger")
        return redirect(url_for('index'))

    filename = f"{template_name}_template.csv"
    try:
        saved_path = save_to_downloads(csv_template_path, filename)
        flash(f"Saved to: {saved_path}", "success")
    except Exception as e:
        flash(f"Could not save file: {e}", "danger")
    return redirect(url_for('generate', sport=sport, template_name=template_name))


@app.route('/delete/<sport>/<template_name>', methods=['POST'])
def delete_template(sport, template_name):
    """Permanently remove a sport/template folder from SCTEMP."""
    template_dir = os.path.join(
        SCTEMP_DIR,
        secure_filename(sport),
        secure_filename(template_name)
    )
    if os.path.exists(template_dir):
        shutil.rmtree(template_dir)
        flash("Template deleted successfully.", "success")
    else:
        flash("Template not found.", "danger")
    return redirect(url_for('index'))


@app.route('/about')
def about():
    return render_template('about.html')


if __name__ == '__main__':
    app.run(debug=True, port=5000)
