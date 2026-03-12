import os, sys, csv, json, shutil, tempfile, threading, urllib.request
from flask import (
    Flask, render_template, request,
    redirect, url_for, flash,
    send_file, make_response
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

# ── Update check ──────────────────────────────────────────────────────────────
APP_VERSION = "1.0.0"
# Replace this URL with your GitHub Gist raw URL after creating it (see README)
UPDATE_CHECK_URL = "https://gist.github.com/CrazyStill/4bfa488d53a8322ccaad43bff389f876"

_update_info = None   # None = still checking | False = up to date | dict = update available


def _parse_version(v):
    try:
        return tuple(int(x) for x in str(v).split('.'))
    except Exception:
        return (0, 0, 0)


def _fetch_update_info():
    global _update_info
    try:
        with urllib.request.urlopen(UPDATE_CHECK_URL, timeout=5) as resp:
            data = json.loads(resp.read().decode())
        if _parse_version(data.get("version", "0")) > _parse_version(APP_VERSION):
            _update_info = data
        else:
            _update_info = False
    except Exception:
        _update_info = False


threading.Thread(target=_fetch_update_info, daemon=True).start()


@app.context_processor
def inject_globals():
    return {'APP_VERSION': APP_VERSION}


def get_sctemp_dir():
    """Return AppData/Roaming/ScorecardCreator/SCTEMP when running as .exe,
    fall back to the local SCTEMP/ folder during development."""
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
    """On first run as packaged app, copy any bundled SCTEMP data into AppData
    if the AppData SCTEMP folder is empty."""
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


def get_downloads_dir():
    """Return the user's Downloads folder path."""
    return os.path.join(os.path.expanduser('~'), 'Downloads')


def save_to_downloads(src_path, filename):
    """Copy a file to the user's Downloads folder. Returns the destination path."""
    downloads = get_downloads_dir()
    os.makedirs(downloads, exist_ok=True)
    dst = os.path.join(downloads, filename)
    shutil.copy2(src_path, dst)
    return dst


def allowed_file(filename, allowed_extensions):
    return (
        '.' in filename
        and filename.rsplit('.', 1)[1].lower() in allowed_extensions
    )


@app.route('/')
def index():
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
    if request.method == 'POST':
        sport         = request.form.get('sport')
        template_name = request.form.get('template_name')
        if not sport or not template_name:
            flash("Sport and Template Name are required.")
            return redirect(request.url)

        sport          = secure_filename(sport)
        template_name  = secure_filename(template_name)
        template_dir   = os.path.join(SCTEMP_DIR, sport, template_name)
        os.makedirs(template_dir, exist_ok=True)

        front_file = request.files.get('front_file')
        if not front_file or not allowed_file(front_file.filename, {'docx'}):
            flash("A valid Word template file (.docx) is required.")
            return redirect(request.url)
        front_file.save(os.path.join(template_dir, 'template_front.docx'))

        csv_file = request.files.get('csv_file')
        if not csv_file or not allowed_file(csv_file.filename, {'csv'}):
            flash("A valid CSV file is required.")
            return redirect(request.url)
        csv_file.save(os.path.join(template_dir, 'template_data.csv'))

        if request.form.get('back_option') == 'yes':
            back_file = request.files.get('back_file')
            if back_file and allowed_file(back_file.filename, {'pdf'}):
                back_file.save(os.path.join(template_dir, 'template_back.pdf'))
            else:
                flash("Back design selected but no valid PDF uploaded.")
                return redirect(request.url)

        return redirect(url_for(
            'mapping',
            sport=sport,
            template_name=template_name
        ))

    return render_template('upload.html')


@app.route('/mapping/<sport>/<template_name>', methods=['GET','POST'])
def mapping(sport, template_name):
    template_dir = os.path.join(
        SCTEMP_DIR,
        secure_filename(sport),
        secure_filename(template_name)
    )
    csv_path     = os.path.join(template_dir, 'template_data.csv')
    mapping_file = os.path.join(template_dir, 'mapping.json')

    if not os.path.exists(csv_path):
        flash("CSV file not found for this template.")
        return redirect(url_for('index'))

    existing_cards_per_page = 4
    existing_mapping = {}
    if os.path.exists(mapping_file):
        data = json.load(open(mapping_file))
        existing_cards_per_page = data.get("cards_per_page", 4)
        existing_mapping = data.get("mapping", {})

    if request.method == 'POST':
        new_csv = request.files.get('new_csv')
        if new_csv and allowed_file(new_csv.filename, {'csv'}):
            new_csv.save(csv_path)
            flash("CSV file updated successfully.")

        with open(csv_path, newline='', encoding='latin-1') as f:
            sample = f.read(1024); f.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample)
            except csv.Error:
                dialect = csv.excel
            reader = csv.reader(f, dialect=dialect)
            headers = next(reader)

        try:
            cards_per_page = int(request.form.get('cards_per_page', 4))
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

        flash("Mapping saved successfully.")
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
    template_dir  = os.path.join(
        SCTEMP_DIR,
        secure_filename(sport),
        secure_filename(template_name)
    )
    front_docx    = os.path.join(template_dir, 'template_front.docx')
    if not os.path.exists(front_docx):
        flash("DOCX template not found.")
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
        preview_pdf_path = os.path.join(temp_dir, "preview.pdf")
        merge_two_pdfs(temp_front_pdf, back_pdf, preview_pdf_path)
        pdf_to_send = preview_pdf_path
    else:
        pdf_to_send = temp_front_pdf

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
            flash("Template updated successfully.")
            return redirect(url_for('preview', sport=sport, template_name=template_name))
        flash("Please upload a valid DOCX file.")
        return redirect(url_for('preview', sport=sport, template_name=template_name))

    return render_template('preview.html', sport=sport, template_name=template_name)


@app.route('/download_template/<sport>/<template_name>')
def download_template(sport, template_name):
    template_dir = os.path.join(
        SCTEMP_DIR,
        secure_filename(sport),
        secure_filename(template_name)
    )
    docx_path    = os.path.join(template_dir, 'template_front.docx')
    if not os.path.exists(docx_path):
        flash("DOCX template not found.")
        return redirect(url_for('index'))

    filename = f"{template_name}_template.docx"
    saved_path = save_to_downloads(docx_path, filename)
    flash(f"Template saved to Downloads: {filename}")
    return redirect(url_for('preview', sport=sport, template_name=template_name))


@app.route('/generate/<sport>/<template_name>', methods=['GET','POST'])
def generate(sport, template_name):
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
            flash("Please upload a valid CSV file.")
            return redirect(request.url)

        with tempfile.TemporaryDirectory() as temp_dir:
            filled_csv_path = os.path.join(temp_dir, 'filled_data.csv')
            filled_csv_file.save(filled_csv_path)

            mapping_json = json.load(open(mapping_file))
            cards_per_page = mapping_json.get("cards_per_page", 4)
            mapping_data   = mapping_json.get("mapping", {})

            output_pdf = generate_scorecard(
                front_path,
                filled_csv_path,
                mapping_data,
                cards_per_page=cards_per_page,
                back_pdf_path=back_path if os.path.exists(back_path) else None,
                temp_dir=temp_dir
            )

            filename = f"{template_name}_Scorecards.pdf"
            saved_path = save_to_downloads(output_pdf, filename)

        flash(f"Scorecards saved to Downloads folder: {filename}")
        return redirect(url_for('index'))

    return render_template('generate.html', sport=sport, template_name=template_name)


@app.route('/download_csv/<sport>/<template_name>')
def download_csv(sport, template_name):
    template_dir     = os.path.join(
        SCTEMP_DIR,
        secure_filename(sport),
        secure_filename(template_name)
    )
    csv_template_path = os.path.join(template_dir, 'template_data.csv')

    if not os.path.exists(csv_template_path):
        flash("CSV template not found.")
        return redirect(url_for('index'))

    filename = f"{template_name}_template.csv"
    save_to_downloads(csv_template_path, filename)
    flash(f"CSV template saved to Downloads: {filename}")
    return redirect(url_for('generate', sport=sport, template_name=template_name))


@app.route('/delete/<sport>/<template_name>', methods=['POST'])
def delete_template(sport, template_name):
    template_dir = os.path.join(
        SCTEMP_DIR,
        secure_filename(sport),
        secure_filename(template_name)
    )
    if os.path.exists(template_dir):
        shutil.rmtree(template_dir)
        flash("Template deleted successfully.")
    else:
        flash("Template not found.")
    return redirect(url_for('index'))


@app.route('/about')
def about():
    return render_template('about.html')


if __name__ == '__main__':
    app.run(debug=True, port=5000)
