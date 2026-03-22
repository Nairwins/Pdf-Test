import sys, json, base64
sys.path.insert(0, '/home/claude')

from flask import Flask, request, render_template, jsonify
from fire import generate_resume_pdf

# ── DOCX import ───────────────────────────────────────────────────────────────
import os, importlib.util

def _load_docx_builder():
    """Load template/docx.py from the project tree."""
    candidates = [
        os.path.join(os.path.dirname(__file__), 'template', 'doc.py'),
        os.path.join(os.path.dirname(__file__), 'template', 'docx.py'),
        os.path.join(os.path.dirname(__file__), 'template_docx.py'),
    ]
    for path in candidates:
        if os.path.exists(path):
            spec = importlib.util.spec_from_file_location('docx_template', path)
            mod  = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            return mod.build_resume_docx
    raise ImportError('Could not find docx template — tried: ' + str(candidates))

build_resume_docx = _load_docx_builder()

app = Flask(__name__, template_folder='templates')


# ── PATH HELPERS ──────────────────────────────────────────────────────────────

def _get(data, path):
    """Get value from data using dot-path like 'experience.0.role'"""
    keys = path.split('.')
    cur = data
    for k in keys:
        if isinstance(cur, list):
            try: cur = cur[int(k)]
            except: return None
        elif isinstance(cur, dict):
            cur = cur.get(k)
        else:
            return None
    return cur


def _set(data, path, value):
    """Set value in data using dot-path. Handles __dates, __comp, __deg, __csv special paths."""
    keys = path.split('.')
    cur = data
    for k in keys[:-1]:
        if isinstance(cur, list):
            cur = cur[int(k)]
        elif isinstance(cur, dict):
            cur = cur.setdefault(k, {})
    last = keys[-1]

    if last == '__dates':
        parts = value.split('–')
        if len(parts) == 2:
            cur['start_date'] = parts[0].strip()
            right = parts[1].strip()
            if '(' in right:
                end, dur = right.split('(', 1)
                cur['end_date'] = end.strip()
                cur['duration'] = dur.rstrip(')').strip()
            else:
                cur['end_date'] = right.strip()
        return

    if last == '__comp':
        parts = value.split('·')
        cur['company'] = parts[0].strip()
        if len(parts) > 1:
            cur['location'] = parts[1].strip()
        return

    if last == '__deg':
        parts = value.split('—')
        cur['degree'] = parts[0].strip()
        if len(parts) > 1:
            cur['field'] = parts[1].strip()
        return

    if last == '__date':
        import re
        state_m = re.search(r'\[(.+?)\]', value)
        if state_m:
            cur['state'] = state_m.group(1)
            value = value[:state_m.start()].strip()
        parts = value.split('–')
        if len(parts) == 2:
            cur['start_date'] = parts[0].strip()
            cur['end_date']   = parts[1].strip()
        else:
            cur['end_date'] = value.strip()
        return

    if last == '__inst':
        parts = value.split('·')
        cur['institution'] = parts[0].strip()
        if len(parts) > 1:
            gpa_str = parts[1].strip()
            if gpa_str.startswith('GPA:'):
                cur['gpa'] = gpa_str[4:].strip()
        return

    if last == '__csv':
        parent_key = keys[-2]
        if parent_key == 'softskills':
            data['softskills'] = [v.strip() for v in value.split(',') if v.strip()]
        elif parent_key == 'languages':
            data['languages'] = [v.strip() for v in value.split(',') if v.strip()]
        elif isinstance(cur, dict):
            cur[parent_key] = [v.strip() for v in value.split(',') if v.strip()]
        return

    if isinstance(cur, list):
        cur[int(last)] = value
    elif isinstance(cur, dict):
        cur[last] = value


# ── ROUTES ────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/render-pdf', methods=['POST'])
def render_pdf():
    try:
        data        = request.get_json()
        resume_data = data.get('data', {})
        color       = data.get('color', '#1a56db')
        pdf_bytes, zones, pw, ph, num_pages = generate_resume_pdf(
            resume_data, main_color=color)
        b64  = base64.b64encode(pdf_bytes).decode()
        name = resume_data.get('name', 'resume').replace(' ', '_')
        return jsonify({
            'success':   True,
            'pdf_b64':   b64,
            'filename':  f'{name}_resume.pdf',
            'zones':     zones,
            'pdf_w':     pw,
            'pdf_h':     ph,
            'num_pages': num_pages,
        })
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/render-docx', methods=['POST'])
def render_docx():
    """Generate and return a DOCX file as base64."""
    try:
        data        = request.get_json()
        resume_data = data.get('data', {})
        color       = data.get('color', '#1a56db')
        docx_bytes  = build_resume_docx(resume_data, main_color=color)
        b64  = base64.b64encode(docx_bytes).decode()
        name = resume_data.get('name', 'resume').replace(' ', '_')
        return jsonify({
            'success':      True,
            'docx_b64':     b64,
            'filename':     f'{name}_resume.docx',
        })
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/edit-field', methods=['POST'])
def edit_field():
    """Apply a single field edit and re-render PDF."""
    try:
        body        = request.get_json()
        resume_data = body.get('data', {})
        path        = body.get('path', '')
        value       = body.get('value', '')
        color       = body.get('color', '#1a56db')

        _set(resume_data, path, value)

        pdf_bytes, zones, pw, ph, num_pages = generate_resume_pdf(
            resume_data, main_color=color)
        b64  = base64.b64encode(pdf_bytes).decode()
        name = resume_data.get('name', 'resume').replace(' ', '_')
        return jsonify({
            'success':   True,
            'pdf_b64':   b64,
            'filename':  f'{name}_resume.pdf',
            'zones':     zones,
            'pdf_w':     pw,
            'pdf_h':     ph,
            'num_pages': num_pages,
            'data':      resume_data,
        })
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/preview', methods=['POST'])
def preview():
    try:
        data        = request.get_json()
        resume_data = json.loads(data.get('json_text', ''))
        return jsonify({'success': True, 'data': resume_data})
    except json.JSONDecodeError as e:
        return jsonify({'success': False, 'error': f'Invalid JSON: {e}'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


if __name__ == '__main__':
    app.run(debug=True, port=5000)