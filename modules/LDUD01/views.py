import io
import json as _json
import mimetypes
import os
from flask import Blueprint, render_template, request, jsonify, session, redirect, url_for, send_file
from functools import wraps
from . import model
from database import get_user_permissions, get_module_config, get_db, get_cursor
from mail_service import (
    queue_mail as _queue_mail,
    trigger_mail_processing as _trigger_mail_processing,
    build_approval_mail_html as _build_approval_mail_html,
)

bp = Blueprint('LDUD01', __name__, template_folder='.')
MODULE_CODE = 'LDUD01'

def _get_user_email_by_id(user_id):
    """Return (email, username) for a user_id."""
    if not user_id:
        return None, None
    from database import get_db, get_cursor
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT email, username FROM users WHERE id=%s', [user_id])
    row = cur.fetchone()
    conn.close()
    return (row['email'], row['username']) if row else (None, None)

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated


# ── Billed lock ──────────────────────────────────────────────────────────────
# Once a vessel's cargo is billed, its LDUD is frozen: the parcel quantities are
# what the bill was computed from, so editing them silently desyncs the bill.
# Two ways out, both outside LDUD01: cancel the invoice behind the bill (which
# voids the ledger rows), or unmark a cutover flag in Admin > Cutover.

def _ldud_vcn_id(ldud_id):
    if not ldud_id:
        return None
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT vcn_id FROM ldud_header WHERE id=%s', [ldud_id])
    row = cur.fetchone()
    conn.close()
    return row['vcn_id'] if row else None


def _ldud_id_for_op(op_id):
    if not op_id:
        return None
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT ldud_id FROM ldud_parcel_ops WHERE id=%s', [op_id])
    row = cur.fetchone()
    conn.close()
    return row['ldud_id'] if row else None


def _billed_locked(ldud_id):
    """Return a 409 response when this LDUD's vessel is billed, else None."""
    from modules.FIN01 import model as fin_model
    vcn_id = _ldud_vcn_id(ldud_id)
    if not vcn_id or not fin_model.is_vcn_billed(vcn_id):
        return None
    blockers = fin_model.vcn_billing_blockers(vcn_id)
    parts = []
    if blockers['bill_numbers']:
        parts.append('billed on ' + ', '.join(blockers['bill_numbers']))
    if blockers['cutover']:
        parts.append('cutover-flagged at go-live')
    return jsonify({'error':
        'This vessel is ' + ' and '.join(parts or ['billed']) + '. '
        'Cancel the invoice behind the bill, or unmark it in Admin › Cutover, '
        'before editing this LDUD.'}), 409

def get_perms():
    if session.get('is_admin'):
        return {'can_read': 1, 'can_add': 1, 'can_edit': 1, 'can_delete': 1}
    return get_user_permissions(session.get('user_id'), MODULE_CODE)

@bp.route('/module/LDUD01/')
@login_required
def view():
    perms = get_perms()
    if not perms.get('can_read'):
        return render_template('no_access.html'), 403
    return render_template('ldud01.html', permissions=perms)

@bp.route('/api/module/LDUD01/data')
@login_required
def get_data():
    try:
        page = int(request.args.get('page', 1))
        size = int(request.args.get('size', 20))
    except (ValueError, TypeError):
        page, size = 1, 20
    try:
        filters = _json.loads(request.args.get('filters', '[]'))
    except _json.JSONDecodeError:
        filters = []
    rows, total = model.get_data(page, size, filters)
    return jsonify({'data': rows, 'last_page': (total + size - 1) // size, 'total': total})

@bp.route('/api/module/LDUD01/vcn_list')
@login_required
def get_vcn_list():
    return jsonify(model.get_vcn_list())

@bp.route('/api/module/LDUD01/parcels-completion/<int:ldud_id>')
@login_required
def parcels_completion(ldud_id):
    return jsonify(model.parcels_completion(ldud_id))

@bp.route('/api/module/LDUD01/vcn_list/export')
@login_required
def get_export_vcn_list():
    return jsonify(model.get_vcn_list())

@bp.route('/api/module/LDUD01/save', methods=['POST'])
@login_required
def save():
    perms = get_perms()
    data = request.json
    is_new = not data.get('id')
    if is_new and not perms.get('can_add'):
        return jsonify({'error': 'No permission to add'}), 403
    if not is_new and not perms.get('can_edit'):
        return jsonify({'error': 'No permission to edit'}), 403

    config = get_module_config('LDUD01')
    is_approver = str(config.get('approver_id', '')) == str(session.get('user_id')) or session.get('is_admin')

    if not is_new:
        locked = _billed_locked(data['id'])
        if locked:
            return locked
        current_status = model.get_doc_status(data['id'])
        if current_status == 'Closed':
            if not is_approver:
                return jsonify({'error': 'Cannot edit a closed record'}), 403
            data['doc_status'] = 'Closed'
        elif current_status == 'Partial Close':
            data['doc_status'] = 'Partial Close'
        else:
            data['doc_status'] = 'Draft'
    else:
        data['doc_status'] = 'Draft'

    # Post-departure SOF times persist only once every LUEU01 parcel is completed
    # (actual End set). The grid also blocks them, but enforce here so the async
    # UI check can't be out-raced.
    gated = ['cast_off_datetime', 'pilot_board_departure', 'pilot_disembarked']
    if not is_new and any(g in data for g in gated):
        if not model.parcels_completion(data['id']).get('all_done'):
            for g in gated:
                data.pop(g, None)

    try:
        row_id, doc_num = model.save_header(data)
    except ValueError as e:
        return jsonify({'error': str(e)}), 400
    return jsonify({'id': row_id, 'doc_num': doc_num, 'doc_status': data.get('doc_status', 'Draft')})


@bp.route('/api/module/LDUD01/closure_check/<int:ldud_id>')
@login_required
def closure_check(ldud_id):
    return jsonify(model.get_closure_eligibility(ldud_id))


@bp.route('/api/module/LDUD01/close', methods=['POST'])
@login_required
def close():
    config = get_module_config('LDUD01')
    is_approver = str(config.get('approver_id', '')) == str(session.get('user_id')) or session.get('is_admin')
    if not is_approver:
        return jsonify({'error': 'No permission to close'}), 403
    data = request.json
    record_id = data.get('id')
    close_type = data.get('close_type')
    password = (data.get('password') or '').strip()
    if not record_id:
        return jsonify({'error': 'Missing id'}), 400
    if close_type not in ['Closed', 'Partial Close']:
        return jsonify({'error': 'Invalid close type'}), 400
    if not password:
        return jsonify({'error': 'Password is required'}), 400

    # Verify password server-side
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT id FROM users WHERE id=%s AND password=%s', [session.get('user_id'), password])
    user = cur.fetchone()
    conn.close()
    if not user:
        return jsonify({'error': 'Incorrect password'}), 403

    # Re-verify eligibility server-side
    eligibility = model.get_closure_eligibility(record_id)
    if not eligibility['eligible']:
        return jsonify({'error': 'Record not eligible for closure', 'missing': eligibility['missing']}), 400
    if close_type == 'Closed' and not eligibility['can_full_close']:
        return jsonify({'error': f"Operations total ({eligibility['ops_total']}) does not match BL total ({eligibility['bl_total']}) — use Partial Close instead"}), 400

    # Enforce: at least one Proof of Quantity document must be uploaded
    conn_doc = get_db()
    cur_doc = get_cursor(conn_doc)
    cur_doc.execute('SELECT COUNT(*) FROM ldud_proof_documents WHERE ldud_id=%s', [record_id])
    doc_count = cur_doc.fetchone()['count']
    conn_doc.close()
    if doc_count == 0:
        return jsonify({'error': 'At least one Proof of Quantity document must be uploaded before closing'}), 400

    model.close_record(record_id, close_type, session.get('username'))
    # Queue notification to approver
    try:
        cfg = get_module_config('LDUD01')
        approver_email, approver_name = _get_user_email_by_id(cfg.get('approver_id'))
        if approver_email:
            # Fetch doc_num and vessel name for the notification
            _conn = get_db()
            _cur = get_cursor(_conn)
            _cur.execute(
                'SELECT lh.doc_num, vh.vessel_name FROM ldud_header lh LEFT JOIN vcn_header vh ON vh.id = lh.vcn_id WHERE lh.id=%s',
                [record_id]
            )
            _row = _cur.fetchone()
            _conn.close()
            doc_num = _row['doc_num'] if _row else f'#{record_id}'
            vessel_name = (_row['vessel_name'] or '—') if _row else '—'
            badge_color = '#059669' if close_type == 'Closed' else '#d97706'
            ldud_url = request.host_url.rstrip('/') + f'/module/LDUD01/'
            _queue_mail(
                to_email=approver_email,
                to_name=approver_name,
                subject=f"[Portbird DPPL] LDUD {doc_num} — {close_type}",
                body_html=_build_approval_mail_html(
                    approver_name=approver_name,
                    action_label=close_type,
                    subtitle='Lay / Despatch — Closure Notification',
                    details=[
                        ('Document No', doc_num),
                        ('Vessel',      vessel_name),
                        ('Status',      close_type),
                    ],
                    action_url=ldud_url,
                    action_btn_label='View in Portbird',
                    submitted_by=session.get('username'),
                    badge_color=badge_color,
                ),
                module_code='LDUD01',
                ref_id=record_id,
            )
            _trigger_mail_processing()
    except Exception:
        pass
    return jsonify({'doc_status': close_type})


@bp.route('/api/module/LDUD01/closure-log/<int:record_id>')
@login_required
def closure_log(record_id):
    return jsonify(model.get_closure_log(record_id))

@bp.route('/api/module/LDUD01/delete', methods=['POST'])
@login_required
def delete():
    perms = get_perms()
    if not perms.get('can_delete'):
        return jsonify({'error': 'No permission to delete'}), 403
    locked = _billed_locked(request.json['id'])
    if locked:
        return locked
    model.delete_header(request.json['id'])
    return jsonify({'success': True})

# Parcel Operations sub-table endpoints
@bp.route('/api/module/LDUD01/parcel_ops/<int:ldud_id>')
@login_required
def get_parcel_ops(ldud_id):
    return jsonify(model.get_parcel_ops(ldud_id))

@bp.route('/api/module/LDUD01/parcel_ops/save', methods=['POST'])
@login_required
def save_parcel_op():
    perms = get_perms()
    if not perms.get('can_add') and not perms.get('can_edit'):
        return jsonify({'error': 'No permission'}), 403
    locked = _billed_locked(request.json.get('ldud_id'))
    if locked:
        return locked
    try:
        row_id = model.save_parcel_op(request.json)
    except ValueError as e:
        return jsonify({'error': str(e)}), 400
    return jsonify({'id': row_id, 'success': True})

@bp.route('/api/module/LDUD01/parcel_ops/delete', methods=['POST'])
@login_required
def delete_parcel_op():
    perms = get_perms()
    # sub-table rows are deletable by anyone who can edit/add (not gated on can_delete)
    if not perms.get('can_add') and not perms.get('can_edit'):
        return jsonify({'error': 'No permission'}), 403
    locked = _billed_locked(_ldud_id_for_op(request.json['id']))
    if locked:
        return locked
    model.delete_parcel_op(request.json['id'])
    return jsonify({'success': True})



# ── Proof of Quantity Documents (stored in DB as BYTEA) ──────────────────────

ALLOWED_EXTENSIONS = {'.pdf', '.jpg', '.jpeg', '.png', '.xlsx', '.xls', '.csv', '.doc', '.docx'}


@bp.route('/api/module/LDUD01/proof_docs/upload', methods=['POST'])
@login_required
def upload_proof_docs():
    ldud_id = request.form.get('ldud_id')
    if not ldud_id:
        return jsonify({'error': 'Missing ldud_id'}), 400

    files = request.files.getlist('files')
    if not files:
        return jsonify({'error': 'No files provided'}), 400

    conn = get_db()
    cur = get_cursor(conn)
    saved = []
    for f in files:
        original = f.filename or ''
        ext = os.path.splitext(original)[1].lower()
        if ext not in ALLOWED_EXTENSIONS:
            continue
        file_bytes = f.read()
        if not file_bytes:
            continue
        mime_type = f.mimetype or mimetypes.guess_type(original)[0] or 'application/octet-stream'
        cur.execute('''
            INSERT INTO ldud_proof_documents (ldud_id, original_filename, file_bytes, mime_type, uploaded_by)
            VALUES (%s, %s, %s, %s, %s) RETURNING id, original_filename, uploaded_at
        ''', [ldud_id, original, file_bytes, mime_type, session.get('username')])
        row = cur.fetchone()
        saved.append({'id': row['id'], 'original_filename': row['original_filename'],
                      'uploaded_at': str(row['uploaded_at'])[:16]})
    conn.commit()
    conn.close()

    if not saved:
        return jsonify({'error': 'No valid files uploaded (allowed: pdf, jpg, png, xlsx, csv, doc)'}), 400
    return jsonify({'success': True, 'docs': saved})


@bp.route('/api/module/LDUD01/proof_docs/<int:ldud_id>')
@login_required
def list_proof_docs(ldud_id):
    # Metadata-only query — never select file_bytes here, keeps the list fast.
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('''
        SELECT id, original_filename, uploaded_by, uploaded_at
        FROM ldud_proof_documents WHERE ldud_id=%s ORDER BY uploaded_at
    ''', [ldud_id])
    docs = [{'id': r['id'], 'original_filename': r['original_filename'],
              'uploaded_by': r['uploaded_by'], 'uploaded_at': str(r['uploaded_at'])[:16]}
            for r in cur.fetchall()]
    conn.close()
    return jsonify({'docs': docs})


@bp.route('/api/module/LDUD01/proof_docs/file/<int:doc_id>')
@login_required
def serve_proof_doc(doc_id):
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT original_filename, file_bytes, mime_type FROM ldud_proof_documents WHERE id=%s', [doc_id])
    row = cur.fetchone()
    conn.close()
    if not row or row['file_bytes'] is None:
        return 'Not found', 404
    ext = os.path.splitext(row['original_filename'])[1].lower()
    inline_types = {'.pdf', '.jpg', '.jpeg', '.png'}
    as_attachment = ext not in inline_types
    return send_file(
        io.BytesIO(bytes(row['file_bytes'])),
        download_name=row['original_filename'],
        mimetype=row['mime_type'] or 'application/octet-stream',
        as_attachment=as_attachment,
    )


@bp.route('/api/module/LDUD01/proof_docs/by_vcn/<int:vcn_id>')
@login_required
def proof_docs_by_vcn(vcn_id):
    """Return proof docs for the LDUD linked to a VCN (used by FIN01 billing/approval pages)."""
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute('SELECT id FROM ldud_header WHERE vcn_id=%s ORDER BY id DESC LIMIT 1', [vcn_id])
    ldud = cur.fetchone()
    if not ldud:
        conn.close()
        return jsonify({'docs': [], 'ldud_id': None})
    ldud_id = ldud['id']
    cur.execute('''
        SELECT id, original_filename, uploaded_by, uploaded_at
        FROM ldud_proof_documents WHERE ldud_id=%s ORDER BY uploaded_at
    ''', [ldud_id])
    docs = [{'id': r['id'], 'original_filename': r['original_filename'],
              'uploaded_by': r['uploaded_by'], 'uploaded_at': str(r['uploaded_at'])[:16]}
            for r in cur.fetchall()]
    conn.close()
    return jsonify({'docs': docs, 'ldud_id': ldud_id})


@bp.route('/api/module/LDUD01/proof_docs/by_bill/<int:bill_id>')
@login_required
def proof_docs_by_bill(bill_id):
    """Return all proof docs for VCN cargo lines on a bill (used by FIN01 approval)."""
    conn = get_db()
    cur = get_cursor(conn)
    # Get VCN-type cargo source IDs from this bill's lines
    cur.execute('''
        SELECT DISTINCT cargo_source_type, cargo_source_id
        FROM bill_lines
        WHERE bill_id=%s AND cargo_source_type IN ('VCN_IMPORT', 'VCN_EXPORT')
          AND cargo_source_id IS NOT NULL
    ''', [bill_id])
    sources = cur.fetchall()

    all_docs = []
    seen_ldud = set()
    for src in sources:
        table = 'vcn_cargo_declaration' if src['cargo_source_type'] == 'VCN_IMPORT' else 'vcn_export_cargo_declaration'
        cur.execute(f'SELECT vcn_id FROM {table} WHERE id=%s', [src['cargo_source_id']])
        decl = cur.fetchone()
        if not decl:
            continue
        cur.execute('SELECT id FROM ldud_header WHERE vcn_id=%s ORDER BY id DESC LIMIT 1', [decl['vcn_id']])
        ldud = cur.fetchone()
        if not ldud or ldud['id'] in seen_ldud:
            continue
        seen_ldud.add(ldud['id'])
        cur.execute('''
            SELECT id, original_filename, uploaded_by, uploaded_at
            FROM ldud_proof_documents WHERE ldud_id=%s ORDER BY uploaded_at
        ''', [ldud['id']])
        for r in cur.fetchall():
            all_docs.append({'id': r['id'], 'original_filename': r['original_filename'],
                             'uploaded_by': r['uploaded_by'], 'uploaded_at': str(r['uploaded_at'])[:16]})
    conn.close()
    return jsonify({'docs': all_docs})
