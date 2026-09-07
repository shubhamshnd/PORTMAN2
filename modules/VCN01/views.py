import io
import json as _json
from flask import Blueprint, render_template, request, jsonify, session, redirect, url_for, send_file
from functools import wraps
import excel_export
from . import model
from database import get_user_permissions, get_module_config
from modules.FIN01 import model as fin_model

bp = Blueprint('VCN01', __name__, template_folder='.')
MODULE_CODE = 'VCN01'
BILLED_LOCK_MSG = 'This vessel has been billed — the record is locked and cannot be changed.'
APPROVED_LOCK_MSG = ('This vessel is Approved and cannot be edited. An administrator must reopen it to Draft first (Admin > Reopen to Draft).')

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated

def _billed_locked(vcn_id):
    """Return a JSON 409 response if the VCN is billed (locked), else None."""
    if vcn_id and fin_model.is_vcn_billed(vcn_id):
        return jsonify({'error': BILLED_LOCK_MSG}), 409
    return None


def _approved_locked(vcn_id):
    """Return a JSON 403 if the VCN is Approved, else None.

    Approved is shut to everyone, approver included — the only way back to an
    editable record is Admin > Reopen to Draft (which itself refuses a billed
    vessel). The screen already hides these controls via isDraftRow(); this is
    the server actually holding the line.
    """
    if vcn_id and model.get_doc_status(vcn_id) == 'Approved':
        return jsonify({'error': APPROVED_LOCK_MSG}), 403
    return None

def get_perms():
    if session.get('is_admin'):
        return {'can_read': 1, 'can_add': 1, 'can_edit': 1, 'can_delete': 1}
    return get_user_permissions(session.get('user_id'), MODULE_CODE)

@bp.route('/module/VCN01/')
@login_required
def view():
    perms = get_perms()
    if not perms.get('can_read'):
        return render_template('no_access.html'), 403
    return render_template('vcn01.html', permissions=perms)

@bp.route('/api/module/VCN01/data')
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
    data, total = model.get_data(page, size, filters)
    return jsonify({'data': data, 'last_page': (total + size - 1) // size, 'total': total})

@bp.route('/api/module/VCN01/vessels')
@login_required
def get_vessels():
    return jsonify(model.get_vessels())

@bp.route('/api/module/VCN01/equipment')
@login_required
def get_equipment():
    return jsonify(model.get_equipment())

@bp.route('/api/module/VCN01/save', methods=['POST'])
@login_required
def save():
    perms = get_perms()
    data = request.json
    is_new = not data.get('id')

    if is_new and not perms.get('can_add'):
        return jsonify({'error': 'No permission to add'}), 403
    if not is_new and not perms.get('can_edit'):
        return jsonify({'error': 'No permission to edit'}), 403

    if not is_new:
        locked = _billed_locked(data['id'])
        if locked:
            return locked
        if 'vcn_doc_num' in data:   # only the number part is editable client-side
            doc_num = (data['vcn_doc_num'] or '').strip()
            if not doc_num:
                return jsonify({'error': 'VCN Doc number cannot be blank'}), 400
            if model.doc_num_taken(doc_num, data['id']):
                return jsonify({'error': f'VCN Doc {doc_num} already exists'}), 400
            data['vcn_doc_num'] = doc_num
        if model.get_doc_status(data['id']) == 'Approved':
            return jsonify({'error': APPROVED_LOCK_MSG}), 403
        data['doc_status'] = 'Draft'
    else:
        data['doc_status'] = 'Draft'

    row_id, doc_num = model.save_header(data)
    return jsonify({'success': True, 'id': row_id, 'vcn_doc_num': doc_num, 'doc_status': data.get('doc_status')})


@bp.route('/api/module/VCN01/approval_check/<int:record_id>')
@login_required
def approval_check(record_id):
    return jsonify(model.get_approval_eligibility(record_id))


@bp.route('/api/module/VCN01/approve', methods=['POST'])
@login_required
def approve():
    config = get_module_config('VCN01')
    is_approver = str(config.get('approver_id', '')) == str(session.get('user_id')) or session.get('is_admin')
    if not is_approver:
        return jsonify({'error': 'No permission to approve'}), 403
    record_id = request.json.get('id')
    if not record_id:
        return jsonify({'error': 'Missing id'}), 400
    eligibility = model.get_approval_eligibility(record_id)
    if not eligibility['eligible']:
        return jsonify({'error': 'Record not eligible for approval', 'missing': eligibility['missing']}), 400
    model.approve_record(record_id, session.get('username'))
    return jsonify({'doc_status': 'Approved'})


@bp.route('/api/module/VCN01/send_to_expected', methods=['POST'])
@login_required
def send_to_expected():
    """Undo an accidental EV01→VCN move — vessel goes back to Expected Vessels."""
    perms = get_perms()
    if not perms.get('can_edit'):
        return jsonify({'error': 'No permission to edit'}), 403
    record_id = (request.json or {}).get('id')
    if not record_id:
        return jsonify({'error': 'Missing id'}), 400
    locked = _billed_locked(record_id)
    if locked:
        return locked
    if model.get_doc_status(record_id) == 'Approved':
        return jsonify({'error': 'Approved records cannot be sent back to Expected — an admin must reopen it first via Admin > Reopen to Draft'}), 400
    try:
        model.send_back_to_expected(record_id, session.get('username'))
    except ValueError as e:
        return jsonify({'error': str(e)}), 409
    return jsonify({'success': True})


@bp.route('/api/module/VCN01/approval-log/<int:record_id>')
@login_required
def approval_log(record_id):
    return jsonify(model.get_approval_log(record_id))


@bp.route('/api/module/VCN01/delete', methods=['POST'])
@login_required
def delete():
    perms = get_perms()
    if not perms.get('can_delete'):
        return jsonify({'error': 'No permission to delete'}), 403
    model.delete_header(request.json.get('id'))
    return jsonify({'success': True})

# Consigner (customer details) endpoints
@bp.route('/api/module/VCN01/consigners/<int:vcn_id>')
@login_required
def get_consigners(vcn_id):
    return jsonify(model.get_consigners(vcn_id))

@bp.route('/api/module/VCN01/parcels/<int:vcn_id>')
@login_required
def get_parcels(vcn_id):
    """Compact parcel list for cross-module pickers (e.g. LDUD).
    Operation-type aware: Import → consigners, Export → export cargo."""
    parcels = [{
        'id': p['id'],
        'parcel_no': p.get('parcel_no'),
        'cargo_name': p.get('cargo_name'),
        'consigner_name': p.get('consigner_name'),
        'quantity': p.get('quantity'),
        'terminals': p.get('terminals', []),
    } for p in model.get_picker_parcels(vcn_id)]
    return jsonify(parcels)

@bp.route('/api/module/VCN01/cargo_quotas/<int:vcn_id>')
@login_required
def get_cargo_quotas(vcn_id):
    return jsonify(model.get_cargo_quotas(vcn_id))

@bp.route('/api/module/VCN01/consigners/save', methods=['POST'])
@login_required
def save_consigner():
    perms = get_perms()
    if not perms.get('can_add') and not perms.get('can_edit'):
        return jsonify({'error': 'No permission'}), 403
    locked = _billed_locked(request.json.get('vcn_id')) or \
             _approved_locked(request.json.get('vcn_id'))
    if locked:
        return locked
    try:
        row_id = model.save_consigner(request.json)
    except ValueError as e:
        return jsonify({'error': str(e)}), 400
    parcel = model.get_parcel(row_id) or {}
    return jsonify({'success': True, 'id': row_id,
                    'parcel_no': parcel.get('parcel_no'),
                    'parcel_seq': parcel.get('parcel_seq'),
                    'cargo_type': model.get_header_cargo_type(request.json.get('vcn_id'))})

@bp.route('/api/module/VCN01/consigners/delete', methods=['POST'])
@login_required
def delete_consigner():
    perms = get_perms()
    # sub-table rows are deletable by anyone who can edit/add (not gated on can_delete)
    if not perms.get('can_add') and not perms.get('can_edit'):
        return jsonify({'error': 'No permission'}), 403
    _cvcn = model.get_consigner_vcn_id(request.json.get('id'))
    locked = _billed_locked(_cvcn) or _approved_locked(_cvcn)
    if locked:
        return locked
    vcn_id = model.delete_consigner(request.json.get('id'))
    return jsonify({'success': True, 'cargo_type': model.get_header_cargo_type(vcn_id)})

@bp.route('/api/module/VCN01/parcels/export')
@login_required
def export_parcels_master():
    """Master export — every parcel on every VCN, import and export, one sheet."""
    if not get_perms().get('can_read'):
        return render_template('no_access.html'), 403
    return excel_export.sheet_response(model.EXPORT_PARCEL_COLS, model.export_all_parcels(),
                                       'Parcels', 'VCN01_Parcels_Master')

# Delays endpoints
@bp.route('/api/module/VCN01/delays/<int:vcn_id>')
@login_required
def get_delays(vcn_id):
    return jsonify(model.get_delays(vcn_id))

@bp.route('/api/module/VCN01/delay_window/<int:vcn_id>')
@login_required
def delay_window(vcn_id):
    """Anchorage -> Pilot Pickup window from LDUD01, bounding the delay entries."""
    return jsonify(model.get_delay_window(vcn_id))

@bp.route('/api/module/VCN01/delays/save', methods=['POST'])
@login_required
def save_delay():
    perms = get_perms()
    if not perms.get('can_add') and not perms.get('can_edit'):
        return jsonify({'error': 'No permission'}), 403
    row_id = model.save_delay(request.json)
    return jsonify({'success': True, 'id': row_id})

@bp.route('/api/module/VCN01/delays/delete', methods=['POST'])
@login_required
def delete_delay():
    perms = get_perms()
    # sub-table rows are deletable by anyone who can edit/add (not gated on can_delete)
    if not perms.get('can_add') and not perms.get('can_edit'):
        return jsonify({'error': 'No permission'}), 403
    model.delete_delay(request.json.get('id'))
    return jsonify({'success': True})

# IGM (FORM III) document — stored as BYTEA on vcn_header
@bp.route('/api/module/VCN01/igm_doc/upload/<int:vcn_id>', methods=['POST'])
@login_required
def upload_igm_doc(vcn_id):
    perms = get_perms()
    if not perms.get('can_add') and not perms.get('can_edit'):
        return jsonify({'error': 'No permission'}), 403
    f = request.files.get('file')
    if not f or not f.filename:
        return jsonify({'error': 'No file provided'}), 400
    if not f.filename.lower().endswith('.pdf'):
        return jsonify({'error': 'Only PDF files accepted'}), 400
    model.save_igm_document(vcn_id, f.filename, f.read())
    return jsonify({'success': True, 'filename': f.filename})

@bp.route('/api/module/VCN01/igm_doc/<int:vcn_id>')
@login_required
def download_igm_doc(vcn_id):
    file_bytes, filename = model.get_igm_document(vcn_id)
    if not file_bytes:
        return jsonify({'error': 'No IGM document uploaded'}), 404
    return send_file(io.BytesIO(file_bytes), mimetype='application/pdf',
                     as_attachment=False, download_name=filename or 'igm.pdf')

# Import cargo is declared via the consigner endpoints above;
# vcn_cargo_declaration stays read-only for historic data (billing/LDUD).

# Export Cargo Declaration endpoints
@bp.route('/api/module/VCN01/export_cargo/<int:vcn_id>')
@login_required
def get_export_cargo(vcn_id):
    return jsonify(model.get_export_cargo_declarations(vcn_id))

@bp.route('/api/module/VCN01/export_cargo/save', methods=['POST'])
@login_required
def save_export_cargo():
    perms = get_perms()
    if not perms.get('can_add') and not perms.get('can_edit'):
        return jsonify({'error': 'No permission'}), 403
    locked = _billed_locked(request.json.get('vcn_id')) or \
             _approved_locked(request.json.get('vcn_id'))
    if locked:
        return locked
    row_id = model.save_export_cargo_declaration(request.json)
    parcel = model.get_export_parcel(row_id) or {}
    return jsonify({'success': True, 'id': row_id,
                    'parcel_no': parcel.get('parcel_no'),
                    'parcel_seq': parcel.get('parcel_seq'),
                    'cargo_type': model.get_header_cargo_type(request.json.get('vcn_id'))})

@bp.route('/api/module/VCN01/export_cargo/delete', methods=['POST'])
@login_required
def delete_export_cargo():
    perms = get_perms()
    # sub-table rows are deletable by anyone who can edit/add (not gated on can_delete)
    if not perms.get('can_add') and not perms.get('can_edit'):
        return jsonify({'error': 'No permission'}), 403
    _evcn = model.get_export_parcel_vcn_id(request.json.get('id'))
    locked = _billed_locked(_evcn) or _approved_locked(_evcn)
    if locked:
        return locked
    vcn_id = model.delete_export_cargo_declaration(request.json.get('id'))
    return jsonify({'success': True, 'cargo_type': model.get_header_cargo_type(vcn_id)})

@bp.route('/api/module/VCN01/export_cargo_names/<int:vcn_id>')
@login_required
def get_export_cargo_names(vcn_id):
    return jsonify(model.get_export_cargo_names_for_vcn(vcn_id))

@bp.route('/api/module/VCN01/all_cargo_names/<int:vcn_id>')
@login_required
def get_all_cargo_names(vcn_id):
    return jsonify(model.get_all_cargo_names_for_vcn(vcn_id))

@bp.route('/api/module/VCN01/export_loading_totals/<int:vcn_id>')
@login_required
def get_export_loading_totals(vcn_id):
    return jsonify(model.get_export_loading_totals(vcn_id))

# Hold Completion (read-only view from LDUD data)
@bp.route('/api/module/VCN01/hold_completion/<int:vcn_id>')
@login_required
def get_hold_completion(vcn_id):
    return jsonify(model.get_hold_completion_by_vcn(vcn_id))

