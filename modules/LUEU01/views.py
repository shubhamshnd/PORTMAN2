from flask import Blueprint, render_template, request, jsonify, session, redirect, url_for
from functools import wraps
import excel_export
from . import model
from database import get_user_permissions, get_db, get_cursor

bp = Blueprint('LUEU01', __name__, template_folder='.')
MODULE_CODE = 'LUEU01'
MODULE_INFO = {'code': 'LUEU01', 'name': 'Load Unload Equipment Utilization'}


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated


def get_perms():
    if session.get('is_admin'):
        return {'can_read': 1, 'can_add': 1, 'can_edit': 1, 'can_delete': 1}
    return get_user_permissions(session.get('user_id'), MODULE_CODE)


LOCKED_MSG = 'This vessel is Closed in LDUD01 — LUEU01 entries are locked.'


def _locked(parcel_op_id):
    """JSON 409 when the op's LDUD is fully Closed, else None."""
    if parcel_op_id and model.ops_locked([parcel_op_id]):
        return jsonify({'success': False, 'error': LOCKED_MSG}), 409
    return None


@bp.route('/module/LUEU01/')
@login_required
def view():
    perms = get_perms()
    if not perms.get('can_read'):
        return render_template('no_access.html'), 403
    return render_template('lueu01.html', permissions=perms)


@bp.route('/api/module/LUEU01/vessels')
@login_required
def get_vessels():
    return jsonify(model.get_vessels_with_started_parcels())


@bp.route('/api/module/LUEU01/parcels/<int:vcn_id>')
@login_required
def get_parcels(vcn_id):
    return jsonify(model.get_started_parcels(vcn_id))


@bp.route('/api/module/LUEU01/export')
@login_required
def export_master():
    """Master export — every parcel operation, all vessels, one sheet."""
    if not get_perms().get('can_read'):
        return render_template('no_access.html'), 403
    return excel_export.sheet_response(model.EXPORT_COLS, model.export_all_parcels(),
                                       'Parcel Operations', 'LUEU01_Parcel_Operations')


@bp.route('/api/module/LUEU01/parcel/times', methods=['POST'])
@login_required
def set_parcel_times():
    perms = get_perms()
    if not perms.get('can_add') and not perms.get('can_edit'):
        return jsonify({'error': 'No permission'}), 403
    data = request.json or {}
    pid = data.get('parcel_op_id')
    if not pid:
        return jsonify({'error': 'Missing parcel_op_id'}), 400
    locked = _locked(pid)
    if locked:
        return locked
    model.set_parcel_times(pid, data.get('start_dt'), data.get('end_dt'))
    return jsonify({'success': True})


@bp.route('/api/module/LUEU01/parcel/expected_start', methods=['POST'])
@login_required
def set_expected_start():
    perms = get_perms()
    if not perms.get('can_add') and not perms.get('can_edit'):
        return jsonify({'error': 'No permission'}), 403
    data = request.json or {}
    pid = data.get('parcel_op_id')
    if not pid:
        return jsonify({'error': 'Missing parcel_op_id'}), 400
    locked = _locked(pid)
    if locked:
        return locked
    model.set_expected_start(pid, data.get('expected_start'), data.get('expected_flow_rate'))
    return jsonify({'success': True})


@bp.route('/api/module/LUEU01/log/<int:parcel_op_id>')
@login_required
def get_log(parcel_op_id):
    return jsonify(model.get_log(parcel_op_id))


@bp.route('/api/module/LUEU01/log/save', methods=['POST'])
@login_required
def save_log():
    perms = get_perms()
    if not perms.get('can_add') and not perms.get('can_edit'):
        return jsonify({'error': 'No permission'}), 403
    data = request.json or {}
    locked = _locked(data.get('parcel_op_id'))
    if locked:
        return locked
    data['created_by'] = session.get('username')
    return jsonify({'id': model.save_log(data), 'success': True})


@bp.route('/api/module/LUEU01/log/delete', methods=['POST'])
@login_required
def delete_log():
    perms = get_perms()
    if not perms.get('can_delete'):
        return jsonify({'error': 'No permission to delete'}), 403
    ids = (request.json or {}).get('ids', [])
    if not ids:
        return jsonify({'error': 'No IDs provided'}), 400
    if model.logs_locked(ids):
        return jsonify({'success': False, 'error': LOCKED_MSG}), 409
    model.soft_delete_log(ids, session.get('username'))
    return jsonify({'success': True, 'deleted_count': len(ids)})


@bp.route('/api/module/LUEU01/parcel/removed', methods=['POST'])
@login_required
def set_parcel_removed():
    """Remove or restore a VCN parcel without opening VCN01.

    The parcel is flagged, never deleted: it stays on screen greyed out and
    can be restored, but stops counting toward billing, closure and targets."""
    perms = get_perms()
    if not perms.get('can_add') and not perms.get('can_edit'):
        return jsonify({'success': False, 'error': 'No permission'}), 403
    data = request.json or {}
    pid = data.get('parcel_op_id')
    parcel_ids = data.get('parcel_ids') or []
    removed = bool(data.get('removed'))
    if not parcel_ids:
        return jsonify({'success': False, 'error': 'Missing parcel_ids'}), 400
    locked = _locked(pid)
    if locked:
        return locked
    from modules.VCN01 import model as vcn_model
    try:
        for rid in parcel_ids:
            vcn_model.set_parcel_removed(rid, removed, data.get('reason'),
                                         session.get('username'))
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})
    return jsonify({'success': True})


@bp.route('/api/module/LUEU01/parcel/additional', methods=['POST'])
@login_required
def set_additional_qty():
    """Amend the BL upward when the vessel discharged more than declared.
    Avoids reopening and editing the VCN for an over-discharge."""
    perms = get_perms()
    if not perms.get('can_add') and not perms.get('can_edit'):
        return jsonify({'success': False, 'error': 'No permission'}), 403
    data = request.json or {}
    pid = data.get('parcel_op_id')
    if not pid:
        return jsonify({'success': False, 'error': 'Missing parcel_op_id'}), 400
    locked = _locked(pid)
    if locked:
        return locked
    try:
        qty = model.set_additional_qty(pid, data.get('quantity'),
                                       data.get('reason'), session.get('username'))
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})
    return jsonify({'success': True, 'additional_qty': qty})


@bp.route('/api/module/LUEU01/parcel/additional/clear', methods=['POST'])
@login_required
def clear_additional_qty():
    perms = get_perms()
    if not perms.get('can_delete'):
        return jsonify({'success': False, 'error': 'No permission to clear'}), 403
    pid = (request.json or {}).get('parcel_op_id')
    if not pid:
        return jsonify({'success': False, 'error': 'Missing parcel_op_id'}), 400
    locked = _locked(pid)
    if locked:
        return locked
    try:
        model.clear_additional_qty(pid, session.get('username'))
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})
    return jsonify({'success': True})


@bp.route('/api/module/LUEU01/parcel/shortclose', methods=['POST'])
@login_required
def shortclose_parcel():
    perms = get_perms()
    if not perms.get('can_add') and not perms.get('can_edit'):
        return jsonify({'success': False, 'error': 'No permission'}), 403
    pid = (request.json or {}).get('parcel_op_id')
    if not pid:
        return jsonify({'success': False, 'error': 'Missing parcel_op_id'}), 400
    locked = _locked(pid)
    if locked:
        return locked
    try:
        model.shortclose_parcel(pid, session.get('username'))
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})
    return jsonify({'success': True})


@bp.route('/api/module/LUEU01/parcel/shortclose/revert', methods=['POST'])
@login_required
def revert_shortclose():
    perms = get_perms()
    if not perms.get('can_delete'):
        return jsonify({'success': False, 'error': 'No permission to revert'}), 403
    pid = (request.json or {}).get('parcel_op_id')
    if not pid:
        return jsonify({'success': False, 'error': 'Missing parcel_op_id'}), 400
    locked = _locked(pid)
    if locked:
        return locked
    try:
        model.revert_shortclose(pid, session.get('username'))
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})
    return jsonify({'success': True})


def _names(sql):
    conn = get_db(); cur = get_cursor(conn)
    cur.execute(sql); rows = cur.fetchall(); conn.close()
    return rows


@bp.route('/api/module/LUEU01/equipment')
@login_required
def get_equipment():
    return jsonify([r['name'] for r in _names('SELECT name FROM equipment ORDER BY name')])


@bp.route('/api/module/LUEU01/delays')
@login_required
def get_delays():
    return jsonify([r['name'] for r in _names('SELECT name FROM port_delay_types ORDER BY name')])


@bp.route('/api/module/LUEU01/uom')
@login_required
def get_uom():
    rows = _names('SELECT name, is_default FROM quantity_uom ORDER BY name')
    return jsonify({'names': [r['name'] for r in rows],
                    'default': next((r['name'] for r in rows if r['is_default']), '')})


@bp.route('/api/module/LUEU01/berths')
@login_required
def get_berths():
    return jsonify([r['berth_name'] for r in _names('SELECT berth_name FROM port_berth_master ORDER BY berth_name')])


@bp.route('/api/module/LUEU01/shift-incharge')
@login_required
def get_shift_incharge():
    return jsonify([r['name'] for r in _names(
        "SELECT name FROM port_shift_incharge WHERE name IS NOT NULL AND name != '' ORDER BY name")])


@bp.route('/api/module/LUEU01/shift-operators')
@login_required
def get_shift_operators():
    return jsonify([r['name'] for r in _names(
        "SELECT name FROM port_shift_operators WHERE name IS NOT NULL AND name != '' ORDER BY name")])
