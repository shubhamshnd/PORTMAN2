"""SRV01 Service Records — xlsx dump.

Approved records only: those are the billing-relevant set, and a draft is by
definition still being filled in.

SRV01's custom fields are EAV, so they differ per service type and cannot all be
fixed columns. Two shapes, one route:

  no service type  -> every approved record, custom fields collapsed into one
                      readable 'Details' cell
  a service type   -> only that type's records, with its custom fields expanded
                      into real columns you can sort and pivot on
"""
from flask import render_template, session, jsonify, request
from functools import wraps
from datetime import datetime

from . import bp
from database import get_db, get_cursor, get_user_permissions
import excel_export

MODULE_CODE = 'RP01'

# Fixed header columns, shared by both shapes. The all-types export adds the
# standard time/reading columns immediately after Date to match the required
# Excel sequence, while service-type-specific exports keep their custom fields
# appended after the base row.
BASE_COLS = [
    ('Record No',    'record_number'),
    ('Month',        'month'),
    ('Service Code', 'service_code'),
    ('Service',      'service_name'),
    ('Bill To Type', 'source_type'),
    ('Bill To',      'source_display'),
    ('Ref Document', 'ref_source_display'),
    ('Date',         'record_date'),
    ('Billable Qty', 'billable_quantity'),
    ('UOM',          'billable_uom'),
    ('Status',       'doc_status'),
    ('Billed',       'billed'),
    ('Bill No',      'bill_number'),
    ('Created By',   'created_by'),
    ('Approved By',  'approved_by'),
    ('Approved On',  'approved_date'),
    ('Remarks',      'remarks'),
]


def login_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if 'user_id' not in session:
            return jsonify({'error': 'Not logged in'}), 401
        return f(*args, **kwargs)
    return wrapper


def _custom_fields(cur, service_type_id):
    """Ordered active field definitions for one service type."""
    cur.execute('''SELECT id, field_label FROM service_field_definitions
                   WHERE service_type_id = %s AND is_active = 1
                   ORDER BY display_order, id''', [service_type_id])
    return [dict(r) for r in cur.fetchall()]


def _all_service_time_reading_columns():
    """Labels that must always appear in the flat 'All service types' export,
    sourced from EAV values — not separate fixed columns."""
    return [
        'Start Time',
        'End Time',
        'Start Reading',
        'End Reading',
    ]


def _format_datetime_value(value):
    """Return Excel-friendly date/time text like 01-08-2026 10:56."""
    if value in (None, ''):
        return ''
    value = str(value).strip()
    for fmt in (
        '%Y-%m-%d %H:%M:%S',
        '%Y-%m-%d %H:%M',
        '%Y-%m-%dT%H:%M:%S',
        '%Y-%m-%dT%H:%M',
        '%Y-%m-%d %H:%M:%S.%f',
        '%Y-%m-%dT%H:%M:%S.%f',
    ):
        try:
            return datetime.strptime(value, fmt).strftime('%d-%m-%Y %H:%M')
        except ValueError:
            continue
    # Fallback: regex-based parse for any ISO-like datetime
    import re
    m = re.match(r'(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2})', value)
    if m:
        return f'{m.group(3)}-{m.group(2)}-{m.group(1)} {m.group(4)}:{m.group(5)}'
    return value


def _month_filter_sql(month, alias='r'):
    """Return SQL+params for a YYYY-MM filter on the stored ISO text date."""
    if not month:
        return "", []
    month = str(month).strip()
    try:
        dt = datetime.strptime(month, '%Y-%m')
    except ValueError:
        return "", []
    start = f"{dt.year:04d}-{dt.month:02d}-01"
    if dt.month == 12:
        end = f"{dt.year + 1:04d}-01-01"
    else:
        end = f"{dt.year:04d}-{dt.month + 1:02d}-01"
    col = f"{alias}.record_date" if alias else "record_date"
    return f" AND {col} >= %s AND {col} < %s", [start, end]


def _status_filter_sql(status, alias='r'):
    """Status filter for report/totals: Approved, Pending or Both."""
    status = (status or 'Approved').strip()
    col = f"{alias}.doc_status" if alias else "doc_status"
    if status == 'Pending':
        return f" AND {col} = 'Pending'", []
    if status == 'Both':
        return f" AND {col} IN ('Approved', 'Pending')", []
    return f" AND {col} = 'Approved'", []


def _fy_months_to_current():
    """Return FY months from April through the current month only."""
    today = datetime.today()
    if today.month >= 4:
        start_year = today.year
        start_month = 4
    else:
        start_year = today.year - 1
        start_month = 4
    months = []
    year = start_year
    month = start_month
    while True:
        months.append((year, month))
        if year == today.year and month == today.month:
            break
        if month == 12:
            year += 1
            month = 1
        else:
            month += 1
    return months


def report_rows(service_type_id=None, month=None, status='Approved'):
    """(cols, rows) for the requested shape. One pass over matching records,
    one pass over their values — not a query per record."""
    conn = get_db()
    cur = get_cursor(conn)

    where, params = "WHERE 1=1", []
    status_where, _ = _status_filter_sql(status)
    where += status_where
    if service_type_id:
        where += ' AND r.service_type_id = %s'
        params.append(service_type_id)

    month_where, month_params = _month_filter_sql(month)
    where += month_where
    params.extend(month_params)

    cur.execute(f'''
        SELECT r.id, r.record_number, r.source_type, r.source_display,
               r.ref_source_display, r.record_date, r.billable_quantity,
               r.billable_uom, r.doc_status, r.is_billed, r.created_by,
               r.approved_by, r.approved_date, r.remarks,
               st.service_code, st.service_name,
               b.bill_number
        FROM service_records r
        LEFT JOIN finance_service_types st ON st.id = r.service_type_id
        LEFT JOIN bill_header b ON b.id = r.bill_id
        {where}
        ORDER BY r.record_number
    ''', params)
    rows = [dict(r) for r in cur.fetchall()]

    for row in rows:
        record_date = row.get('record_date')
        if record_date:
            try:
                row['month'] = datetime.strptime(str(record_date), '%Y-%m-%d').strftime('%b-%y')
            except ValueError:
                row['month'] = str(record_date)[:7]
        else:
            row['month'] = ''
        row['billed'] = bool(row.pop('is_billed', 0))

    if not rows:
        # Still return the shape the caller asked for — a sheet with no rows must
        # carry the same headers as one with rows, or the two can't be appended.
        if service_type_id:
            cols = BASE_COLS + [(d['field_label'], f"cf_{d['id']}") for d in _custom_fields(cur, service_type_id)]
        else:
            cols = [
                ('Record No',    'record_number'),
                ('Month',        'month'),
                ('Service Code', 'service_code'),
                ('Service',      'service_name'),
                ('Bill To Type', 'source_type'),
                ('Bill To',      'source_display'),
                ('Ref Document', 'ref_source_display'),
                ('Date',         'record_date'),
                ('Start Time',    '_req_start_time'),
                ('End Time',      '_req_end_time'),
                ('Start Reading', '_req_start_reading'),
                ('End Reading',   '_req_end_reading'),
            ]
            cols.extend(BASE_COLS[8:])
            cols.append(('Details', 'details'))
        conn.close()
        return cols, rows

    # Every recorded value for these records, in one query.
    ids = [r['id'] for r in rows]
    cur.execute('''SELECT v.service_record_id, v.field_value,
                          d.id AS field_id, d.field_label, d.display_order
                   FROM service_record_values v
                   JOIN service_field_definitions d ON d.id = v.field_definition_id
                   WHERE v.service_record_id = ANY(%s)
                   ORDER BY d.display_order, d.id''', [ids])
    values = [dict(r) for r in cur.fetchall()]

    if service_type_id:
        defs = _custom_fields(cur, service_type_id)
        conn.close()
        by_record = {}
        for v in values:
            by_record.setdefault(v['service_record_id'], {})[v['field_id']] = v['field_value']
        for row in rows:
            vals = by_record.get(row['id'], {})
            for d in defs:
                row[f"cf_{d['id']}"] = vals.get(d['id'], '')
        cols = BASE_COLS + [(d['field_label'], f"cf_{d['id']}") for d in defs]
        return cols, rows

    conn.close()

    # Build lookup: {record_id: {field_label: field_value}}  (label-based, avoids key-type issues)
    by_label = {}
    for v in values:
        by_label.setdefault(v['service_record_id'], {})[v['field_label']] = v['field_value']

    # Details cell (all non-empty EAV pairs)
    details = {}
    for v in values:
        if str(v['field_value'] or '').strip() == '':
            continue
        details.setdefault(v['service_record_id'], []).append(
            f"{v['field_label']}: {v['field_value']}")

    # ── Required columns (ALWAYS present, fixed synthetic keys) ─────────────
    # These 4 columns must appear even when no records carry those EAV labels.
    REQUIRED = [
        ('Start Time',    '_req_start_time',    True),   # (header, row-key, is_datetime)
        ('End Time',      '_req_end_time',       True),
        ('Start Reading', '_req_start_reading',  False),
        ('End Reading',   '_req_end_reading',    False),
    ]
    required_label_set = {label for label, _, _ in REQUIRED}

    for row in rows:
        row['details'] = ' | '.join(details.get(row['id'], []))
        rec = by_label.get(row['id'], {})
        for label, key, is_dt in REQUIRED:
            raw = rec.get(label, '')
            row[key] = _format_datetime_value(raw) if is_dt else raw

    # ── Extra EAV columns (service-specific, not in REQUIRED) ───────────────
    seen_labels = set(required_label_set)
    seen_ids    = set()
    extra_defs  = []
    for v in values:
        lbl = v['field_label']
        fid = int(v['field_id'])
        if lbl in seen_labels or fid in seen_ids:
            continue
        seen_labels.add(lbl)
        seen_ids.add(fid)
        extra_defs.append({'field_id': fid, 'field_label': lbl})

    by_record = {}
    for v in values:
        by_record.setdefault(v['service_record_id'], {})[int(v['field_id'])] = v['field_value']

    for row in rows:
        for d in extra_defs:
            row[f"cf_{d['field_id']}"] = by_record.get(row['id'], {}).get(d['field_id'], '')

    # ── Assemble column spec ─────────────────────────────────────────────────
    before_required = [
        ('Record No',    'record_number'),
        ('Month',        'month'),
        ('Service Code', 'service_code'),
        ('Service',      'service_name'),
        ('Bill To Type', 'source_type'),
        ('Bill To',      'source_display'),
        ('Ref Document', 'ref_source_display'),
        ('Date',         'record_date'),
    ]
    required_cols = [(label, key) for label, key, _ in REQUIRED]
    extra_cols    = [(d['field_label'], f"cf_{d['field_id']}") for d in extra_defs]
    after_cols    = BASE_COLS[8:]

    cols = before_required + required_cols + extra_cols + after_cols + [('Details', 'details')]
    return cols, rows


@bp.route('/module/RP01/service-records/')
@login_required
def srv_report_page():
    perms = ({'can_read': 1} if session.get('is_admin')
             else get_user_permissions(session.get('user_id'), MODULE_CODE))
    if not perms.get('can_read'):
        return render_template('no_access.html'), 403
    return render_template('service_record_report.html')


@bp.route('/api/module/RP01/service-records/summary')
@login_required
def srv_report_summary():
    """Counts plus the service types worth offering — only those that actually
    match the selected status and month filters."""
    month = request.args.get('month') or ''
    status = (request.args.get('status') or 'Approved').strip()
    conn = get_db()
    cur = get_cursor(conn)

    status_where, _ = _status_filter_sql(status, 's')
    month_where, month_params = _month_filter_sql(month, 's')
    approved_where = "WHERE 1=1" + status_where
    billed_where = "WHERE 1=1" + status_where + " AND COALESCE(s.is_billed,0) = 1"
    type_where = "WHERE 1=1" + status_where
    if month_where:
        approved_where += month_where
        billed_where += month_where
        type_where += month_where

    cur.execute(f"SELECT COUNT(*) AS c FROM service_records s {approved_where}", month_params)
    approved = cur.fetchone()['c']
    cur.execute(f"SELECT COUNT(*) AS c FROM service_records s {billed_where}", month_params)
    billed = cur.fetchone()['c']
    cur.execute('''SELECT st.id, st.service_code, st.service_name, COUNT(*) AS records
                   FROM service_records s
                   JOIN finance_service_types st ON st.id = s.service_type_id
                   {type_where}
                   GROUP BY st.id, st.service_code, st.service_name
                   ORDER BY st.service_name'''.format(type_where=type_where), month_params)
    types = [dict(r) for r in cur.fetchall()]

    month_status_where = _status_filter_sql(status, 's')[0]
    months = []
    for year, month_no in _fy_months_to_current():
        value = f"{year:04d}-{month_no:02d}"
        label = datetime.strptime(value + '-01', '%Y-%m-%d').strftime('%b')
        months.append({'value': value, 'label': label})
    conn.close()
    return jsonify({'approved': approved, 'billed': billed,
                    'unbilled': approved - billed, 'service_types': types,
                    'months': months})


@bp.route('/api/module/RP01/service-records/export')
@login_required
def srv_report_export():
    raw = request.args.get('service_type_id') or ''
    service_type_id = int(raw) if raw.isdigit() else None
    month = request.args.get('month') or ''
    status = (request.args.get('status') or 'Approved').strip()
    cols, rows = report_rows(service_type_id, month=month, status=status)
    stem = f'Service_Records_{service_type_id}' if service_type_id else 'Service_Records'
    if month:
        stem = f'{stem}_{month.replace("-", "")}'
    if status and status != 'Approved':
        stem = f'{stem}_{status.lower()}'
    return excel_export.sheet_response(cols, rows, 'Service Records', stem)
