"""
DPR Report — Monthly Cargo Break-up & Performance Report (JSW & JNPT Analysis)
Port Bird Module: RP01
Purpose: Analyse monthly cargo break up output as per JSW and JNPT calculation to cross check billing process.
Features:
  - 07:00 AM month-end cut-off rule for JSW & JNPT months
  - JNPT Month: Based on vessel sailing out (cast-off datetime)
  - JSW Month: Based on actual cargo discharged during that month
  - Unified historical MIS data (mis_history + mis_vessel_master) and live Port Bird data
  - Short shipment & excess cargo quantity adjustments
  - Rich interactive browser preview table with filters (FY, Month, JSW/JNPT Mode, Terminal, Search)
  - Formatted Excel export matching the official JJLTPL format
"""

import io
import math
import re
from datetime import datetime, date, timedelta, time
from decimal import Decimal
from functools import wraps

from flask import (
    render_template,
    request,
    jsonify,
    send_file,
    session,
    redirect,
    url_for,
    Response,
)
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from database import get_db, get_cursor, get_user_permissions
from .. import bp

MODULE_CODE = 'RP01'


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


def _parse_dt(val):
    if not val:
        return None
    if isinstance(val, datetime):
        return val
    if isinstance(val, date):
        return datetime.combine(val, time(0, 0))
    s = str(val).strip().replace('T', ' ')
    for length in (19, 16, 10):
        sub = s[:length]
        for fmt in (
            '%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%Y-%m-%d',
            '%d-%m-%Y %H:%M:%S', '%d-%m-%Y %H:%M', '%d-%m-%Y',
            '%d/%m/%Y %H:%M:%S', '%d/%m/%Y %H:%M', '%d/%m/%Y'
        ):
            try:
                return datetime.strptime(sub, fmt)
            except Exception:
                pass
    return None


def _format_display_dt(dt):
    if not dt:
        return ''
    d = _parse_dt(dt)
    if not d:
        return str(dt)
    return d.strftime('%d/%m/%y %H:%M')


def _calc_month_label(dt):
    """
    07:00 AM cutoff: Subtract 7 hours to align with 07:00 AM operational day/month window.
    Vessels sailing or cargo discharging up to 07:00 AM on 1st of month count in previous month.
    """
    if not dt:
        return ''
    d = _parse_dt(dt)
    if not d:
        return ''
    adjusted = d - timedelta(hours=7)
    return adjusted.strftime('%b-%y')


def _calc_fin_year(dt):
    if not dt:
        return ''
    d = _parse_dt(dt)
    if not d:
        return ''
    adjusted = d - timedelta(hours=7)
    y = adjusted.year
    m = adjusted.month
    sy = y if m >= 4 else y - 1
    ey = (sy + 1) % 100
    return f"{sy}-{ey:02d}"


def _safe_float(val, default=0.0):
    if val is None:
        return default
    try:
        return float(val)
    except (ValueError, TypeError):
        return default


def _load_cargo_master_map(cur):
    cur.execute("""
        SELECT cargo_name, cargo_type, cargo_category, cargo_category_2, cargo_sub_category, cargo_sub_category_2
        FROM vessel_cargo;
    """)
    cargo_map = {}
    for r in cur.fetchall():
        cn = (r.get('cargo_name') or '').strip().lower()
        if cn and cn not in cargo_map:
            cargo_map[cn] = r
    return cargo_map


def _get_active_financial_year():
    now = datetime.now()
    # If before April, previous calendar year started the FY
    start_year = now.year if now.month >= 4 else now.year - 1
    return f"{start_year}-{(start_year + 1) % 100:02d}"


def _extract_doc_sl(vmd_str, vcn_doc_str, vm_doc_str=None):
    """
    Extracts numeric sequence from vessel master doc (e.g. 'VM12' -> 12, 'VM2/MT CLEAN THRASHER' -> 2),
    or falls back to trailing digits in vcn_doc_num.
    """
    for s in (vmd_str, vm_doc_str):
        if s:
            m = re.search(r'VM\s*(\d+)', str(s), re.IGNORECASE)
            if m:
                return int(m.group(1))
            m = re.search(r'(\d+)', str(s))
            if m:
                return int(m.group(1))
    if vcn_doc_str:
        m = re.search(r'(\d+)$', str(vcn_doc_str))
        if m:
            return int(m.group(1))
    return None


def fetch_all_dpr_records(fin_year=None, month=None, mode='ALL', terminal='ALL'):
    """
    Fetches unified DPR records across historical MIS and live Port Bird pipeline.
    Calculates 07:00 AM cutoff for JSW & JNPT months, short shipment adjustments,
    and assigns Int SL (vessel sequence) and Sl. No.
    """
    conn = get_db()
    records = []
    seen_keys = set()

    try:
        cur = get_cursor(conn)
        cargo_map = _load_cargo_master_map(cur)

        # -------------------------------------------------------------
        # 1. Historical Dataset: mis_history + mis_vessel_master
        # -------------------------------------------------------------
        cur.execute("""
            SELECT 
                m.id,
                COALESCE(m.fin_year, '') AS fin_year,
                COALESCE(m.month_jsw, '') AS month_jsw,
                COALESCE(m.month_jnpt, '') AS month_jnpt,
                COALESCE(m.vcn_no, '') AS vcn_no,
                COALESCE(m.vessel_name, '') AS vessel_name,
                COALESCE(m.customer, '') AS company,
                COALESCE(m.payment_by, '') AS payment_by,
                COALESCE(m.cargo_sub_category_2, '') AS category0,
                COALESCE(m.cargo_sub_category, '') AS category1,
                COALESCE(m.cargo_category, '') AS cargo,
                COALESCE(m.cargo_name, '') AS cargo1,
                v.ops_commenced AS op_start,
                v.cargo_completion AS op_comp,
                COALESCE(m.terminal, '') AS unloading_terminal,
                COALESCE(m.quantity, 0)::float8 AS quantity,
                COALESCE(m.overseas_coastal, '') AS overseas_coastal,
                COALESCE(m.import_export, '') AS type,
                COALESCE(m.gangway_agent, '') AS vessel_agent,
                COALESCE(m.remarks, '') AS remarks,
                COALESCE(m.importer, '') AS importer
            FROM mis_history m
            LEFT JOIN mis_vessel_master v ON v.vcn_no = m.vcn_no
            WHERE (%s IS NULL OR m.fin_year = %s)
            ORDER BY m.id ASC;
        """, (fin_year, fin_year))

        hist_rows = cur.fetchall()
        for r in hist_rows:
            row = dict(r)
            c_name = row['cargo1'].strip()
            cg_info = cargo_map.get(c_name.lower(), {})
            cat0 = row['category0'] or cg_info.get('cargo_sub_category_2') or cg_info.get('cargo_category_2') or cg_info.get('cargo_category') or ''
            cat1 = row['category1'] or cg_info.get('cargo_sub_category') or cg_info.get('cargo_category') or ''
            cargo_type = row['cargo'] or cg_info.get('cargo_category') or cg_info.get('cargo_sub_category') or ''

            key = ('hist', row['id'])
            seen_keys.add(key)

            records.append({
                'source': 'historical',
                'month_jsw': row['month_jsw'],
                'month_jnpt': row['month_jnpt'],
                'vcn_no': row['vcn_no'],
                'vessel_name': row['vessel_name'],
                'company': row['company'],
                'payment_by': row['payment_by'],
                'category0': cat0,
                'category1': cat1,
                'cargo': cargo_type,
                'cargo1': c_name,
                'op_start': _format_display_dt(row['op_start']),
                'op_comp': _format_display_dt(row['op_comp']),
                'unloading_terminal': row['unloading_terminal'],
                'quantity': round(_safe_float(row['quantity']), 3),
                'overseas_coastal': row['overseas_coastal'],
                'type': row['type'],
                'vessel_agent': row['vessel_agent'],
                'remarks': row['remarks'],
                'importer': row['importer'],
                'fin_year': row['fin_year'],
                'sort_dt': _parse_dt(row['op_start']) or datetime.min,
            })

        # -------------------------------------------------------------
        # 2. Live Dataset: Port Bird Operational Pipeline
        # -------------------------------------------------------------
        # A. Query parcels with LDUD parcel operations
        cur.execute("""
            SELECT 
                v.id AS vcn_id,
                COALESCE(v.via_number, v.vcn_doc_num) AS vcn_no,
                v.vcn_doc_num,
                v.vessel_master_doc,
                v.vessel_name,
                COALESCE(v.vessel_agent_name, '') AS vessel_agent,
                COALESCE(v.vessel_run_type, '') AS overseas_coastal,
                COALESCE(v.operation_type, '') AS type,
                COALESCE(v.remarks, '') AS vcn_remarks,
                COALESCE(v.customer_name, '') AS customer_name,
                COALESCE(v.importer_exporter_name, '') AS importer_exporter_name,
                v.doc_date,
                l.id AS ldud_id,
                l.doc_status AS ldud_status,
                l.cast_off_datetime,
                l.discharge_commenced,
                l.discharge_completed,
                po.id AS parcel_op_id,
                po.cargo_name AS po_cargo_name,
                po.terminal_name AS po_terminal_name,
                po.start_dt AS po_start,
                po.end_dt AS po_end,
                po.quantity AS po_qty,
                po.parcel_ids
            FROM vcn_header v
            JOIN ldud_header l ON l.vcn_id = v.id AND COALESCE(l.is_deleted, FALSE) = FALSE
            JOIN ldud_parcel_ops po ON po.ldud_id = l.id
            WHERE COALESCE(v.doc_status, '') != 'Cancelled'
            ORDER BY v.id DESC, po.id ASC;
        """)
        live_parcel_rows = cur.fetchall()

        # Pre-cache vessel master document numbers from vessels table
        cur.execute("SELECT doc_num, vessel_name FROM vessels WHERE doc_num IS NOT NULL;")
        vessels_doc_by_name = {
            (r['vessel_name'] or '').strip().lower(): r['doc_num'].strip()
            for r in cur.fetchall()
            if r['vessel_name']
        }

        # Query all logs from lueu_parcel_log for month-wise discharge splitting
        cur.execute("""
            SELECT 
                l.parcel_op_id,
                l.entry_date,
                l.from_time,
                l.to_time,
                COALESCE(l.quantity, 0)::float8 AS quantity,
                COALESCE(l.is_shortclose, FALSE) AS is_shortclose,
                COALESCE(l.remarks, '') AS remarks
            FROM lueu_parcel_log l
            WHERE COALESCE(l.is_deleted, FALSE) = FALSE
            ORDER BY l.parcel_op_id, l.entry_date, l.from_time;
        """)
        all_logs = cur.fetchall()

        # Group logs by (parcel_op_id, jsw_month)
        po_logs_by_month = {}
        for log in all_logs:
            p_id = log['parcel_op_id']
            ed = log['entry_date']
            ft = log['from_time'] or '00:00'
            tt = log['to_time'] or ft
            s_dt = _parse_dt(f"{ed} {ft}")
            e_dt = _parse_dt(f"{ed} {tt}")
            m_jsw = _calc_month_label(s_dt)
            if not m_jsw:
                continue

            key = (p_id, m_jsw)
            if key not in po_logs_by_month:
                po_logs_by_month[key] = {
                    'min_start': s_dt,
                    'max_end': e_dt,
                    'actual_qty': 0.0,
                    'shortclose_qty': 0.0,
                    'month_jsw': m_jsw
                }
            g = po_logs_by_month[key]
            if s_dt and (not g['min_start'] or s_dt < g['min_start']):
                g['min_start'] = s_dt
            if e_dt and (not g['max_end'] or e_dt > g['max_end']):
                g['max_end'] = e_dt

            is_sc = log['is_shortclose'] or 'short' in (log['remarks'] or '').lower()
            q = float(log['quantity'])
            if is_sc:
                g['shortclose_qty'] += q
            else:
                g['actual_qty'] += q

        # Pre-cache consigners by id and by vcn_id
        cur.execute("""
            SELECT id, vcn_id, consigner_name, importer_name, cargo_name, unload_terminal, quantity
            FROM vcn_consigners
            UNION ALL
            SELECT id, vcn_id, consigner_name, importer_name, cargo_name, unload_terminal, quantity
            FROM vcn_export_cargo_declaration;
        """)
        consigners_by_id = {}
        consigners_by_vcn = {}
        for cr in cur.fetchall():
            consigners_by_id[cr['id']] = cr
            consigners_by_vcn.setdefault(cr['vcn_id'], []).append(cr)

        live_ldud_vcns = set()

        for lr in live_parcel_rows:
            d = dict(lr)
            live_ldud_vcns.add(d['vcn_id'])

            vcn_no = (d['vcn_no'] or '').strip()
            vessel_name = (d['vessel_name'] or '').strip()
            if not vcn_no and not vessel_name:
                continue

            # Extract Vessel Serial Number from vessel_master_doc (e.g. 'VM12' -> 12, 'VM2/MT CLEAN THRASHER' -> 2)
            doc_sl = _extract_doc_sl(
                d.get('vessel_master_doc'),
                d.get('vcn_doc_num'),
                vessels_doc_by_name.get(vessel_name.lower())
            )

            # Consigner / Importer resolution
            cargo_name = d['po_cargo_name'] or ''
            term_val = d['po_terminal_name'] or ''
            company = ''
            importer = ''

            pids = [int(x) for x in str(d.get('parcel_ids') or '').split(',') if x.strip().isdigit()]
            for pid in pids:
                if pid in consigners_by_id:
                    c_info = consigners_by_id[pid]
                    company = c_info.get('consigner_name') or company
                    importer = c_info.get('importer_name') or importer
                    if not cargo_name:
                        cargo_name = c_info.get('cargo_name') or ''
                    if not term_val:
                        term_val = c_info.get('unload_terminal') or ''
                    break

            if not company and d['vcn_id'] in consigners_by_vcn:
                for vc in consigners_by_vcn[d['vcn_id']]:
                    if not cargo_name or vc.get('cargo_name', '').strip().lower() == cargo_name.strip().lower():
                        company = vc.get('consigner_name') or company
                        importer = vc.get('importer_name') or importer
                        if not term_val:
                            term_val = vc.get('unload_terminal') or ''
                        break

            # Payment party: Vessel Call Screen party (NEVER fallback to vessel agent!)
            payment_by = (
                (d.get('customer_name') or '').strip()
                or (d.get('importer_exporter_name') or '').strip()
                or company
                or importer
                or ''
            )

            # Cargo hierarchy
            c_key = cargo_name.strip().lower()
            cg_info = cargo_map.get(c_key, {})
            cat0 = cg_info.get('cargo_sub_category_2') or cg_info.get('cargo_category_2') or cg_info.get('cargo_category') or ''
            cat1 = cg_info.get('cargo_sub_category') or cg_info.get('cargo_category') or ''
            cargo_type = cg_info.get('cargo_category') or cg_info.get('cargo_sub_category') or ''

            # JNPT Month: Based on vessel sailing out (cast_off_datetime - 7 hours)
            cast_off_raw = d['cast_off_datetime']
            dt_jnpt = _parse_dt(cast_off_raw or d['discharge_completed'] or d['po_end'] or d['discharge_commenced'] or d['po_start'] or d['doc_date'])
            m_jnpt = _calc_month_label(dt_jnpt)

            # Check if this PO has logs in po_logs_by_month
            p_id = d['parcel_op_id']
            matching_months = [k[1] for k in po_logs_by_month if k[0] == p_id]

            if matching_months:
                for m_jsw in matching_months:
                    grp = po_logs_by_month[(p_id, m_jsw)]
                    fy = _calc_fin_year(dt_jnpt or grp['min_start'])
                    if fin_year and fy != fin_year:
                        continue

                    po_qty = _safe_float(d['po_qty'])
                    actual_qty = grp['actual_qty']
                    short_qty = grp['shortclose_qty']

                    if actual_qty > 0:
                        final_qty = actual_qty
                    elif len(matching_months) == 1:
                        final_qty = max(0.0, po_qty - short_qty)
                    else:
                        final_qty = 0.0

                    records.append({
                        'source': 'live_po_split',
                        'doc_sl': doc_sl,
                        'month_jsw': m_jsw,
                        'month_jnpt': m_jnpt,
                        'vcn_no': vcn_no,
                        'vessel_name': vessel_name,
                        'company': company,
                        'payment_by': payment_by,
                        'category0': cat0,
                        'category1': cat1,
                        'cargo': cargo_type,
                        'cargo1': cargo_name,
                        'op_start': _format_display_dt(grp['min_start']),
                        'op_comp': _format_display_dt(grp['max_end']),
                        'unloading_terminal': term_val,
                        'quantity': round(final_qty, 3),
                        'overseas_coastal': d['overseas_coastal'],
                        'type': d['type'],
                        'vessel_agent': d['vessel_agent'],
                        'remarks': d['vcn_remarks'],
                        'importer': importer,
                        'fin_year': fy,
                        'sort_dt': grp['min_start'] or dt_jnpt or datetime.min,
                    })
            else:
                # No equipment logs yet: single row based on parcel op dates
                dt_jsw = _parse_dt(d['po_start'] or d['discharge_commenced'] or cast_off_raw or d['doc_date'])
                m_jsw = _calc_month_label(dt_jsw)
                fy = _calc_fin_year(dt_jnpt or dt_jsw)
                if fin_year and fy != fin_year:
                    continue

                po_qty = _safe_float(d['po_qty'])

                records.append({
                    'source': 'live_po_nolog',
                    'doc_sl': doc_sl,
                    'month_jsw': m_jsw,
                    'month_jnpt': m_jnpt,
                    'vcn_no': vcn_no,
                    'vessel_name': vessel_name,
                    'company': company,
                    'payment_by': payment_by,
                    'category0': cat0,
                    'category1': cat1,
                    'cargo': cargo_type,
                    'cargo1': cargo_name,
                    'op_start': _format_display_dt(d['po_start'] or d['discharge_commenced']),
                    'op_comp': _format_display_dt(d['po_end'] or d['discharge_completed']),
                    'unloading_terminal': term_val,
                    'quantity': round(po_qty, 3),
                    'overseas_coastal': d['overseas_coastal'],
                    'type': d['type'],
                    'vessel_agent': d['vessel_agent'],
                    'remarks': d['vcn_remarks'],
                    'importer': importer,
                    'fin_year': fy,
                    'sort_dt': dt_jsw or dt_jnpt or datetime.min,
                })

        # B. For vessel calls without LDUD operations yet, load declared parcels directly
        cur.execute("""
            SELECT 
                v.id AS vcn_id,
                COALESCE(v.via_number, v.vcn_doc_num) AS vcn_no,
                v.vcn_doc_num,
                v.vessel_master_doc,
                v.vessel_name,
                COALESCE(v.vessel_agent_name, '') AS vessel_agent,
                COALESCE(v.vessel_run_type, '') AS overseas_coastal,
                COALESCE(v.operation_type, '') AS type,
                COALESCE(v.remarks, '') AS vcn_remarks,
                COALESCE(v.customer_name, '') AS customer_name,
                COALESCE(v.importer_exporter_name, '') AS importer_exporter_name,
                v.doc_date
            FROM vcn_header v
            WHERE COALESCE(v.doc_status, '') != 'Cancelled'
            ORDER BY v.id DESC;
        """)
        all_live_vcns = cur.fetchall()

        for av in all_live_vcns:
            v_id = av.get('vcn_id') or av.get('id')
            if v_id in live_ldud_vcns:
                continue

            vcn_consigners = consigners_by_vcn.get(v_id, [])
            dt_base = _parse_dt(av['doc_date'])
            m_jsw = _calc_month_label(dt_base)
            m_jnpt = _calc_month_label(dt_base)
            fy = _calc_fin_year(dt_base)

            if fin_year and fy != fin_year:
                continue

            vcn_no = (av['vcn_no'] or '').strip()
            vessel_name = (av['vessel_name'] or '').strip()

            # Extract Vessel Serial Number from vessel_master_doc (e.g. 'VM12' -> 12)
            doc_sl = _extract_doc_sl(
                av.get('vessel_master_doc'),
                av.get('vcn_doc_num'),
                vessels_doc_by_name.get(vessel_name.lower())
            )

            # Payment party: NEVER vessel agent!
            payment_by = (av.get('customer_name') or '').strip() or (av.get('importer_exporter_name') or '').strip()

            if not vcn_consigners:
                # Vessel with no parcel records declared yet
                records.append({
                    'source': 'live_declared',
                    'doc_sl': doc_sl,
                    'month_jsw': m_jsw,
                    'month_jnpt': m_jnpt,
                    'vcn_no': vcn_no,
                    'vessel_name': vessel_name,
                    'company': '',
                    'payment_by': payment_by,
                    'category0': '',
                    'category1': '',
                    'cargo': '',
                    'cargo1': '',
                    'op_start': '',
                    'op_comp': '',
                    'unloading_terminal': '',
                    'quantity': 0.0,
                    'overseas_coastal': av['overseas_coastal'],
                    'type': av['type'],
                    'vessel_agent': av['vessel_agent'],
                    'remarks': av['vcn_remarks'],
                    'importer': '',
                    'fin_year': fy,
                    'sort_dt': dt_base or datetime.min,
                })
            else:
                for cr in vcn_consigners:
                    c_name = cr.get('cargo_name') or ''
                    cg_info = cargo_map.get(c_name.strip().lower(), {})
                    cat0 = cg_info.get('cargo_sub_category_2') or cg_info.get('cargo_category_2') or cg_info.get('cargo_category') or ''
                    cat1 = cg_info.get('cargo_sub_category') or cg_info.get('cargo_category') or ''
                    cargo_type = cg_info.get('cargo_category') or cg_info.get('cargo_sub_category') or ''
                    q_val = _safe_float(cr.get('quantity'))
                    cr_party = cr.get('consigner_name') or cr.get('importer_name') or ''

                    records.append({
                        'source': 'live_declared',
                        'doc_sl': doc_sl,
                        'month_jsw': m_jsw,
                        'month_jnpt': m_jnpt,
                        'vcn_no': vcn_no,
                        'vessel_name': vessel_name,
                        'company': cr.get('consigner_name') or '',
                        'payment_by': payment_by or cr_party,
                        'category0': cat0,
                        'category1': cat1,
                        'cargo': cargo_type,
                        'cargo1': c_name,
                        'op_start': '',
                        'op_comp': '',
                        'unloading_terminal': cr.get('unload_terminal') or '',
                        'quantity': round(q_val, 3),
                        'overseas_coastal': av['overseas_coastal'],
                        'type': av['type'],
                        'vessel_agent': av['vessel_agent'],
                        'remarks': av['vcn_remarks'],
                        'importer': cr.get('importer_name') or '',
                        'fin_year': fy,
                        'sort_dt': dt_base or datetime.min,
                    })

    finally:
        conn.close()

    # Group parcels by vessel: calculate earliest datetime per vessel call and sort order
    def _month_sort_key(m_str):
        for fmt in ('%b-%y', '%b-%Y', '%B-%y', '%B-%Y'):
            try:
                return datetime.strptime(m_str, fmt)
            except Exception:
                pass
        return datetime.min

    vessel_sort_meta = {}
    for r in records:
        v_key = (r['vcn_no'].strip().upper(), r['vessel_name'].strip().upper())
        m_dt = _month_sort_key(r['month_jsw'] or r['month_jnpt'] or '')
        doc_sl = r.get('doc_sl') if r.get('doc_sl') is not None else 999999
        dt = r['sort_dt']
        meta = (m_dt, doc_sl, dt)
        if v_key not in vessel_sort_meta or meta < vessel_sort_meta[v_key]:
            vessel_sort_meta[v_key] = meta

    # Sort records so all parcels of each vessel stay grouped together chronologically
    records.sort(key=lambda x: (
        vessel_sort_meta.get((x['vcn_no'].strip().upper(), x['vessel_name'].strip().upper()), (datetime.min, 999999, datetime.min)),
        x['vcn_no'].strip().upper(),
        x['vessel_name'].strip().upper(),
        x['sort_dt']
    ))

    # Assign global vessel index and month-wise vessel index across the full FY dataset
    global_vessel_idx = 0
    curr_global_vessel = None
    month_vessel_indices = {}

    for r in records:
        v_key = (r['vcn_no'], r['vessel_name'])
        m_key = r['month_jsw'] or r['month_jnpt'] or 'Unknown'
        if v_key != curr_global_vessel:
            curr_global_vessel = v_key
            global_vessel_idx += 1
            month_vessel_indices[m_key] = month_vessel_indices.get(m_key, 0) + 1
            r['vessel_global_sl'] = global_vessel_idx
            r['vessel_month_sl'] = month_vessel_indices[m_key]
            r['is_vessel_first'] = True
        else:
            r['vessel_global_sl'] = global_vessel_idx
            r['vessel_month_sl'] = month_vessel_indices.get(m_key, 1)
            r['is_vessel_first'] = False

    # Filter by Month and Mode
    filtered = []
    for r in records:
        if month and month.upper() != 'ALL':
            if mode == 'JSW':
                if r['month_jsw'].upper() != month.upper():
                    continue
            elif mode == 'JNPT':
                if r['month_jnpt'].upper() != month.upper():
                    continue
            else:
                if r['month_jsw'].upper() != month.upper() and r['month_jnpt'].upper() != month.upper():
                    continue

        if terminal and terminal.upper() != 'ALL':
            if terminal.upper() not in r['unloading_terminal'].upper():
                continue

        filtered.append(r)

    # In the filtered view, format Sl. No and Int SL:
    # First parcel of each vessel shows Sl. No and Int SL; subsequent parcels are blank
    curr_vessel = None
    curr_month = None
    vessel_month_counter = 0
    vessels_in_view = set()
    total_qty = 0.0

    for r in filtered:
        total_qty += r['quantity']
        v_key = (r['vcn_no'].strip().upper(), r['vessel_name'].strip().upper())
        m_label = r['month_jsw'] or r['month_jnpt'] or ''
        vessels_in_view.add(v_key)

        if v_key != curr_vessel:
            curr_vessel = v_key
            if month and month.upper() != 'ALL':
                vessel_month_counter += 1
            else:
                if m_label != curr_month:
                    curr_month = m_label
                    vessel_month_counter = 1
                else:
                    vessel_month_counter += 1

            # Sl. No: Exact number from doc no (e.g. VCN-2627-001 -> 1), or global sequence for historical
            r['sl_no'] = r.get('doc_sl') if r.get('doc_sl') is not None else r['vessel_global_sl']
            r['int_sl'] = vessel_month_counter
            r['is_first_parcel'] = True
        else:
            r['sl_no'] = ''
            r['int_sl'] = ''
            r['is_first_parcel'] = False

    vessel_count = len(vessels_in_view)
    return filtered, total_qty, vessel_count


# =====================================================================
# ROUTES
# =====================================================================

@bp.route('/module/RP01/dpr-report/')
@login_required
def dpr_report_page():
    perms = get_perms()
    if not perms.get('can_read'):
        return render_template('no_access.html'), 403
    return render_template('dpr_report/dpr_report.html', permissions=perms)


@bp.route('/api/module/RP01/dpr-report/meta')
@login_required
def dpr_report_meta():
    try:
        conn = get_db()
        cur = get_cursor(conn)
        
        # 1. Dynamically gather all financial years from historical MIS and live VCNs
        cur.execute("""
            SELECT DISTINCT fin_year FROM mis_history WHERE fin_year IS NOT NULL AND TRIM(fin_year) != ''
            UNION
            SELECT DISTINCT fin_year FROM mis_vessel_master WHERE fin_year IS NOT NULL AND TRIM(fin_year) != '';
        """)
        years_set = {r['fin_year'].strip() for r in cur.fetchall() if r['fin_year']}
        
        cur.execute("SELECT DISTINCT doc_date FROM vcn_header WHERE doc_date IS NOT NULL;")
        for r in cur.fetchall():
            fy = _calc_fin_year(r['doc_date'])
            if fy:
                years_set.add(fy)

        cur_fy = _get_active_financial_year()
        if cur_fy:
            years_set.add(cur_fy)

        years = sorted(list(years_set), reverse=True)
        default_fy = cur_fy if cur_fy in years else (years[0] if years else "")

        # 2. Dynamically extract months that ACTUALLY EXIST in the database per financial year
        cur.execute("""
            SELECT fin_year, month_jsw AS m FROM mis_history WHERE month_jsw IS NOT NULL AND TRIM(month_jsw) != ''
            UNION
            SELECT fin_year, month_jnpt AS m FROM mis_history WHERE month_jnpt IS NOT NULL AND TRIM(month_jnpt) != '';
        """)
        raw_months_by_fy = {}
        for r in cur.fetchall():
            fy = (r['fin_year'] or '').strip()
            m = (r['m'] or '').strip()
            if fy and m:
                raw_months_by_fy.setdefault(fy, set()).add(m)

        # Include live pipeline actual months from Port Bird
        cur.execute("""
            SELECT 
                v.doc_date, l.cast_off_datetime, l.discharge_commenced, l.discharge_completed,
                po.start_dt, po.end_dt
            FROM vcn_header v
            JOIN ldud_header l ON l.vcn_id = v.id AND COALESCE(l.is_deleted, FALSE) = FALSE
            JOIN ldud_parcel_ops po ON po.ldud_id = l.id
            WHERE COALESCE(v.doc_status, '') != 'Cancelled';
        """)
        for r in cur.fetchall():
            op_start = r['start_dt'] or r['discharge_commenced']
            cast_off = r['cast_off_datetime']
            dt_jsw = _parse_dt(op_start or cast_off or r['doc_date'])
            dt_jnpt = _parse_dt(cast_off or op_start or r['doc_date'])
            m_jsw = _calc_month_label(dt_jsw)
            m_jnpt = _calc_month_label(dt_jnpt)
            fy = _calc_fin_year(dt_jnpt or dt_jsw)
            if fy:
                if m_jsw:
                    raw_months_by_fy.setdefault(fy, set()).add(m_jsw)
                if m_jnpt:
                    raw_months_by_fy.setdefault(fy, set()).add(m_jnpt)

        # Include live pipeline equipment logs actual months
        cur.execute("""
            SELECT entry_date, from_time FROM lueu_parcel_log WHERE COALESCE(is_deleted, FALSE) = FALSE;
        """)
        for r in cur.fetchall():
            ed = r['entry_date']
            ft = r['from_time'] or '00:00'
            dt = _parse_dt(f"{ed} {ft}")
            if dt:
                m = _calc_month_label(dt)
                fy = _calc_fin_year(dt)
                if fy and m:
                    raw_months_by_fy.setdefault(fy, set()).add(m)
                    years_set.add(fy)

        # Helper to sort months chronologically based on their parsed datetime
        def _month_sort_key(m_str):
            for fmt in ('%b-%y', '%b-%Y', '%B-%y', '%B-%Y'):
                try:
                    return datetime.strptime(m_str, fmt)
                except Exception:
                    pass
            return datetime.min

        months_by_fy = {}
        for fy, m_set in raw_months_by_fy.items():
            months_by_fy[fy] = sorted(list(m_set), key=_month_sort_key)

        # 3. Dynamically gather all distinct terminals from MIS and Port Bird
        cur.execute("""
            SELECT DISTINCT terminal FROM mis_history WHERE terminal IS NOT NULL AND TRIM(terminal) != ''
            UNION
            SELECT DISTINCT unload_terminal FROM vcn_consigners WHERE unload_terminal IS NOT NULL AND TRIM(unload_terminal) != ''
            UNION
            SELECT DISTINCT terminal_name FROM ldud_parcel_ops WHERE terminal_name IS NOT NULL AND TRIM(terminal_name) != '';
        """)
        raw_terms = set()
        for r in cur.fetchall():
            val = list(r.values())[0]
            if val:
                for part in str(val).split(','):
                    p = part.strip()
                    if p:
                        raw_terms.add(p)
        terminals = sorted(list(raw_terms))
        
        conn.close()

        return jsonify({
            'years': years,
            'current_fy': default_fy,
            'months': months_by_fy.get(default_fy, []),
            'months_by_fy': months_by_fy,
            'terminals': terminals
        })
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@bp.route('/api/module/RP01/dpr-report/data')
@login_required
def dpr_report_data():
    try:
        cur_fy = _get_active_financial_year()
        fin_year = request.args.get('fin_year', cur_fy).strip()
        month = request.args.get('month', '').strip()
        mode = request.args.get('mode', 'ALL').strip().upper()
        terminal = request.args.get('terminal', 'ALL').strip()

        records, total_qty, vessel_count = fetch_all_dpr_records(
            fin_year=fin_year,
            month=month if month != 'ALL' else None,
            mode=mode,
            terminal=terminal
        )

        return jsonify({
            'fin_year': fin_year,
            'month': month or 'ALL',
            'mode': mode,
            'terminal': terminal,
            'records': records,
            'total_quantity': round(total_qty, 3),
            'vessel_count': vessel_count,
            'record_count': len(records)
        })
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@bp.route('/api/module/RP01/dpr-report/export')
@login_required
def dpr_report_export():
    try:
        cur_fy = _get_active_financial_year()
        fin_year = request.args.get('fin_year', cur_fy).strip()
        month = request.args.get('month', '').strip()
        mode = request.args.get('mode', 'ALL').strip().upper()
        terminal = request.args.get('terminal', 'ALL').strip()

        records, total_qty, vessel_count = fetch_all_dpr_records(
            fin_year=fin_year,
            month=month if month != 'ALL' else None,
            mode=mode,
            terminal=terminal
        )

        wb = Workbook()
        ws = wb.active
        ws.title = "MONTHLY MIS - JJLTPL"

        # Styles
        font_title = Font(name="Calibri", size=11, bold=True, color="000000")
        font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        font_cell = Font(name="Calibri", size=10)
        font_total = Font(name="Calibri", size=10, bold=True, color="000000")

        fill_header = PatternFill("solid", fgColor="5B67A5")  # Purple/Indigo matching attached excel
        fill_total = PatternFill("solid", fgColor="FEF08A")   # Light yellow total highlight
        thin_border_side = Side(style="thin", color="CBD5E1")
        border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

        align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
        align_right = Alignment(horizontal="right", vertical="center")

        # Row 1: Sheet Title
        ws.cell(row=1, column=1, value="MONTHLY MIS - JJLTPL").font = font_title
        ws.row_dimensions[1].height = 20

        # Row 2: Table Headers
        headers = [
            ("Sl. No", 8, align_center),
            ("Int SL", 8, align_center),
            ("Month JSW", 12, align_center),
            ("Month JNPT", 12, align_center),
            ("VCN No", 12, align_center),
            ("Vessel Name", 26, align_left),
            ("Company", 28, align_left),
            ("Payment will be made by", 28, align_left),
            ("Category0", 16, align_left),
            ("Category 1", 16, align_left),
            ("Cargo", 16, align_left),
            ("Cargo1", 22, align_left),
            ("Operation Start", 18, align_center),
            ("Operation Completed", 18, align_center),
            ("Unloading Terminal", 18, align_center),
            ("Quantity", 16, align_right),
            ("Overseas/Costal", 16, align_center),
            ("Type", 12, align_center),
            ("Remarks", 20, align_left),
            ("Importer", 28, align_left),
            ("Year", 12, align_center),
            ("Agent", 24, align_left)
        ]

        ws.row_dimensions[2].height = 32
        for col_idx, (h_name, width, h_align) in enumerate(headers, 1):
            cell = ws.cell(row=2, column=col_idx, value=h_name)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = align_center
            cell.border = border
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = width

        # Data Rows
        curr_row = 3
        for r in records:
            ws.row_dimensions[curr_row].height = 20
            row_data = [
                (r['sl_no'], align_center, None),
                (r['int_sl'], align_center, None),
                (r['month_jsw'], align_center, None),
                (r['month_jnpt'], align_center, None),
                (r['vcn_no'], align_center, None),
                (r['vessel_name'], align_left, None),
                (r['company'], align_left, None),
                (r['payment_by'], align_left, None),
                (r['category0'], align_left, None),
                (r['category1'], align_left, None),
                (r['cargo'], align_left, None),
                (r['cargo1'], align_left, None),
                (r['op_start'], align_center, None),
                (r['op_comp'], align_center, None),
                (r['unloading_terminal'], align_center, None),
                (r['quantity'], align_right, "#,##0.000"),
                (r['overseas_coastal'], align_center, None),
                (r['type'], align_center, None),
                (r['remarks'], align_left, None),
                (r['importer'], align_left, None),
                (r['fin_year'], align_center, None),
                (r['vessel_agent'], align_left, None),
            ]

            for col_idx, (val, c_align, num_fmt) in enumerate(row_data, 1):
                cell = ws.cell(row=curr_row, column=col_idx, value=val)
                cell.font = font_cell
                cell.alignment = c_align
                cell.border = border
                if num_fmt:
                    cell.number_format = num_fmt

            curr_row += 1

        # Total Row
        ws.row_dimensions[curr_row].height = 22
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=curr_row, column=col_idx)
            cell.font = font_total
            cell.fill = fill_total
            cell.border = border

        ws.cell(row=curr_row, column=6, value="Total").alignment = align_left
        qty_cell = ws.cell(row=curr_row, column=16, value=round(total_qty, 3))
        qty_cell.number_format = "#,##0.000"
        qty_cell.alignment = align_right

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        filename_month = month if month and month.upper() != 'ALL' else 'FullYear'
        filename = f"DPR_Report_{fin_year}_{filename_month}.xlsx"

        return send_file(
            output,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500
