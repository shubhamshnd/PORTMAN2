"""
Other Statistics Report — RP01 Module
====================================
Folder: modules/RP01/RP01/statistics_report

Cross-checks vessel call, parcel, terminal, pipeline, agent, run type,
flag, and port data from actual database tables.

Aggregates statistics based on vessels whose Cast-Off or Completion date-time
falls within the selected reporting period (FY / Month).

Sections:
1. Terminal + Pipeline + Cargo Wise
2. Vessel Agent Wise
3. Flag Wise (cross-checked with vessel_flags)
4. Port Wise (cross-checked with port_master)
5. Vessel Run Type Wise (supplementary summary)
6. Vessel-Level Cross-Check Detail
7. Validation Errors
"""

import calendar
import io
import re
from datetime import datetime, date
from functools import wraps

from flask import render_template, session, redirect, url_for, request, jsonify, send_file
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

from database import get_db, get_cursor, get_user_permissions
from .. import bp

MODULE_CODE = 'RP01'
CUTOFF_DATE = date(2026, 7, 1)

MONTH_NAMES = [
    "All", "April", "May", "June", "July", "August",
    "September", "October", "November", "December", "January", "February", "March"
]

MONTH_MAP = {
    "April": 4, "May": 5, "June": 6, "July": 7, "August": 8, "September": 9,
    "October": 10, "November": 11, "December": 12, "January": 1, "February": 2, "March": 3
}

MONTH_SHORT_MAP = {
    "Apr": 4, "May": 5, "Jun": 6, "Jul": 7, "Aug": 8, "Sep": 9,
    "Oct": 10, "Nov": 11, "Dec": 12, "Jan": 1, "Feb": 2, "Mar": 3
}


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
    if isinstance(val, (datetime, date)):
        return val if isinstance(val, datetime) else datetime.combine(val, datetime.min.time())
    s = str(val).strip().replace('T', ' ')
    for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%Y-%m-%d',
                '%d-%m-%Y %H:%M:%S', '%d-%m-%Y %H:%M', '%d-%m-%Y',
                '%d/%m/%Y %H:%M:%S', '%d/%m/%Y %H:%M', '%d/%m/%Y'):
        try:
            return datetime.strptime(s[:len(fmt.replace('%Y', '2026').replace('%m', '12').replace('%d', '12').replace('%H', '12').replace('%M', '12').replace('%S', '12'))], fmt)
        except Exception:
            pass
    try:
        return datetime.fromisoformat(s)
    except Exception:
        return None


def _get_fy_start_year(fin_year: str) -> int:
    try:
        return int(fin_year.split('-')[0])
    except Exception:
        return datetime.now().year if datetime.now().month >= 4 else datetime.now().year - 1


def _period_bounds(fin_year: str, month: str):
    """Return (start_dt, end_dt, is_all_months) as datetime objects."""
    start_y = _get_fy_start_year(fin_year)
    if not month or month.lower() == 'all':
        return datetime(start_y, 4, 1, 0, 0, 0), datetime(start_y + 1, 4, 1, 0, 0, 0), True
    
    m_num = MONTH_MAP.get(month.capitalize())
    if not m_num:
        for k, v in MONTH_SHORT_MAP.items():
            if month.lower().startswith(k.lower()):
                m_num = v
                break
    if not m_num:
        m_num = 4
    
    y = start_y if m_num >= 4 else start_y + 1
    start_dt = datetime(y, m_num, 1, 0, 0, 0)
    if m_num == 12:
        end_dt = datetime(y + 1, 1, 1, 0, 0, 0)
    else:
        end_dt = datetime(y, m_num + 1, 1, 0, 0, 0)
    return start_dt, end_dt, False


def _load_masters(cur):
    """Load master lookup tables for validation and mapping."""
    # 1. Vessel Flag Master
    cur.execute("SELECT name, flag_type FROM vessel_flags WHERE name IS NOT NULL AND name <> ''")
    flag_master = {r['name'].strip().upper(): (r['flag_type'] or '').strip().upper() for r in cur.fetchall()}

    # 2. Terminal Master
    cur.execute("SELECT terminal_name FROM terminal_master WHERE is_active IS NOT FALSE")
    terminal_master = {r['terminal_name'].strip().upper() for r in cur.fetchall() if r['terminal_name']}

    # 3. Pipeline Master
    cur.execute("SELECT pipeline_name FROM pipeline_master WHERE is_active IS NOT FALSE")
    pipeline_master = {r['pipeline_name'].strip().upper() for r in cur.fetchall() if r['pipeline_name']}

    # 4. Vessel Agent Master
    cur.execute("SELECT name, agent_code FROM vessel_agents WHERE COALESCE(is_active, 1) <> 0")
    agent_master = {}
    for r in cur.fetchall():
        if r['name']:
            agent_master[r['name'].strip().upper()] = r['name'].strip()
        if r['agent_code']:
            agent_master[r['agent_code'].strip().upper()] = r['name'].strip() if r['name'] else r['agent_code'].strip()

    # 5. Vessel Run Type Master
    cur.execute("SELECT name FROM vessel_run_types")
    run_type_master = {r['name'].strip().upper() for r in cur.fetchall() if r['name']}

    # 6. Port Master (VPM01)
    cur.execute("SELECT name, port_code FROM port_master WHERE name IS NOT NULL AND TRIM(name) <> ''")
    port_by_name = {}
    port_by_code = {}
    port_canonical_name = {}
    port_list = []
    for r in cur.fetchall():
        p_name = (r['name'] or '').strip()
        p_code = (r['port_code'] or '').strip()
        if p_name:
            p_upper = p_name.upper()
            port_by_name[p_upper] = p_code
            port_canonical_name[p_upper] = p_name
            port_list.append({'name': p_name, 'name_upper': p_upper, 'code': p_code})
        if p_code:
            port_by_code[p_code.upper()] = p_name

    return {
        'flags': flag_master,
        'terminals': terminal_master,
        'pipelines': pipeline_master,
        'agents': agent_master,
        'run_types': run_type_master,
        'port_by_name': port_by_name,
        'port_by_code': port_by_code,
        'port_canonical_name': port_canonical_name,
        'port_list': port_list
    }


def _resolve_port(raw_port: str, raw_code: str, masters: dict) -> tuple:
    """
    Dynamically resolve Port Code and Port Name against Port Master (VPM01).
    Prioritizes vessel Load Port and dynamically looks up registered Port Code.
    Handles exact name match, exact code match, and flexible prefix/suffix matching.
    """
    raw_port = (raw_port or '').strip()
    raw_code = (raw_code or '').strip()

    if not raw_port and not raw_code:
        return 'Missing Port Code', 'Missing Port Name'

    # If code was already provided (e.g. from historical mis_vessel_master)
    if raw_code and raw_code.upper() in masters['port_by_code']:
        return raw_code.upper(), masters['port_by_code'][raw_code.upper()]

    target_name = raw_port or raw_code
    target_upper = target_name.upper()

    # 1. Exact Port Name match in Port Master
    if target_upper in masters['port_by_name']:
        matched_code = masters['port_by_name'][target_upper]
        canonical = masters.get('port_canonical_name', {}).get(target_upper, target_name)
        return (matched_code if matched_code else canonical), canonical

    # 2. Check if target_name is already a Port Code in Port Master
    if target_upper in masters['port_by_code']:
        return target_upper, masters['port_by_code'][target_upper]

    # 3. Flexible / Partial matching against Port Master (e.g., 'Rosario, Argentina' -> 'Rosario')
    target_clean = re.sub(r'[\.,;].*$', '', target_name).strip().upper()
    for p in masters.get('port_list', []):
        p_name_up = p['name_upper']
        p_base = re.sub(r'[\.,;].*$', '', p_name_up).strip()
        if (target_clean and target_clean == p_base) or \
           (len(p_base) >= 4 and (p_base in target_upper or target_clean in p_name_up)):
            matched_code = p['code'] if p['code'] else p['name']
            return matched_code, p['name']

    # 4. Dynamic fallback when port name is not yet registered in Port Master
    code_val = raw_code or target_name
    return code_val, target_name


def _fetch_live_data(cur, start_dt, end_dt, masters):
    """
    Fetch live operational records (post-cutover: ldud_header, vcn_header,
    ldud_parcel_ops, lueu_parcel_log, vcn_consigners, vcn_export_cargo_declaration).
    """
    cur.execute("""
        SELECT
            lh.id AS ldud_id,
            lh.doc_num AS ldud_doc_num,
            lh.cast_off_datetime,
            lh.discharge_completed,
            lh.alongside_datetime,
            vh.id AS vcn_id,
            vh.vcn_doc_num,
            vh.via_number,
            vh.vessel_name,
            vh.vessel_master_doc,
            vh.vessel_agent_name,
            vh.vessel_run_type,
            vh.berth_name,
            vh.loa,
            vh.draft,
            vh.pbl,
            vh.operation_type,
            vh.load_port,
            vh.discharge_port,
            ves.nationality AS vessel_nationality,
            ves.doc_num AS vessel_doc_num
        FROM ldud_header lh
        JOIN vcn_header vh ON vh.id = lh.vcn_id
        LEFT JOIN vessels ves ON (
            ves.doc_num = split_part(COALESCE(vh.vessel_master_doc, ''), '/', 1)
            OR UPPER(REPLACE(TRIM(ves.vessel_name), 'MT ', '')) = UPPER(REPLACE(TRIM(vh.vessel_name), 'MT ', ''))
        )
        WHERE COALESCE(lh.is_deleted, FALSE) = FALSE
          AND (
              (lh.cast_off_datetime IS NOT NULL AND NULLIF(TRIM(lh.cast_off_datetime), '') IS NOT NULL)
              OR
              (lh.discharge_completed IS NOT NULL AND NULLIF(TRIM(lh.discharge_completed), '') IS NOT NULL)
          )
        ORDER BY lh.cast_off_datetime NULLS LAST, lh.discharge_completed NULLS LAST, vh.vcn_doc_num
    """)
    vessel_calls = cur.fetchall()

    records = []
    seen_calls = set()

    for vc in vessel_calls:
        dt_val = _parse_dt(vc['cast_off_datetime']) or _parse_dt(vc['discharge_completed'])
        if not dt_val:
            continue
        if dt_val < start_dt or dt_val >= end_dt:
            continue

        ldud_id = vc['ldud_id']
        seen_calls.add(ldud_id)

        cast_off_date = dt_val.strftime('%d-%m-%Y')
        cast_off_time = dt_val.strftime('%H:%M')

        vcn_num = vc['vcn_doc_num'] or vc['via_number'] or 'Missing VCN'
        vessel_name = vc['vessel_name'] or 'Missing Vessel'
        agent_name = (vc['vessel_agent_name'] or '').strip()
        run_type = (vc['vessel_run_type'] or '').strip()
        flag_name = (vc['vessel_nationality'] or '').strip()
        op_type = (vc['operation_type'] or '').strip().capitalize()

        # Prioritize vessel Load Port as requested, with fallback to discharge port
        load_port_val = (vc['load_port'] or '').strip()
        if not load_port_val:
            load_port_val = (vc['discharge_port'] or '').strip()

        port_code, port_name = _resolve_port(load_port_val, '', masters)

        cur.execute("""
            SELECT
                po.id AS po_id,
                po.parcel_ids,
                po.cargo_name,
                po.terminal_name,
                po.quantity AS po_quantity
            FROM ldud_parcel_ops po
            WHERE po.ldud_id = %s
            ORDER BY po.id
        """, [ldud_id])
        parcel_ops = cur.fetchall()

        if not parcel_ops:
            records.append({
                'source': 'Live',
                'vcn_no': vcn_num,
                'vessel_name': vessel_name,
                'terminal': 'Missing Terminal',
                'pipeline': 'Missing Pipeline',
                'cargo': 'Missing Cargo',
                'quantity_mt': 0.0,
                'short_close_qty': 0.0,
                'agent_name': agent_name,
                'run_type': run_type,
                'flag_name': flag_name,
                'port_code': port_code,
                'port_name': port_name,
                'cast_off_date': cast_off_date,
                'cast_off_time': cast_off_time,
                'has_parcels': False
            })
            continue

        for po in parcel_ops:
            po_id = po['po_id']
            terminal = (po['terminal_name'] or '').strip()
            cargo = (po['cargo_name'] or '').strip()

            parcel_ids = [int(x.strip()) for x in str(po['parcel_ids'] or '').split(',') if x.strip().isdigit()]
            tbl = 'vcn_export_cargo_declaration' if op_type == 'Export' else 'vcn_consigners'
            pipeline = ''
            if parcel_ids:
                cur.execute(f"""
                    SELECT pipeline_name, unload_terminal
                    FROM {tbl}
                    WHERE id = ANY(%s)
                """, [parcel_ids])
                p_rows = cur.fetchall()
                pipes = list(dict.fromkeys(r['pipeline_name'].strip() for r in p_rows if r['pipeline_name'] and r['pipeline_name'].strip()))
                pipeline = ', '.join(pipes)
                if not terminal:
                    terms = list(dict.fromkeys(r['unload_terminal'].strip() for r in p_rows if r['unload_terminal'] and r['unload_terminal'].strip()))
                    terminal = ', '.join(terms)

            # Query actual handled quantity excluding Short Close
            cur.execute("""
                SELECT
                    COALESCE(SUM(quantity), 0) AS handled_qty,
                    COALESCE(SUM(CASE
                        WHEN COALESCE(is_shortclose, FALSE) = TRUE
                          OR LOWER(COALESCE(remarks, '')) LIKE '%%short%%close%%'
                        THEN quantity ELSE 0 END
                    ), 0) AS sc_qty,
                    COUNT(*) AS log_count
                FROM lueu_parcel_log
                WHERE parcel_op_id = %s
                  AND is_deleted IS NOT TRUE
            """, [po_id])
            log_res = cur.fetchone()

            log_count = log_res['log_count'] if log_res else 0
            if log_count > 0:
                handled_qty = float(log_res['handled_qty'] or 0.0) - float(log_res['sc_qty'] or 0.0)
                sc_qty = float(log_res['sc_qty'] or 0.0)
            else:
                handled_qty = float(po['po_quantity'] or 0.0)
                sc_qty = 0.0

            records.append({
                'source': 'Live',
                'vcn_no': vcn_num,
                'vessel_name': vessel_name,
                'terminal': terminal,
                'pipeline': pipeline,
                'cargo': cargo,
                'quantity_mt': round(max(handled_qty, 0.0), 3),
                'short_close_qty': round(max(sc_qty, 0.0), 3),
                'agent_name': agent_name,
                'run_type': run_type,
                'flag_name': flag_name,
                'port_code': port_code,
                'port_name': port_name,
                'cast_off_date': cast_off_date,
                'cast_off_time': cast_off_time,
                'has_parcels': True
            })

    return records


def _fetch_legacy_data(cur, fin_year, month, start_dt, end_dt, masters):
    """
    Fetch historical records from mis_vessel_master and mis_history.
    """
    cur.execute("""
        SELECT
            mvm.vcn_no,
            mvm.vessel_name,
            mvm.agent,
            mvm.overseas_coastal,
            mvm.foreign_indian,
            mvm.flag,
            mvm.port_code,
            mvm.port_of_loading,
            mvm.unload_pipeline,
            mvm.unloading_terminal,
            mvm.cargo,
            mvm.quantity,
            mvm.cast_off,
            mvm.sail_cast_off,
            mvm.cargo_completion,
            mvm.fin_year,
            mvm.month
        FROM mis_vessel_master mvm
        WHERE mvm.fin_year = %s
        ORDER BY mvm.id
    """, [fin_year])
    all_vessels = cur.fetchall()

    records = []
    for mvm in all_vessels:
        dt_val = _parse_dt(mvm['cast_off']) or _parse_dt(mvm['sail_cast_off']) or _parse_dt(mvm['cargo_completion'])
        if dt_val:
            if dt_val < start_dt or dt_val >= end_dt:
                continue
            cast_off_date = dt_val.strftime('%d-%m-%Y')
            cast_off_time = dt_val.strftime('%H:%M')
        else:
            m_text = str(mvm['month'] or '').strip()
            if month and month.lower() != 'all':
                m_short = month[:3].lower()
                if m_short not in m_text.lower():
                    continue
            cast_off_date = 'Missing Cast-Off Date'
            cast_off_time = ''

        vcn_num = (mvm['vcn_no'] or '').strip() or 'Missing VCN'
        vessel_name = (mvm['vessel_name'] or '').strip() or 'Missing Vessel'
        agent_name = (mvm['agent'] or '').strip()
        run_type = (mvm['overseas_coastal'] or '').strip()
        if not run_type:
            fi = (mvm['foreign_indian'] or '').strip().upper()
            if fi in ('F', 'FOREIGN'):
                run_type = 'Foreign'
            elif fi in ('I', 'INDIAN', 'C', 'COASTAL', 'COSTAL'):
                run_type = 'Costal'
        flag_name = (mvm['flag'] or '').strip()
        port_code_raw = (mvm['port_code'] or '').strip()
        load_port_raw = (mvm['port_of_loading'] or '').strip()
        port_code, port_name = _resolve_port(load_port_raw, port_code_raw, masters)
        default_pipe = (mvm['unload_pipeline'] or '').strip()

        cur.execute("""
            SELECT
                terminal,
                cargo_name,
                quantity,
                customer,
                payment_by
            FROM mis_history
            WHERE vcn_no = %s
            ORDER BY id
        """, [vcn_num])
        hist_parcels = cur.fetchall()

        if hist_parcels:
            for hp in hist_parcels:
                t_val = (hp['terminal'] or '').strip()
                c_val = (hp['cargo_name'] or '').strip()
                q_val = float(hp['quantity'] or 0.0)
                records.append({
                    'source': 'Historical',
                    'vcn_no': vcn_num,
                    'vessel_name': vessel_name,
                    'terminal': t_val,
                    'pipeline': default_pipe,
                    'cargo': c_val,
                    'quantity_mt': round(max(q_val, 0.0), 3),
                    'short_close_qty': 0.0,
                    'agent_name': agent_name,
                    'run_type': run_type,
                    'flag_name': flag_name,
                    'port_code': port_code,
                    'port_name': port_name,
                    'cast_off_date': cast_off_date,
                    'cast_off_time': cast_off_time,
                    'has_parcels': True
                })
        else:
            t_val = (mvm['unloading_terminal'] or '').strip()
            c_val = (mvm['cargo'] or '').strip()
            q_val = float(mvm['quantity'] or 0.0)
            records.append({
                'source': 'Historical',
                'vcn_no': vcn_num,
                'vessel_name': vessel_name,
                'terminal': t_val,
                'pipeline': default_pipe,
                'cargo': c_val,
                'quantity_mt': round(max(q_val, 0.0), 3),
                'short_close_qty': 0.0,
                'agent_name': agent_name,
                'run_type': run_type,
                'flag_name': flag_name,
                'port_code': port_code,
                'port_name': port_name,
                'cast_off_date': cast_off_date,
                'cast_off_time': cast_off_time,
                'has_parcels': bool(t_val or c_val or q_val)
            })

    return records


def get_other_statistics_data(fin_year: str, month: str):
    start_dt, end_dt, is_all = _period_bounds(fin_year, month)

    conn = get_db()
    try:
        cur = get_cursor(conn)
        masters = _load_masters(cur)

        raw_records = []
        if end_dt <= datetime.combine(CUTOFF_DATE, datetime.min.time()):
            raw_records = _fetch_legacy_data(cur, fin_year, month, start_dt, end_dt, masters)
        elif start_dt >= datetime.combine(CUTOFF_DATE, datetime.min.time()):
            raw_records = _fetch_live_data(cur, start_dt, end_dt, masters)
        else:
            legacy_end = datetime.combine(CUTOFF_DATE, datetime.min.time())
            rec_leg = _fetch_legacy_data(cur, fin_year, month, start_dt, legacy_end, masters)
            rec_live = _fetch_live_data(cur, legacy_end, end_dt, masters)
            raw_records = rec_leg + rec_live
    finally:
        conn.close()

    processed_records = []
    validation_errors = []
    seen_vcns = set()

    for r in raw_records:
        vcn = r['vcn_no']
        vessel = r['vessel_name']
        seen_vcns.add(vcn)

        # 1. Terminal verification
        terminal = r['terminal']
        if not terminal:
            terminal = 'Missing Terminal'
            validation_errors.append({'vcn': vcn, 'vessel': vessel, 'field': 'Terminal', 'issue': 'Terminal is missing in parcel data'})
        else:
            for t in [x.strip() for x in terminal.split(',') if x.strip()]:
                if t.upper() not in masters['terminals']:
                    validation_errors.append({'vcn': vcn, 'vessel': vessel, 'field': 'Terminal', 'issue': f'Terminal "{t}" not found in Terminal Master'})

        # 2. Pipeline verification
        pipeline = r['pipeline']
        if not pipeline:
            pipeline = 'Missing Pipeline'
            validation_errors.append({'vcn': vcn, 'vessel': vessel, 'field': 'Pipeline', 'issue': 'Pipeline is missing in parcel data'})
        else:
            for p in [x.strip() for x in pipeline.split(',') if x.strip()]:
                if p.upper() not in masters['pipelines']:
                    validation_errors.append({'vcn': vcn, 'vessel': vessel, 'field': 'Pipeline', 'issue': f'Pipeline "{p}" not found in Pipeline Master'})

        # 3. Cargo verification
        cargo = r['cargo']
        if not cargo:
            cargo = 'Missing Cargo'
            validation_errors.append({'vcn': vcn, 'vessel': vessel, 'field': 'Cargo', 'issue': 'Cargo name is missing'})

        # 4. Quantity verification
        qty = float(r['quantity_mt'] or 0.0)
        if qty <= 0.0:
            validation_errors.append({'vcn': vcn, 'vessel': vessel, 'field': 'Quantity', 'issue': 'Handled quantity is 0 or missing'})

        # 5. Vessel Agent verification
        agent = r['agent_name']
        if not agent:
            agent = 'Missing Agent'
            validation_errors.append({'vcn': vcn, 'vessel': vessel, 'field': 'Agent', 'issue': 'Vessel agent is missing'})
        elif agent.upper() not in masters['agents']:
            validation_errors.append({'vcn': vcn, 'vessel': vessel, 'field': 'Agent', 'issue': f'Agent "{agent}" not found in Vessel Agent Master'})

        # 6. Vessel Run Type verification
        raw_run_type = (r['run_type'] or '').strip()
        run_type_lower = raw_run_type.lower()
        if 'fore' in run_type_lower or 'over' in run_type_lower or run_type_lower == 'f':
            run_type = 'Foreign'
        elif 'cost' in run_type_lower or 'coast' in run_type_lower or 'ind' in run_type_lower or run_type_lower in ('c', 'i'):
            run_type = 'Costal'
        else:
            run_type = raw_run_type

        if not run_type:
            run_type = 'Missing Vessel Run Type'
            validation_errors.append({'vcn': vcn, 'vessel': vessel, 'field': 'Vessel Run Type', 'issue': 'Vessel run type is missing'})
        elif run_type.upper() not in masters['run_types']:
            validation_errors.append({'vcn': vcn, 'vessel': vessel, 'field': 'Vessel Run Type', 'issue': f'Run type "{run_type}" not recognized in Master'})

        # 7. Flag and Flag Type verification
        # User requirement: If vessel run type is Foreign ("foren") show 'FF',
        # and if vessel run type is Coastal ("costel") show 'IF'.
        if run_type == 'Foreign' or 'fore' in run_type_lower or 'over' in run_type_lower or run_type_lower == 'f':
            flag_type = 'FF'
        elif run_type == 'Costal' or 'cost' in run_type_lower or 'coast' in run_type_lower or 'ind' in run_type_lower or run_type_lower in ('c', 'i'):
            flag_type = 'IF'
        else:
            flag_type = masters['flags'].get((r['flag_name'] or '').strip().upper(), '')

        flag_name = (r['flag_name'] or '').strip()
        if not flag_name or flag_name.lower() in ('missing flag', 'none', 'null'):
            if flag_type == 'FF':
                flag_name = 'Foreign'
            elif flag_type == 'IF':
                flag_name = 'India'
            else:
                flag_name = 'Missing Flag'
                if not flag_type:
                    flag_type = 'Missing Flag'
                validation_errors.append({'vcn': vcn, 'vessel': vessel, 'field': 'Flag', 'issue': 'Vessel flag / nationality is missing'})
        else:
            if not flag_type:
                flag_type = masters['flags'].get(flag_name.upper(), 'Missing Flag')

        # 8. Port verification (resolving Load Port dynamically against Port Master)
        port_code, port_name = _resolve_port(r['port_name'], r['port_code'], masters)
        if not port_code or port_code == 'Missing Port Code':
            port_code = 'Missing Port Code'
            port_name = 'Missing Port Name'
            validation_errors.append({'vcn': vcn, 'vessel': vessel, 'field': 'Port', 'issue': 'Both Port Code and Port Name are missing in vessel load port'})
        elif port_name.upper() not in masters['port_by_name'] and port_code.upper() not in masters['port_by_code']:
            validation_errors.append({'vcn': vcn, 'vessel': vessel, 'field': 'Port Code', 'issue': f'Load Port "{port_name}" not registered in Port Master (VPM01)'})

        # 9. Cast-off date verification
        c_date = (r['cast_off_date'] or '').strip()
        c_time = (r['cast_off_time'] or '').strip()
        if not c_date or 'Missing' in c_date:
            validation_errors.append({'vcn': vcn, 'vessel': vessel, 'field': 'Cast-Off', 'issue': 'Cast-Off / Completion date-time is missing'})
            last_cast_off = c_date or 'Missing Cast-Off Date'
        else:
            last_cast_off = f"{c_date} {c_time}".strip()

        processed_records.append({
            'vcn': vcn,
            'vessel_name': vessel,
            'terminal': terminal,
            'pipeline': pipeline,
            'cargo': cargo,
            'qty_mt': qty,
            'short_close_qty': r['short_close_qty'],
            'agent_name': agent,
            'run_type': run_type,
            'flag_name': flag_name,
            'flag_type': flag_type,
            'port_code': port_code,
            'port_name': port_name,
            'cast_off_date': r['cast_off_date'],
            'cast_off_time': r['cast_off_time'],
            'last_cast_off': last_cast_off,
        })

    # Summary 1: Terminal and Pipeline Wise
    t_p_c_agg = {}
    for r in processed_records:
        key = (r['terminal'], r['pipeline'], r['cargo'])
        t_p_c_agg[key] = t_p_c_agg.get(key, 0.0) + r['qty_mt']
    summary_terminal_pipeline_cargo = [
        {'terminal': k[0], 'pipeline': k[1], 'cargo': k[2], 'qty_mt': round(v, 3)}
        for k, v in sorted(t_p_c_agg.items(), key=lambda x: (x[0][0], x[0][1], x[0][2]))
    ]

    # Summary 2: Vessel Agent Wise
    agent_agg = {}
    for r in processed_records:
        key = r['agent_name']
        agent_agg[key] = agent_agg.get(key, 0.0) + r['qty_mt']
    summary_agent = [
        {'agent_name': k, 'qty_mt': round(v, 3)}
        for k, v in sorted(agent_agg.items(), key=lambda x: x[0])
    ]

    # Summary 3: Flag Wise
    flag_agg = {}
    for r in processed_records:
        key = (r['flag_type'], r['flag_name'])
        flag_agg[key] = flag_agg.get(key, 0.0) + r['qty_mt']
    summary_flag = [
        {'flag_type': k[0], 'flag_name': k[1], 'qty_mt': round(v, 3)}
        for k, v in sorted(flag_agg.items(), key=lambda x: (x[0][0], x[0][1]))
    ]

    # Summary 4: Port Wise
    port_agg = {}
    for r in processed_records:
        key = (r['port_code'], r['port_name'])
        port_agg[key] = port_agg.get(key, 0.0) + r['qty_mt']
    summary_port = [
        {'port_code': k[0], 'port_name': k[1], 'qty_mt': round(v, 3)}
        for k, v in sorted(port_agg.items(), key=lambda x: (x[0][0], x[0][1]))
    ]

    # Supplementary: Vessel Run Type Wise
    run_type_agg = {}
    for r in processed_records:
        key = r['run_type']
        run_type_agg[key] = run_type_agg.get(key, 0.0) + r['qty_mt']
    summary_run_type = [
        {'run_type': k, 'qty_mt': round(v, 3)}
        for k, v in sorted(run_type_agg.items(), key=lambda x: x[0])
    ]

    total_qty = sum(r['qty_mt'] for r in processed_records)
    total_short_close = sum(r['short_close_qty'] for r in processed_records)

    return {
        'fin_year': fin_year,
        'month': month,
        'summary': {
            'total_vessels': len(seen_vcns),
            'total_parcels': len(processed_records),
            'total_qty_mt': round(total_qty, 3),
            'short_close_qty_mt': round(total_short_close, 3),
            'error_count': len(validation_errors)
        },
        'terminal_pipeline_cargo': summary_terminal_pipeline_cargo,
        'vessel_agent': summary_agent,
        'flag_wise': summary_flag,
        'port_wise': summary_port,
        'run_type_wise': summary_run_type,
        'vessel_detail': processed_records,
        'validation_errors': validation_errors
    }


def _get_available_fin_years():
    conn = get_db()
    cur = get_cursor(conn)
    cur.execute("SELECT DISTINCT fin_year FROM mis_vessel_master WHERE fin_year IS NOT NULL AND fin_year <> ''")
    years = {r['fin_year'].strip() for r in cur.fetchall()}
    conn.close()
    years.add('2024-25')
    years.add('2025-26')
    years.add('2026-27')
    return sorted(list(years), reverse=True)


# ══════════════════════════════════════════════════════════════════
#  ROUTES
# ══════════════════════════════════════════════════════════════════

@bp.route('/module/RP01/statistics-report/')
@bp.route('/module/RP01/other-statistics/')
@login_required
def statistics_report_index():
    perms = get_perms()
    if not perms.get('can_read'):
        return render_template('no_access.html'), 403

    today = date.today()
    current_month_name = today.strftime('%B')
    current_fy = f"{today.year}-{str(today.year + 1)[-2:]}" if today.month >= 4 else f"{today.year - 1}-{str(today.year)[-2:]}"

    fin_years = _get_available_fin_years()
    if current_fy not in fin_years:
        fin_years.insert(0, current_fy)
    default_fy = current_fy
    default_month = current_month_name if current_month_name in MONTH_NAMES else 'April'

    return render_template(
        'statistics_report/statistics_report.html',
        username=session.get('username'),
        permissions=perms,
        fin_years=fin_years,
        month_names=MONTH_NAMES,
        default_fy=default_fy,
        default_month=default_month
    )


@bp.route('/api/module/RP01/statistics-report/data', methods=['GET'])
@bp.route('/api/module/RP01/other-statistics/data', methods=['GET'])
@login_required
def statistics_report_api_data():
    perms = get_perms()
    if not perms.get('can_read'):
        return jsonify({'error': 'Unauthorized'}), 403

    today = date.today()
    cur_month = today.strftime('%B')
    cur_fy = f"{today.year}-{str(today.year + 1)[-2:]}" if today.month >= 4 else f"{today.year - 1}-{str(today.year)[-2:]}"

    fin_year = request.args.get('fin_year', cur_fy).strip()
    month = request.args.get('month', cur_month).strip()

    try:
        data = get_other_statistics_data(fin_year, month)
        return jsonify(data)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@bp.route('/api/module/RP01/statistics-report/export', methods=['GET'])
@bp.route('/api/module/RP01/other-statistics/export', methods=['GET'])
@login_required
def statistics_report_api_export():
    perms = get_perms()
    if not perms.get('can_read'):
        return jsonify({'error': 'Unauthorized'}), 403

    today = date.today()
    cur_month = today.strftime('%B')
    cur_fy = f"{today.year}-{str(today.year + 1)[-2:]}" if today.month >= 4 else f"{today.year - 1}-{str(today.year)[-2:]}"

    fin_year = request.args.get('fin_year', cur_fy).strip()
    month = request.args.get('month', cur_month).strip()

    try:
        data = get_other_statistics_data(fin_year, month)
    except Exception as e:
        return jsonify({'error': f'Export failed: {e}'}), 500

    wb = Workbook()
    wb.remove(wb.active)

    font_title = Font(name="Calibri", size=14, bold=True, color="1E3A8A")
    font_sub = Font(name="Calibri", size=10, italic=True, color="4B5563")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_data = Font(name="Calibri", size=11)
    font_bold = Font(name="Calibri", size=11, bold=True)
    font_total = Font(name="Calibri", size=11, bold=True, color="1E3A8A")

    fill_header = PatternFill("solid", fgColor="1E3A8A")
    fill_total = PatternFill("solid", fgColor="E0E7FF")
    fill_error = PatternFill("solid", fgColor="FEE2E2")

    thin_border = Border(
        left=Side(style="thin", color="D1D5DB"),
        right=Side(style="thin", color="D1D5DB"),
        top=Side(style="thin", color="D1D5DB"),
        bottom=Side(style="thin", color="D1D5DB")
    )
    total_border = Border(
        left=Side(style="thin", color="D1D5DB"),
        right=Side(style="thin", color="D1D5DB"),
        top=Side(style="thin", color="1E3A8A"),
        bottom=Side(style="double", color="1E3A8A")
    )

    def style_header(ws, row_idx, num_cols):
        ws.row_dimensions[row_idx].height = 26
        for c in range(1, num_cols + 1):
            cell = ws.cell(row=row_idx, column=c)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border

    def auto_fit_columns(ws, max_cols):
        for col_idx in range(1, max_cols + 1):
            max_len = 0
            col_letter = get_column_letter(col_idx)
            for row in ws.iter_rows(min_col=col_idx, max_col=col_idx):
                val = row[0].value
                if val is not None:
                    s_len = len(str(val))
                    if s_len > max_len:
                        max_len = s_len
            ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

    # ─────────────────────────────────────────────────────────────
    # SHEET 1: Terminal_Pipeline_Cargo
    # ─────────────────────────────────────────────────────────────
    ws1 = wb.create_sheet(title="Terminal_Pipeline_Cargo")


    headers1 = ["Terminal", "Pipeline", "Cargo", "Qty Handled in MT"]
    for i, h in enumerate(headers1, 1):
        ws1.cell(row=4, column=i, value=h)
    style_header(ws1, 4, len(headers1))

    cur_row = 5
    tot_qty1 = 0.0
    for item in data['terminal_pipeline_cargo']:
        ws1.cell(row=cur_row, column=1, value=item['terminal']).font = font_data
        ws1.cell(row=cur_row, column=2, value=item['pipeline']).font = font_data
        ws1.cell(row=cur_row, column=3, value=item['cargo']).font = font_data
        c4 = ws1.cell(row=cur_row, column=4, value=item['qty_mt'])
        c4.font = font_data
        c4.number_format = "#,##0.000"
        for c in range(1, 5):
            ws1.cell(row=cur_row, column=c).border = thin_border
        tot_qty1 += item['qty_mt']
        cur_row += 1

    ws1.cell(row=cur_row, column=1, value="Total").font = font_total
    ws1.cell(row=cur_row, column=2, value="").font = font_total
    ws1.cell(row=cur_row, column=3, value="").font = font_total
    c_tot = ws1.cell(row=cur_row, column=4, value=round(tot_qty1, 3))
    c_tot.font = font_total
    c_tot.number_format = "#,##0.000"
    for c in range(1, 5):
        ws1.cell(row=cur_row, column=c).fill = fill_total
        ws1.cell(row=cur_row, column=c).border = total_border
    auto_fit_columns(ws1, 4)

    # ─────────────────────────────────────────────────────────────
    # SHEET 2: Vessel_Agent
    # ─────────────────────────────────────────────────────────────
    ws2 = wb.create_sheet(title="Vessel_Agent")


    headers2 = ["Vessel Agent Name", "Qty Handled in MT"]
    for i, h in enumerate(headers2, 1):
        ws2.cell(row=4, column=i, value=h)
    style_header(ws2, 4, len(headers2))

    cur_row = 5
    tot_qty2 = 0.0
    for item in data['vessel_agent']:
        ws2.cell(row=cur_row, column=1, value=item['agent_name']).font = font_data
        c2 = ws2.cell(row=cur_row, column=2, value=item['qty_mt'])
        c2.font = font_data
        c2.number_format = "#,##0.000"
        for c in range(1, 3):
            ws2.cell(row=cur_row, column=c).border = thin_border
        tot_qty2 += item['qty_mt']
        cur_row += 1

    ws2.cell(row=cur_row, column=1, value="Total").font = font_total
    c_tot2 = ws2.cell(row=cur_row, column=2, value=round(tot_qty2, 3))
    c_tot2.font = font_total
    c_tot2.number_format = "#,##0.000"
    for c in range(1, 3):
        ws2.cell(row=cur_row, column=c).fill = fill_total
        ws2.cell(row=cur_row, column=c).border = total_border
    auto_fit_columns(ws2, 2)

    # ─────────────────────────────────────────────────────────────
    # SHEET 3: Flag_Wise
    # ─────────────────────────────────────────────────────────────
    ws3 = wb.create_sheet(title="Flag_Wise")

    headers3 = ["Flag Type", "Flag Name", "Qty Handled in MT"]
    for i, h in enumerate(headers3, 1):
        ws3.cell(row=4, column=i, value=h)
    style_header(ws3, 4, len(headers3))

    cur_row = 5
    tot_qty3 = 0.0
    for item in data['flag_wise']:
        ws3.cell(row=cur_row, column=1, value=item['flag_type']).font = font_data
        ws3.cell(row=cur_row, column=2, value=item['flag_name']).font = font_data
        c3 = ws3.cell(row=cur_row, column=3, value=item['qty_mt'])
        c3.font = font_data
        c3.number_format = "#,##0.000"
        for c in range(1, 4):
            ws3.cell(row=cur_row, column=c).border = thin_border
        tot_qty3 += item['qty_mt']
        cur_row += 1

    ws3.cell(row=cur_row, column=1, value="Total").font = font_total
    ws3.cell(row=cur_row, column=2, value="").font = font_total
    c_tot3 = ws3.cell(row=cur_row, column=3, value=round(tot_qty3, 3))
    c_tot3.font = font_total
    c_tot3.number_format = "#,##0.000"
    for c in range(1, 4):
        ws3.cell(row=cur_row, column=c).fill = fill_total
        ws3.cell(row=cur_row, column=c).border = total_border
    auto_fit_columns(ws3, 3)

    # ─────────────────────────────────────────────────────────────
    # SHEET 4: Port_Wise
    # ─────────────────────────────────────────────────────────────
    ws4 = wb.create_sheet(title="Port_Wise")

    headers4 = ["Port Code", "Port Name", "Qty Handled in MT"]
    for i, h in enumerate(headers4, 1):
        ws4.cell(row=4, column=i, value=h)
    style_header(ws4, 4, len(headers4))

    cur_row = 5
    tot_qty4 = 0.0
    for item in data['port_wise']:
        ws4.cell(row=cur_row, column=1, value=item['port_code']).font = font_data
        ws4.cell(row=cur_row, column=2, value=item['port_name']).font = font_data
        c4 = ws4.cell(row=cur_row, column=3, value=item['qty_mt'])
        c4.font = font_data
        c4.number_format = "#,##0.000"
        for c in range(1, 4):
            ws4.cell(row=cur_row, column=c).border = thin_border
        tot_qty4 += item['qty_mt']
        cur_row += 1

    ws4.cell(row=cur_row, column=1, value="Total").font = font_total
    ws4.cell(row=cur_row, column=2, value="").font = font_total
    c_tot4 = ws4.cell(row=cur_row, column=3, value=round(tot_qty4, 3))
    c_tot4.font = font_total
    c_tot4.number_format = "#,##0.000"
    for c in range(1, 4):
        ws4.cell(row=cur_row, column=c).fill = fill_total
        ws4.cell(row=cur_row, column=c).border = total_border
    auto_fit_columns(ws4, 3)

    # ─────────────────────────────────────────────────────────────
    # SHEET 5: Vessel_Detail
    # ─────────────────────────────────────────────────────────────
    ws5 = wb.create_sheet(title="Vessel_Detail")

    headers5 = [
        "VCN Doc No", "Vessel Name", "Terminal", "Pipeline", "Cargo", "Qty MT",
        "Vessel Agent", "Vessel Run Type", "Flag Name", "Flag Type",
        "Last Cast-Off Date & Time"
    ]
    for i, h in enumerate(headers5, 1):
        ws5.cell(row=4, column=i, value=h)
    style_header(ws5, 4, len(headers5))

    cur_row = 5
    tot_qty5 = 0.0
    for r in data['vessel_detail']:
        ws5.cell(row=cur_row, column=1, value=r['vcn']).font = font_data
        ws5.cell(row=cur_row, column=2, value=r['vessel_name']).font = font_data
        ws5.cell(row=cur_row, column=3, value=r['terminal']).font = font_data
        ws5.cell(row=cur_row, column=4, value=r['pipeline']).font = font_data
        ws5.cell(row=cur_row, column=5, value=r['cargo']).font = font_data
        c6 = ws5.cell(row=cur_row, column=6, value=r['qty_mt'])
        c6.font = font_data
        c6.number_format = "#,##0.000"
        ws5.cell(row=cur_row, column=7, value=r['agent_name']).font = font_data
        ws5.cell(row=cur_row, column=8, value=r['run_type']).font = font_data
        ws5.cell(row=cur_row, column=9, value=r['flag_name']).font = font_data
        ws5.cell(row=cur_row, column=10, value=r['flag_type']).font = font_data
        ws5.cell(row=cur_row, column=11, value=r.get('last_cast_off', '')).font = font_data

        for c in range(1, 12):
            ws5.cell(row=cur_row, column=c).border = thin_border
        tot_qty5 += r['qty_mt']
        cur_row += 1

    ws5.cell(row=cur_row, column=1, value="Total").font = font_total
    for c in range(2, 6):
        ws5.cell(row=cur_row, column=c, value="").font = font_total
    c_tot5 = ws5.cell(row=cur_row, column=6, value=round(tot_qty5, 3))
    c_tot5.font = font_total
    c_tot5.number_format = "#,##0.000"
    for c in range(7, 12):
        ws5.cell(row=cur_row, column=c, value="").font = font_total
    for c in range(1, 12):
        ws5.cell(row=cur_row, column=c).fill = fill_total
        ws5.cell(row=cur_row, column=c).border = total_border
    auto_fit_columns(ws5, 11)

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)

    if not month or month.lower() == 'all':
        filename = f"Other_Statistics_Report_FY_{fin_year}.xlsx"
    else:
        start_y = _get_fy_start_year(fin_year)
        m_num = MONTH_MAP.get(month.capitalize(), 4)
        yr = start_y if m_num >= 4 else start_y + 1
        filename = f"Other_Statistics_Report_{month.capitalize()}_{yr}.xlsx"

    return send_file(
        bio,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename
    )
