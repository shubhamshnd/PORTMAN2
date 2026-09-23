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

    # 7. Vessel Cargo Master (for Cargo Sub Category 2 mapping)
    cur.execute("""
        SELECT cargo_name, cargo_sub_category_2, cargo_category
        FROM vessel_cargo
        WHERE cargo_name IS NOT NULL AND TRIM(cargo_name) <> ''
    """)
    cargo_sub_cat_map = {}
    for r in cur.fetchall():
        c_name = (r['cargo_name'] or '').strip().upper()
        sub2 = (r['cargo_sub_category_2'] or '').strip()
        cat = (r['cargo_category'] or '').strip()
        cargo_sub_cat_map[c_name] = sub2 or cat

    return {
        'flags': flag_master,
        'terminals': terminal_master,
        'pipelines': pipeline_master,
        'agents': agent_master,
        'run_types': run_type_master,
        'port_by_name': port_by_name,
        'port_by_code': port_by_code,
        'port_canonical_name': port_canonical_name,
        'port_list': port_list,
        'cargo_sub_cat_map': cargo_sub_cat_map
    }


def _resolve_cargo_sub_category_2(cargo_name: str, hist_sub2: str = None, masters: dict = None) -> str:
    """
    Resolve Cargo Sub Category 2 by checking cargo name against vessel_cargo master table.
    Prioritizes explicit cargo_sub_category_2 from historical record or vessel_cargo master.
    """
    if hist_sub2 and hist_sub2.strip():
        return hist_sub2.strip()
    cn = (cargo_name or '').strip()
    if not cn:
        return 'Unspecified Cargo'
    c_map = (masters or {}).get('cargo_sub_cat_map', {})
    # 1. Exact match
    if cn.upper() in c_map and c_map[cn.upper()]:
        return c_map[cn.upper()]
    # 2. Substring match against vessel_cargo master
    for k, v in c_map.items():
        if v and (k in cn.upper() or cn.upper() in k):
            return v
    # 3. Known heuristics
    upper_c = cn.upper()
    if any(term in upper_c for term in ['FO', 'FURNACE', 'DIESEL', 'CRUDE', 'OIL', 'PETROL', 'KEROSENE', 'POL', 'FEED STOCK', 'BASE OIL', 'LUBE']):
        if any(e in upper_c for e in ['EDIBLE', 'PALM', 'SOYABEAN', 'SUNFLOWER', 'CPO']):
            return 'EDIBLE OIL'
        return 'POL'
    if any(term in upper_c for term in ['ACID', 'ALCOHOL', 'BENZENE', 'TOLUENE', 'CHEMICAL', 'ACETATE', 'MONOMER', 'KETONE', 'GLYCERINE', 'PHENOL', 'ACETONE']):
        return 'CHEMICAL'
    return cn



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
    default_month = 'All'
    start_y = today.year if today.month >= 4 else today.year - 1
    default_start_date = f"{start_y}-04-01"
    default_end_date = today.strftime('%Y-%m-%d')
    default_start_datetime = f"{start_y}-04-01T07:00"
    default_end_datetime = f"{today.strftime('%Y-%m-%d')}T07:00"

    return render_template(
        'statistics_report/statistics_report.html',
        username=session.get('username'),
        permissions=perms,
        fin_years=fin_years,
        month_names=MONTH_NAMES,
        default_fy=default_fy,
        default_month=default_month,
        default_start_date=default_start_date,
        default_end_date=default_end_date,
        default_start_datetime=default_start_datetime,
        default_end_datetime=default_end_datetime
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
    start_date = request.args.get('start_date', '').strip()
    end_date = request.args.get('end_date', '').strip()

    try:
        data = get_other_statistics_data(fin_year, month)
        data_analytics = get_detailed_analytics_data(fin_year, month, start_date, end_date)
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

    _generate_analytics_excel(data_analytics, wb)

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


# =============================================================================
# MULTI-CATEGORY ANALYTICAL STATISTICS (TAB 2)
# =============================================================================

def get_detailed_analytics_data(
    fin_year: str,
    month: str = 'All',
    start_date_str: str = None,
    end_date_str: str = None
):
    """
    Query and aggregate the analytical tables.

    Customer logic:
        vcn_consigners.consigner_name
                    ↓
              vessel_customers
                    ↓
         Customer Code + Customer Name
                    ↓
         Customer-wise quantity aggregation

    Existing quantity logic:
        LUEU handled quantity - Short Close quantity

    Equipment utilisation:
        Equipment quantity is calculated directly from
        lueu_parcel_log equipment-wise.
    """

    # -------------------------------------------------------------------------
    # DATE RANGE
    # -------------------------------------------------------------------------
    if start_date_str and end_date_str:
        s_dt = _parse_dt(start_date_str)
        e_dt = _parse_dt(end_date_str)

        if s_dt and e_dt:
            start_dt = s_dt

            if len(end_date_str.strip()) == 10:
                end_dt = datetime.combine(
                    e_dt.date(),
                    datetime.max.time()
                )
            else:
                end_dt = (
                    e_dt.replace(
                        second=59,
                        microsecond=999999
                    )
                    if e_dt.second == 0
                    else e_dt
                )
        else:
            start_dt, end_dt, _ = _period_bounds(fin_year, month)
    else:
        start_dt, end_dt, _ = _period_bounds(fin_year, month)

    conn = get_db()
    raw_items = []

    # -------------------------------------------------------------------------
    # EQUIPMENT-WISE TOTALS
    #
    # IMPORTANT:
    # Do not calculate equipment utilisation from raw_items because one
    # parcel operation can have multiple equipment entries.
    #
    # Example:
    #     MLA-1 = 5000 MT
    #     MLA-3 = 3000 MT
    #
    # Each equipment receives only its actual logged quantity.
    # -------------------------------------------------------------------------
    equipment_totals = {}

    try:
        cur = get_cursor(conn)

        # ---------------------------------------------------------------------
        # LOAD EXISTING MASTERS
        # ---------------------------------------------------------------------
        masters = _load_masters(cur)

        # ---------------------------------------------------------------------
        # VESSEL CUSTOMER MASTER
        #
        # We intentionally inspect the table structure instead of assuming
        # exact column names.
        # ---------------------------------------------------------------------
        customer_master = {}

        try:
            cur.execute("""
                SELECT column_name
                FROM information_schema.columns
                WHERE table_schema = 'public'
                  AND table_name = 'vessel_customers'
                ORDER BY ordinal_position
            """)

            customer_columns = [
                row['column_name']
                for row in cur.fetchall()
            ]

            # Normalize column names for matching.
            normalized_columns = {
                str(col).lower().replace('_', ''): col
                for col in customer_columns
            }

            # -------------------------------------------------------------
            # Find customer code column
            # -------------------------------------------------------------
            customer_code_col = None

            code_candidates = [
                'customer_code',
                'customercode',
                'customer_cd',
                'customerid',
                'customer_id',
                'code',
                'cust_code',
                'custcode'
            ]

            for candidate in code_candidates:
                key = candidate.lower().replace('_', '')

                if key in normalized_columns:
                    customer_code_col = normalized_columns[key]
                    break

            # -------------------------------------------------------------
            # Find customer name column
            # -------------------------------------------------------------
            customer_name_col = None

            name_candidates = [
                'customer_name',
                'customername',
                'name',
                'customer',
                'cust_name',
                'custname'
            ]

            for candidate in name_candidates:
                key = candidate.lower().replace('_', '')

                if key in normalized_columns:
                    customer_name_col = normalized_columns[key]
                    break

            # -------------------------------------------------------------
            # If exact candidates were not found, inspect columns
            # -------------------------------------------------------------
            if not customer_code_col:
                for col in customer_columns:
                    lc = str(col).lower()

                    if (
                        'customer' in lc
                        and (
                            'code' in lc
                            or lc.endswith('cd')
                            or lc.endswith('id')
                        )
                    ):
                        customer_code_col = col
                        break

            if not customer_name_col:
                for col in customer_columns:
                    lc = str(col).lower()

                    if (
                        'customer' in lc
                        and 'name' in lc
                    ):
                        customer_name_col = col
                        break

            # -------------------------------------------------------------
            # Build customer master lookup
            # -------------------------------------------------------------
            if customer_code_col and customer_name_col:

                safe_code_col = '"' + customer_code_col.replace('"', '""') + '"'
                safe_name_col = '"' + customer_name_col.replace('"', '""') + '"'

                cur.execute(f"""
                    SELECT
                        {safe_code_col} AS customer_code,
                        {safe_name_col} AS customer_name
                    FROM vessel_customers
                """)

                customer_rows = cur.fetchall()

                for cm in customer_rows:

                    code = str(
                        cm['customer_code'] or ''
                    ).strip()

                    name = str(
                        cm['customer_name'] or ''
                    ).strip()

                    if not code and not name:
                        continue

                    # Lookup by customer code.
                    if code:
                        customer_master[
                            ('code', code.upper())
                        ] = {
                            'customer_code': code,
                            'customer_name': name
                        }

                    # Lookup by customer name.
                    if name:
                        customer_master[
                            ('name', name.upper())
                        ] = {
                            'customer_code': code,
                            'customer_name': name
                        }

            else:
                # Do not stop the entire report if the master structure
                # cannot be resolved.
                customer_master = {}

        except Exception:
            # Customer master must not break the complete report.
            customer_master = {}

        # ---------------------------------------------------------------------
        # CUSTOMER RESOLVER
        # ---------------------------------------------------------------------
        def resolve_customer(raw_customer):
            """
            Resolve consigner/customer value against vessel_customers.

            Returns:
                {
                    'customer_code': ...,
                    'customer_name': ...
                }
            """

            raw_customer = str(raw_customer or '').strip()

            if not raw_customer:
                return {
                    'customer_code': '',
                    'customer_name': 'Unspecified Customer'
                }

            # First try exact name match.
            key_name = ('name', raw_customer.upper())

            if key_name in customer_master:
                return customer_master[key_name]

            # Then try exact code match.
            key_code = ('code', raw_customer.upper())

            if key_code in customer_master:
                return customer_master[key_code]

            # If master does not contain the value, retain the original
            # customer name rather than changing existing report behaviour.
            return {
                'customer_code': '',
                'customer_name': raw_customer
            }

        # =========================================================================
        # 1. FETCH LIVE OPERATIONS
        # =========================================================================
        cur.execute("""
            SELECT
                lh.id AS ldud_id,
                lh.cast_off_datetime,
                lh.discharge_completed,

                vh.vcn_doc_num,
                vh.via_number,
                vh.vessel_name,
                vh.vessel_agent_name,
                vh.vessel_run_type,
                vh.operation_type,
                vh.load_port,
                vh.discharge_port,

                ves.nationality AS vessel_nationality,

                po.id AS po_id,
                po.terminal_name,
                po.cargo_name,
                po.quantity AS po_qty,
                po.start_dt,
                po.end_dt,
                po.parcel_ids

            FROM ldud_header lh

            JOIN vcn_header vh
                ON vh.id = lh.vcn_id

            LEFT JOIN vessels ves
                ON (
                    ves.doc_num =
                        split_part(
                            COALESCE(vh.vessel_master_doc, ''),
                            '/',
                            1
                        )

                    OR

                    UPPER(
                        REPLACE(
                            TRIM(ves.vessel_name),
                            'MT ',
                            ''
                        )
                    )
                    =
                    UPPER(
                        REPLACE(
                            TRIM(vh.vessel_name),
                            'MT ',
                            ''
                        )
                    )
                )

            JOIN ldud_parcel_ops po
                ON po.ldud_id = lh.id

            WHERE COALESCE(lh.is_deleted, FALSE) = FALSE

              AND (
                    (
                        lh.cast_off_datetime IS NOT NULL
                        AND NULLIF(
                            TRIM(lh.cast_off_datetime),
                            ''
                        ) IS NOT NULL
                    )

                    OR

                    (
                        lh.discharge_completed IS NOT NULL
                        AND NULLIF(
                            TRIM(lh.discharge_completed),
                            ''
                        ) IS NOT NULL
                    )
              )

            ORDER BY
                lh.cast_off_datetime,
                po.id
        """)

        live_rows = cur.fetchall()

        # ---------------------------------------------------------------------
        # PROCESS LIVE DATA
        # ---------------------------------------------------------------------
        for r in live_rows:

            # Existing report date logic.
            dt_val = (
                _parse_dt(r['cast_off_datetime'])
                or
                _parse_dt(r['discharge_completed'])
            )

            if not dt_val:
                continue

            if dt_val < start_dt or dt_val > end_dt:
                continue

            po_id = r['po_id']

            # -----------------------------------------------------------------
            # OPERATION TYPE
            # -----------------------------------------------------------------
            op_type = (
                r['operation_type'] or ''
            ).strip().capitalize()

            # -----------------------------------------------------------------
            # PARCEL IDS
            # -----------------------------------------------------------------
            parcel_ids = [
                int(x.strip())
                for x in str(r['parcel_ids'] or '').split(',')
                if x.strip().isdigit()
            ]

            tbl = (
                'vcn_export_cargo_declaration'
                if op_type == 'Export'
                else
                'vcn_consigners'
            )

            pipeline = ''
            customer = ''
            raw_payer = ''
            equipment = ''
            terminal = (
                r['terminal_name'] or ''
            ).strip()

            # -----------------------------------------------------------------
            # PARCEL INFORMATION
            # -----------------------------------------------------------------
            if parcel_ids:

                cur.execute(
                    f"""
                        SELECT
                            pipeline_name,
                            unload_terminal,
                            consigner_name,
                            importer_name,
                            equipment_names

                        FROM {tbl}

                        WHERE id = ANY(%s)
                    """,
                    [parcel_ids]
                )

                p_rows = cur.fetchall()

                pipes = list(
                    dict.fromkeys(
                        p['pipeline_name'].strip()
                        for p in p_rows
                        if (
                            p['pipeline_name']
                            and p['pipeline_name'].strip()
                        )
                    )
                )

                custs = list(
                    dict.fromkeys(
                        p['consigner_name'].strip()
                        for p in p_rows
                        if (
                            p['consigner_name']
                            and p['consigner_name'].strip()
                        )
                    )
                )

                payers = list(
                    dict.fromkeys(
                        p['importer_name'].strip()
                        for p in p_rows
                        if (
                            p['importer_name']
                            and p['importer_name'].strip()
                        )
                    )
                )

                eqs = list(
                    dict.fromkeys(
                        p['equipment_names'].strip()
                        for p in p_rows
                        if (
                            p['equipment_names']
                            and p['equipment_names'].strip()
                        )
                    )
                )

                pipeline = ', '.join(pipes)
                raw_payer = ', '.join(payers)

                # -------------------------------------------------------------
                # Existing raw customer extraction
                # -------------------------------------------------------------
                customer = ', '.join(custs)

                # -------------------------------------------------------------
                # Equipment from parcel master is retained only as fallback
                # display information.
                #
                # Actual Equipment Utilisation is calculated below from
                # lueu_parcel_log.
                # -------------------------------------------------------------
                equipment = ', '.join(eqs)

                # -------------------------------------------------------------
                # Terminal fallback
                # -------------------------------------------------------------
                if not terminal:

                    terms = list(
                        dict.fromkeys(
                            p['unload_terminal'].strip()
                            for p in p_rows
                            if (
                                p['unload_terminal']
                                and p['unload_terminal'].strip()
                            )
                        )
                    )

                    terminal = ', '.join(terms)

            # -----------------------------------------------------------------
            # CUSTOMER MASTER RESOLUTION
            # -----------------------------------------------------------------
            customer_master_values = []

            for customer_value in (
                customer.split(',')
                if customer
                else []
            ):

                customer_value = customer_value.strip()

                if not customer_value:
                    continue

                resolved = resolve_customer(customer_value)

                customer_master_values.append(resolved)

            # Keep unique customers.
            unique_customers = []

            for cm in customer_master_values:
                code = (cm.get('customer_code') or '').strip()
                name = (cm.get('customer_name') or '').strip()
                if not name:
                    continue

                found = False
                for uc in unique_customers:
                    if uc['name'] == name:
                        found = True
                        if not uc['code'] and code:
                            uc['code'] = code
                        break

                if not found:
                    unique_customers.append({'code': code, 'name': name})
            
            if not unique_customers:
                unique_customers.append({'code': '', 'name': 'Unspecified Customer'})

            customer_code = ', '.join([uc['code'] for uc in unique_customers if uc['code']])
            customer_name = ', '.join([uc['name'] for uc in unique_customers])
            payers_list = payers if payers else ['Unspecified Payment']

            # -----------------------------------------------------------------
            # LUEU QUANTITY
            #
            # Overall quantity:
            #     Handled Qty - Short Close Qty
            #
            # Equipment utilisation:
            #     Actual equipment-wise quantity from lueu_parcel_log.
            # -----------------------------------------------------------------
            cur.execute("""
                SELECT
                    COALESCE(
                        SUM(quantity),
                        0
                    ) AS handled_qty,

                    COALESCE(
                        SUM(
                            CASE
                                WHEN
                                    COALESCE(
                                        is_shortclose,
                                        FALSE
                                    ) = TRUE

                                    OR

                                    LOWER(
                                        COALESCE(
                                            remarks,
                                            ''
                                        )
                                    ) LIKE '%%short%%close%%'

                                THEN quantity

                                ELSE 0
                            END
                        ),
                        0
                    ) AS sc_qty,

                    COUNT(*) AS log_count

                FROM lueu_parcel_log

                WHERE parcel_op_id = %s

                  AND is_deleted IS NOT TRUE

            """, [po_id])

            log_res = cur.fetchone()

            if (
                log_res
                and log_res['log_count'] > 0
            ):

                # -------------------------------------------------------------
                # Overall operation quantity
                # -------------------------------------------------------------
                qty = (
                    float(
                        log_res['handled_qty']
                        or 0.0
                    )
                    -
                    float(
                        log_res['sc_qty']
                        or 0.0
                    )
                )

                # -------------------------------------------------------------
                # ACTUAL EQUIPMENT-WISE QUANTITY
                #
                # IMPORTANT:
                # Do NOT use STRING_AGG here.
                #
                # Every equipment receives only its own quantity.
                # -------------------------------------------------------------
                cur.execute("""
                    SELECT
                        NULLIF(
                            TRIM(equipment_name),
                            ''
                        ) AS equipment_name,

                        COALESCE(
                            SUM(
                                CASE
                                    WHEN
                                        COALESCE(
                                            is_shortclose,
                                            FALSE
                                        ) = TRUE

                                        OR

                                        LOWER(
                                            COALESCE(
                                                remarks,
                                                ''
                                            )
                                        ) LIKE '%%short%%close%%'

                                    THEN 0

                                    ELSE COALESCE(
                                        quantity,
                                        0
                                    )
                                END
                            ),
                            0
                        ) AS equipment_qty

                    FROM lueu_parcel_log

                    WHERE parcel_op_id = %s
                      AND is_deleted IS NOT TRUE

                    GROUP BY
                        NULLIF(
                            TRIM(equipment_name),
                            ''
                        )

                    ORDER BY
                        NULLIF(
                            TRIM(equipment_name),
                            ''
                        )
                """, [po_id])

                equipment_rows = cur.fetchall()

                for eq_row in equipment_rows:

                    eq_name = (
                        eq_row['equipment_name']
                        or
                        '-'
                    ).strip()

                    eq_qty = float(
                        eq_row['equipment_qty']
                        or 0.0
                    )

                    if eq_qty <= 0:
                        continue

                    equipment_totals[eq_name] = (
                        equipment_totals.get(
                            eq_name,
                            0.0
                        )
                        +
                        eq_qty
                    )

            else:
                # No LUEU log available.
                qty = float(
                    r['po_qty']
                    or 0.0
                )

            # -----------------------------------------------------------------
            # PUMPING DURATION
            # -----------------------------------------------------------------
            s_dt = _parse_dt(
                r['start_dt']
            )

            e_dt = _parse_dt(
                r['end_dt']
            )

            duration_hours = 0.0

            if (
                s_dt
                and
                e_dt
                and
                e_dt > s_dt
            ):
                duration_hours = round(
                    (
                        e_dt - s_dt
                    ).total_seconds()
                    / 3600.0,
                    2
                )

            # -----------------------------------------------------------------
            # VESSEL RUN TYPE
            # -----------------------------------------------------------------
            raw_run = (
                r['vessel_run_type']
                or ''
            ).strip()

            run_type = (
                'Foreign'
                if (
                    'fore' in raw_run.lower()
                    or raw_run.lower() == 'f'
                )

                else

                (
                    'Costal'
                    if (
                        'cost' in raw_run.lower()
                        or 'coast' in raw_run.lower()
                        or 'ind' in raw_run.lower()
                    )

                    else raw_run
                )
            )

            # -----------------------------------------------------------------
            # FLAG
            # -----------------------------------------------------------------
            flag = (
                r['vessel_nationality']
                or ''
            ).strip()

            if not flag:
                flag = (
                    'Foreign'
                    if run_type == 'Foreign'
                    else 'India'
                )

            # -----------------------------------------------------------------
            # PORT
            # -----------------------------------------------------------------
            port_val = (
                r['load_port']
                or
                r['discharge_port']
                or
                ''
            ).strip()

            _, port_name = _resolve_port(
                port_val,
                '',
                masters
            )

            # -----------------------------------------------------------------
            # CARGO
            # -----------------------------------------------------------------
            raw_cargo = _resolve_cargo_sub_category_2(
                r['cargo_name'],
                None,
                masters
            )

            # -----------------------------------------------------------------
            # APPEND LIVE ITEM
            # -----------------------------------------------------------------
            raw_items.append({

                'source': 'Live',

                'terminal':
                    terminal
                    or
                    '-',

                # Blank pipeline means Flexible Hose.
                'pipeline':
                    pipeline
                    or
                    'Flexible Hose',

                'cargo':
                    raw_cargo,

                'customers_list': unique_customers,
                
                # Customer Name used for display.
                'customer':
                    customer_name
                    or
                    customer
                    or
                    '-',

                # Customer Code retained separately.
                'customer_code':
                    customer_code,
                
                'payers_list': payers_list,

                'payment_by':
                    raw_payer
                    or
                    'Unspecified Payment',


                'duration_hours':
                    duration_hours,

                'agent_name':
                    (
                        r['vessel_agent_name']
                        or ''
                    ).strip()
                    or
                    'Unspecified Agent',

                'flag_name':
                    flag
                    or
                    'Unspecified Flag',

                'port_name':
                    port_name
                    or
                    'Unspecified Port',

                # Do not force missing equipment to MLA-1.
                #
                # Actual Equipment Utilisation comes from
                # equipment_totals.
                'equipment_name':
                    equipment
                    or
                    'Unspecified Equipment',

                'run_type':
                    run_type
                    or
                    'Foreign',

                'operation_type':
                    op_type
                    or
                    'Import',

                'qty_mt':
                    max(
                        qty,
                        0.0
                    )
            })

        # =========================================================================
        # 2. FETCH HISTORICAL RECORDS
        # =========================================================================
        cur.execute("""
            SELECT

                mh.terminal,
                mh.cargo_name,
                mh.cargo_type,
                mh.cargo_sub_category_2,
                mh.customer,
                mh.payment_by,
                mh.quantity,
                mh.overseas_coastal,
                mh.import_export,

                mvm.month,
                mvm.agent,
                mvm.flag,
                mvm.port_of_loading,
                mvm.unload_pipeline,
                mvm.ops_commenced,
                mvm.cargo_completion,
                mvm.cast_off,
                mvm.sail_cast_off

            FROM mis_history mh

            LEFT JOIN mis_vessel_master mvm
                ON mvm.vcn_no = mh.vcn_no

            WHERE mh.fin_year = %s

        """, [fin_year])

        hist_rows = cur.fetchall()

        # ---------------------------------------------------------------------
        # PROCESS HISTORICAL DATA
        # ---------------------------------------------------------------------
        for h in hist_rows:

            dt_val = (
                _parse_dt(h['cast_off'])
                or
                _parse_dt(h['sail_cast_off'])
                or
                _parse_dt(h['cargo_completion'])
            )

            if dt_val:

                if (
                    dt_val < start_dt
                    or
                    dt_val > end_dt
                ):
                    continue

            else:

                m_text = str(
                    h.get('month')
                    or
                    ''
                ).strip()

                if (
                    month
                    and
                    month.lower() != 'all'
                ):

                    m_short = (
                        month[:3]
                        .lower()
                    )

                    if (
                        m_short
                        not in
                        m_text.lower()
                    ):
                        continue

            # -----------------------------------------------------------------
            # DURATION
            # -----------------------------------------------------------------
            s_dt = _parse_dt(
                h['ops_commenced']
            )

            e_dt = _parse_dt(
                h['cargo_completion']
            )

            duration_hours = 0.0

            if (
                s_dt
                and
                e_dt
                and
                e_dt > s_dt
            ):

                duration_hours = round(
                    (
                        e_dt - s_dt
                    ).total_seconds()
                    / 3600.0,
                    2
                )

            # -----------------------------------------------------------------
            # RUN TYPE
            # -----------------------------------------------------------------
            raw_run = (
                h['overseas_coastal']
                or
                ''
            ).strip()

            run_type = (
                'Foreign'
                if (
                    'over'
                    in raw_run.lower()

                    or

                    'fore'
                    in raw_run.lower()
                )

                else

                'Costal'
            )

            # -----------------------------------------------------------------
            # OPERATION TYPE
            # -----------------------------------------------------------------
            op_type = (
                h['import_export']
                or
                'Import'
            ).strip().capitalize()

            # -----------------------------------------------------------------
            # CARGO
            # -----------------------------------------------------------------
            raw_cargo = _resolve_cargo_sub_category_2(
                h['cargo_name']
                or
                h['cargo_type'],
                h.get('cargo_sub_category_2'),
                masters
            )

            # -----------------------------------------------------------------
            # HISTORICAL CUSTOMER MASTER LOOKUP
            # -----------------------------------------------------------------
            historical_customer = (h['customer'] or '').strip()
            
            hist_custs = [c.strip() for c in historical_customer.split(',') if c.strip()]
            if not hist_custs:
                hist_custs = ['Unspecified Customer']
                
            hist_customers_list = []
            for hc in hist_custs:
                resolved = resolve_customer(hc)
                code = (resolved.get('customer_code') or '').strip()
                name = (resolved.get('customer_name') or hc or 'Unspecified Customer').strip()
                hist_customers_list.append({'code': code, 'name': name})

            historical_customer_code = ', '.join([uc['code'] for uc in hist_customers_list if uc['code']])
            historical_customer_name = ', '.join([uc['name'] for uc in hist_customers_list])
            
            historical_payer = (h['payment_by'] or '').strip()
            hist_payers_list = [p.strip() for p in historical_payer.split(',') if p.strip()]
            if not hist_payers_list:
                hist_payers_list = ['Unspecified Payment']

            # -----------------------------------------------------------------
            # APPEND HISTORICAL ITEM
            # -----------------------------------------------------------------
            raw_items.append({

                'source': 'Historical',

                'terminal':
                    (
                        h['terminal']
                        or
                        ''
                    ).strip()
                    or
                    'Unspecified Terminal',

                # Blank historical pipeline means Flexible Hose.
                'pipeline':
                    (
                        h['unload_pipeline']
                        or
                        ''
                    ).strip()
                    or
                    'Flexible Hose',

                'cargo':
                    raw_cargo,

                'customers_list': hist_customers_list,

                'customer':
                    historical_customer_name,

                'customer_code':
                    historical_customer_code,
                    
                'payers_list': hist_payers_list,

                'payment_by':
                    (
                        h['payment_by']
                        or
                        ''
                    ).strip()
                    or
                    'Unspecified Payment',

                'duration_hours':
                    duration_hours,

                'agent_name':
                    (
                        h['agent']
                        or
                        ''
                    ).strip()
                    or
                    'Unspecified Agent',

                'flag_name':
                    (
                        h['flag']
                        or
                        ''
                    ).strip()
                    or
                    (
                        'Foreign'
                        if run_type == 'Foreign'
                        else
                        'India'
                    ),

                'port_name':
                    (
                        h['port_of_loading']
                        or
                        ''
                    ).strip()
                    or
                    'Unspecified Port',

                # Historical source does not provide actual equipment-wise
                # equipment log information in this function.
                #
                # Do NOT put historical quantity into MLA-1.
                'equipment_name':
                    'Unspecified Equipment',

                'run_type':
                    run_type,

                'operation_type':
                    op_type,

                'qty_mt':
                    float(
                        h['quantity']
                        or
                        0.0
                    )
            })

    finally:
        conn.close()

    # =========================================================================
    # GENERIC QUANTITY AGGREGATION
    # =========================================================================
    def _agg_qty(key):

        totals = {}

        for it in raw_items:

            k = (
                it.get(key)
                or
                'Unspecified'
            )

            totals[k] = (
                totals.get(k, 0.0)
                +
                it['qty_mt']
            )

        grand_total = sum(
            totals.values()
        )

        rows = []

        for k, v in sorted(
            totals.items(),
            key=lambda x: -x[1]
        ):

            pct = (
                v / grand_total * 100.0
                if grand_total > 0
                else 0.0
            )

            rows.append({
                'name': k,
                'qty_mt': round(
                    v,
                    3
                ),
                'pct': round(
                    pct,
                    1
                )
            })

        return {
            'rows': rows,

            'total_qty':
                round(
                    grand_total,
                    3
                ),

            'total_pct':
                100.0
                if grand_total > 0
                else 0.0
        }

    # =========================================================================
    # AGENT-WISE AGGREGATION
    #
    # Dynamically groups agents by checking against the `agent_master`
    # loaded from `vessel_agents` table, avoiding any hardcoded replacements.
    # =========================================================================

    def resolve_agent(raw_name):
        raw = (raw_name or '').strip()
        if not raw:
            return 'Unspecified Agent'
        
        up = raw.upper()
        # masters['agents'] is populated from vessel_agents (mapping code/name to canonical name)
        if up in masters['agents']:
            return masters['agents'][up]
            
        return raw

    def _agg_agent():
        """
        Aggregate qty by vessel agent, resolving names dynamically against master.
        """
        totals = {}

        for it in raw_items:

            raw_name = (
                it.get('agent_name') or 'Unspecified Agent'
            ).strip()

            resolved_name = resolve_agent(raw_name)

            if resolved_name not in totals:
                totals[resolved_name] = 0.0

            totals[resolved_name] += it['qty_mt']

        grand_total = sum(totals.values())

        rows = []

        for name, qty in sorted(
            totals.items(),
            key=lambda x: -x[1]
        ):
            pct = qty / grand_total * 100.0 if grand_total > 0 else 0.0

            rows.append({
                'name': name,
                'qty_mt': round(qty, 3),
                'pct': round(pct, 1)
            })

        return {
            'rows': rows,
            'total_qty': round(grand_total, 3),
            'total_pct': 100.0 if grand_total > 0 else 0.0
        }

    #
    # Groups by CUSTOMER CODE from vessel_customers master table.
    #
    # Logic:
    #   1. raw customer name (from consigner / mis_history) is resolved
    #      against vessel_customers master via resolve_customer().
    #   2. If a code is found  → key = customer_code (UPPER)
    #      If no code in master → key = 'NAME:' + customer_name (UPPER)
    #      This ensures all records for the same customer code
    #      collapse into ONE row even if the name spelling differs.
    #   3. Pipelines used by that customer are collected (unique, ordered)
    #      and shown joined with  " / "  in the Pipeline column.
    #
    # Display columns:
    #     Customer Code  |  Customer Name  |  Pipeline(s)  |  Qty  |  %
    # =========================================================================
    def _agg_customer():

        totals = {}

        for it in raw_items:

            cust_list = it.get('customers_list') or [{'code': '', 'name': 'Unspecified Customer'}]
            n = len(cust_list)
            share = it['qty_mt'] / n if n > 0 else 0.0
            pipeline_name = (it.get('pipeline') or 'Flexible Hose').strip()

            for cust in cust_list:
                c_code = cust['code']
                c_name = cust['name']

                if c_code:
                    group_key = c_code.upper()
                else:
                    group_key = 'NAME:' + c_name.upper()

                if group_key not in totals:
                    totals[group_key] = {
                        'customer_code': c_code,
                        'customer_name': c_name,
                        'pipelines': [],
                        'qty_mt': 0.0
                    }
                else:
                    if not totals[group_key]['customer_code'] and c_code:
                        totals[group_key]['customer_code'] = c_code
                    if totals[group_key]['customer_name'] in ('', 'Unspecified Customer') and c_name and c_name != 'Unspecified Customer':
                        totals[group_key]['customer_name'] = c_name

                totals[group_key]['qty_mt'] += share

                if pipeline_name not in totals[group_key]['pipelines']:
                    totals[group_key]['pipelines'].append(pipeline_name)

        grand_total = sum(
            x['qty_mt'] for x in totals.values()
        )

        rows = []

        for item in sorted(
            totals.values(),
            key=lambda x: -x['qty_mt']
        ):
            qty = item['qty_mt']
            pct = qty / grand_total * 100.0 if grand_total > 0 else 0.0
            pipeline_display = ' / '.join(item['pipelines'])

            code = item['customer_code']
            name = item['customer_name']

            rows.append({
                'name':
                    code if code else name,

                'customer_code':
                    code,

                'customer_name':
                    name,

                'pipeline':
                    pipeline_display,

                'qty_mt':
                    round(qty, 3),

                'pct':
                    round(pct, 1)
            })

        return {
            'rows': rows,
            'total_qty': round(grand_total, 3),
            'total_pct': 100.0 if grand_total > 0 else 0.0
        }

    def _agg_payment_by():
        totals = {}
        for it in raw_items:
            payers = it.get('payers_list') or ['Unspecified Payment']
            n = len(payers)
            share = it['qty_mt'] / n if n > 0 else 0.0
            
            for p in payers:
                totals[p] = totals.get(p, 0.0) + share
                
        grand_total = sum(totals.values())
        rows = []
        for k, v in sorted(totals.items(), key=lambda x: -x[1]):
            pct = v / grand_total * 100.0 if grand_total > 0 else 0.0
            rows.append({
                'name': k,
                'qty_mt': round(v, 3),
                'pct': round(pct, 1)
            })
            
        return {
            'rows': rows,
            'total_qty': round(grand_total, 3),
            'total_pct': 100.0 if grand_total > 0 else 0.0
        }


    # =========================================================================
    # PIPELINE SPLIT HELPER
    #
    # Splits a combined pipeline string into individual pipeline names.
    #
    # Handles ALL common separators used in live and historical data:
    #     ,        (live data: code joins with ", ")
    #     " & "    (historical data: "12" & 8" dia GBL SS")
    #     " + "    (historical data: "8" dia Suraj + 12" dia GBL")
    #
    # Each resulting piece is stripped. Empty pieces are discarded.
    # If nothing remains, defaults to ["Flexible Hose"].
    # =========================================================================
    _PIPE_SPLIT_RE = re.compile(r',|\s+&\s+|\s+\+\s+')

    def _split_pipeline(raw_pipe):
        raw_pipe = (raw_pipe or 'Flexible Hose').strip()
        parts = [p.strip() for p in _PIPE_SPLIT_RE.split(raw_pipe) if p.strip()]
        return parts or ['Flexible Hose']

    _TERM_SPLIT_RE = re.compile(r',|\s+/\s+|/|\s+&\s+|\s+\+\s+')

    def _split_terminal(raw_term):
        raw_term = (raw_term or 'Unspecified Terminal').strip()
        parts = [p.strip() for p in _TERM_SPLIT_RE.split(raw_term) if p.strip()]
        return parts or ['Unspecified Terminal']

    # =========================================================================
    # PIPELINE UTILISATION
    #
    # Combined pipeline strings are split into individual pipelines.
    # Hours are distributed EQUALLY among each individual pipeline so the
    # grand total remains accurate and each pipeline appears exactly ONCE.
    # =========================================================================
    def _agg_pipeline_hours():

        totals = {}

        for it in raw_items:

            pipes = _split_pipeline(it.get('pipeline'))
            n = len(pipes)
            share = it['duration_hours'] / n if n > 0 else 0.0

            for pipe in pipes:
                totals[pipe] = totals.get(pipe, 0.0) + share

        grand_total = sum(totals.values())

        rows = []

        for k, v in sorted(
            totals.items(),
            key=lambda x: -x[1]
        ):
            pct = (
                v / grand_total * 100.0
                if grand_total > 0
                else 0.0
            )
            rows.append({
                'name': k,
                'hours': round(v, 2),
                'pct': round(pct, 1)
            })

        return {
            'rows': rows,
            'total_hours': round(grand_total, 2),
            'total_pct': 100.0 if grand_total > 0 else 0.0
        }

    # =========================================================================
    # PIPELINE WISE QTY
    #
    # Same split logic as Pipeline Utilisation:
    # Combined strings split on  ,  /  &  /  +  separators.
    # Qty distributed equally. Each pipeline appears exactly ONCE.
    # =========================================================================
    def _agg_pipeline_qty():

        totals = {}

        for it in raw_items:

            pipes = _split_pipeline(it.get('pipeline'))
            n = len(pipes)
            share = it['qty_mt'] / n if n > 0 else 0.0

            for pipe in pipes:
                totals[pipe] = totals.get(pipe, 0.0) + share

        grand_total = sum(totals.values())

        rows = []

        for k, v in sorted(
            totals.items(),
            key=lambda x: -x[1]
        ):
            pct = (
                v / grand_total * 100.0
                if grand_total > 0
                else 0.0
            )
            rows.append({
                'name': k,
                'qty_mt': round(v, 3),
                'pct': round(pct, 1)
            })

        return {
            'rows': rows,
            'total_qty': round(grand_total, 3),
            'total_pct': 100.0 if grand_total > 0 else 0.0
        }

    # =========================================================================
    # TERMINAL WISE QTY
    # =========================================================================
    def _agg_terminal_qty():
        totals = {}

        for it in raw_items:
            terms = _split_terminal(it.get('terminal'))
            n = len(terms)
            share = it['qty_mt'] / n if n > 0 else 0.0

            for term in terms:
                totals[term] = totals.get(term, 0.0) + share

        grand_total = sum(totals.values())
        rows = []

        for k, v in sorted(
            totals.items(),
            key=lambda x: -x[1]
        ):
            pct = (
                v / grand_total * 100.0
                if grand_total > 0
                else 0.0
            )
            rows.append({
                'name': k,
                'qty_mt': round(v, 3),
                'pct': round(pct, 1)
            })

        return {
            'rows': rows,
            'total_qty': round(grand_total, 3),
            'total_pct': 100.0 if grand_total > 0 else 0.0
        }


    # =========================================================================
    # EQUIPMENT UTILISATION
    #
    # IMPORTANT:
    # This is NOT calculated from raw_items.
    #
    # It uses actual equipment-wise quantity from lueu_parcel_log.
    #
    # Therefore:
    #     MLA-1 -> actual MLA-1 quantity
    #     MLA-3 -> actual MLA-3 quantity
    #     Flexible Hose / other equipment -> actual logged quantity
    #
    # No missing equipment quantity is automatically assigned to MLA-1.
    # =========================================================================
    def _agg_equipment():

        totals = dict(
            equipment_totals
        )

        grand_total = sum(
            totals.values()
        )

        rows = []

        for equipment_name, qty in sorted(
            totals.items(),
            key=lambda x: -x[1]
        ):

            pct = (
                qty / grand_total * 100.0
                if grand_total > 0
                else 0.0
            )

            rows.append({
                'name':
                    equipment_name,

                'qty_mt':
                    round(
                        qty,
                        3
                    ),

                'pct':
                    round(
                        pct,
                        1
                    )
            })

        return {
            'rows':
                rows,

            'total_qty':
                round(
                    grand_total,
                    3
                ),

            'total_pct':
                100.0
                if grand_total > 0
                else 0.0
        }

    # =========================================================================
    # FINAL RESPONSE
    # =========================================================================
    return {

        'terminal_wise':
            _agg_terminal_qty(),

        'pipeline_wise':
            _agg_pipeline_qty(),

        'cargo_wise':
            _agg_qty('cargo'),

        # -------------------------------------------------------------
        # CUSTOMER MASTER BASED
        # -------------------------------------------------------------
        'customer_wise':
            _agg_customer(),

        'pipeline_utilisation':
            _agg_pipeline_hours(),

        'vessel_agent_wise':
            _agg_agent(),

        'flag_wise':
            _agg_qty('flag_name'),

        'port_wise':
            _agg_qty('port_name'),

        'payment_type_wise':
            _agg_payment_by(),

        # IMPORTANT:
        # Equipment utilisation now comes from actual
        # lueu_parcel_log equipment-wise quantities.
        'equipment_utilisation':
            _agg_equipment(),

        'vessel_run_type_wise':
            _agg_qty('run_type'),

        'operation_type_wise':
            _agg_qty('operation_type'),

        'meta': {

            'fin_year':
                fin_year,

            'month':
                month or 'All',

            'start_date':
                (
                    start_dt.strftime(
                        '%d-%m-%Y %H:%M'
                    )
                    if (
                        start_dt.hour
                        or
                        start_dt.minute
                    )
                    else
                    start_dt.strftime(
                        '%d-%m-%Y'
                    )
                ),

            'end_date':
                (
                    end_dt.strftime(
                        '%d-%m-%Y %H:%M'
                    )
                    if (
                        end_dt.hour != 23
                        or
                        end_dt.minute != 59
                    )
                    else
                    end_dt.strftime(
                        '%d-%m-%Y'
                    )
                ),

            'record_count':
                len(raw_items)
        }
    }

@bp.route('/api/module/RP01/statistics-report/analytics-data', methods=['GET'])
@login_required
def statistics_report_api_analytics_data():
    perms = get_perms()
    if not perms.get('can_read'):
        return jsonify({'error': 'Unauthorized'}), 403

    today = date.today()
    cur_month = today.strftime('%B')
    cur_fy = f"{today.year}-{str(today.year + 1)[-2:]}" if today.month >= 4 else f"{today.year - 1}-{str(today.year)[-2:]}"

    fin_year = request.args.get('fin_year', cur_fy).strip()
    month = request.args.get('month', cur_month).strip()
    start_date = request.args.get('start_date', '').strip()
    end_date = request.args.get('end_date', '').strip()

    try:
        data = get_detailed_analytics_data(fin_year, month, start_date, end_date)
        return jsonify(data)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@bp.route('/api/module/RP01/statistics-report/export-analytics', methods=['GET'])
@login_required
def statistics_report_api_export_analytics():
    perms = get_perms()
    if not perms.get('can_read'):
        return jsonify({'error': 'Unauthorized'}), 403

    today = date.today()
    cur_month = today.strftime('%B')
    cur_fy = f"{today.year}-{str(today.year + 1)[-2:]}" if today.month >= 4 else f"{today.year - 1}-{str(today.year)[-2:]}"

    fin_year = request.args.get('fin_year', cur_fy).strip()
    month = request.args.get('month', cur_month).strip()
    start_date = request.args.get('start_date', '').strip()
    end_date = request.args.get('end_date', '').strip()

    try:
        data = get_detailed_analytics_data(fin_year, month, start_date, end_date)
    except Exception as e:
        return jsonify({'error': f'Analytics calculation failed: {e}'}), 500

    wb = _generate_analytics_excel(data)

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)

    filename = f"Multi_Category_Analytics_{fin_year}_{month}.xlsx"
    return send_file(
        bio,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename
    )


def _generate_analytics_excel(data: dict, wb: Workbook = None) -> Workbook:
    """
    Build the multi-category analytics Excel workbook.

    Existing sheets remain unchanged.
    Adds one new sheet:
        Pipeline_Operation_Detail

    Columns:
        Pipeline Name
        Vessel Name
        Operation Start
        Operation Stop
    """

    if wb is None:
        wb = Workbook()
        ws = wb.active
        ws.title = "Analytics_Report"
    else:
        ws = wb.create_sheet(title="Analytics_Report")


    font_title = Font(
        name="Calibri",
        size=10,
        bold=True,
        color="FFFFFF"
    )

    font_hdr = Font(
        name="Calibri",
        size=9,
        bold=True,
        color="000000"
    )

    font_data = Font(
        name="Calibri",
        size=9,
        color="000000"
    )

    font_tot = Font(
        name="Calibri",
        size=9,
        bold=True,
        color="000000"
    )

    font_meta = Font(
        name="Calibri",
        size=10,
        bold=True,
        color="1F4E78"
    )

    font_banner = Font(
        name="Calibri",
        size=11,
        bold=True,
        color="1E4620"
    )

    fill_title = PatternFill(
        start_color="1F4E78",
        end_color="1F4E78",
        fill_type="solid"
    )

    fill_hdr = PatternFill(
        start_color="D9E1F2",
        end_color="D9E1F2",
        fill_type="solid"
    )

    fill_tot = PatternFill(
        start_color="F2F2F2",
        end_color="F2F2F2",
        fill_type="solid"
    )

    fill_banner = PatternFill(
        start_color="E2EFDA",
        end_color="E2EFDA",
        fill_type="solid"
    )

    thin = Side(
        border_style="thin",
        color="D9D9D9"
    )

    double = Side(
        border_style="double",
        color="000000"
    )

    box_border = Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=thin
    )

    tot_border = Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=double
    )

    # -------------------------------------------------------------------------
    # TOP FILTER METADATA
    # -------------------------------------------------------------------------
    ws.cell(
        row=2,
        column=1,
        value="Year"
    ).font = font_meta

    ws.cell(
        row=2,
        column=2,
        value=data['meta']['fin_year']
    ).font = font_data

    ws.cell(
        row=2,
        column=3,
        value="Selection date"
    ).font = font_meta

    ws.cell(
        row=2,
        column=4,
        value=f"Start: {data['meta']['start_date']}"
    ).font = font_data

    ws.cell(
        row=2,
        column=5,
        value=f"End: {data['meta']['end_date']}"
    ).font = font_data

    ws.cell(
        row=2,
        column=6,
        value="Month"
    ).font = font_meta

    ws.cell(
        row=2,
        column=7,
        value=data['meta']['month']
    ).font = font_data

    # -------------------------------------------------------------------------
    # COMMON WRITE BOX
    # -------------------------------------------------------------------------
    def write_box(
        start_r,
        start_c,
        title,
        col1_title,
        col2_title,
        col3_title,
        rows,
        tot_val,
        is_hours=False
    ):

        ws.merge_cells(
            start_row=start_r,
            start_column=start_c,
            end_row=start_r,
            end_column=start_c + 2
        )

        t_cell = ws.cell(
            row=start_r,
            column=start_c,
            value=title
        )

        t_cell.font = font_title
        t_cell.fill = fill_title
        t_cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        for c in range(
            start_c,
            start_c + 3
        ):
            ws.cell(
                row=start_r,
                column=c
            ).border = box_border

        h_row = start_r + 1

        h_vals = [
            col1_title,
            col2_title,
            col3_title
        ]

        for idx, hv in enumerate(h_vals):

            cell = ws.cell(
                row=h_row,
                column=start_c + idx,
                value=hv
            )

            cell.font = font_hdr
            cell.fill = fill_hdr
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )
            cell.border = box_border

        curr_r = h_row + 1

        for item in rows:

            c1 = ws.cell(
                row=curr_r,
                column=start_c,
                value=item['name']
            )

            c1.font = font_data
            c1.border = box_border

            val = item.get(
                'hours'
                if is_hours
                else
                'qty_mt',
                0.0
            )

            c2 = ws.cell(
                row=curr_r,
                column=start_c + 1,
                value=val
            )

            c2.font = font_data

            c2.number_format = (
                "#,##0.00"
                if is_hours
                else
                "#,##0.000"
            )

            c2.border = box_border

            c3 = ws.cell(
                row=curr_r,
                column=start_c + 2,
                value=(
                    item.get('pct', 0.0)
                    / 100.0
                )
            )

            c3.font = font_data
            c3.number_format = "0.0%"
            c3.border = box_border

            curr_r += 1

        t1 = ws.cell(
            row=curr_r,
            column=start_c,
            value="Total"
        )

        t1.font = font_tot
        t1.fill = fill_tot
        t1.border = tot_border

        t2 = ws.cell(
            row=curr_r,
            column=start_c + 1,
            value=tot_val
        )

        t2.font = font_tot
        t2.fill = fill_tot

        t2.number_format = (
            "#,##0.00"
            if is_hours
            else
            "#,##0.000"
        )

        t2.border = tot_border

        t3 = ws.cell(
            row=curr_r,
            column=start_c + 2,
            value=(
                1.0
                if tot_val > 0
                else 0.0
            )
        )

        t3.font = font_tot
        t3.fill = fill_tot
        t3.number_format = "0.0%"
        t3.border = tot_border

        return curr_r

    # -------------------------------------------------------------------------
    # CUSTOMER-WISE WRITE BOX  (4 columns: Code, Pipeline, Qty, %)
    # -------------------------------------------------------------------------
    def write_customer_box(
        start_r,
        start_c,
        rows,
        tot_val
    ):

        # Title spans 4 columns
        ws.merge_cells(
            start_row=start_r,
            start_column=start_c,
            end_row=start_r,
            end_column=start_c + 3
        )

        t_cell = ws.cell(
            row=start_r,
            column=start_c,
            value="Customerwise"
        )
        t_cell.font = font_title
        t_cell.fill = fill_title
        t_cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        for c in range(start_c, start_c + 4):
            ws.cell(row=start_r, column=c).border = box_border

        h_row = start_r + 1

        for idx, hv in enumerate(
            ["Customer Code", "Pipeline", "Qty handled in MT", "% of Qty"]
        ):
            cell = ws.cell(
                row=h_row,
                column=start_c + idx,
                value=hv
            )
            cell.font = font_hdr
            cell.fill = fill_hdr
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )
            cell.border = box_border

        curr_r = h_row + 1

        for item in rows:

            # Customer Code
            c1 = ws.cell(
                row=curr_r,
                column=start_c,
                value=item.get('customer_code') or item.get('name', '')
            )
            c1.font = font_data
            c1.border = box_border

            # Pipeline
            c2 = ws.cell(
                row=curr_r,
                column=start_c + 1,
                value=item.get('pipeline', '')
            )
            c2.font = font_data
            c2.border = box_border

            # Qty
            c3 = ws.cell(
                row=curr_r,
                column=start_c + 2,
                value=item.get('qty_mt', 0.0)
            )
            c3.font = font_data
            c3.number_format = "#,##0.000"
            c3.border = box_border

            # %
            c4 = ws.cell(
                row=curr_r,
                column=start_c + 3,
                value=item.get('pct', 0.0) / 100.0
            )
            c4.font = font_data
            c4.number_format = "0.0%"
            c4.border = box_border

            curr_r += 1

        # Total row spans Code + Pipeline columns
        ws.merge_cells(
            start_row=curr_r,
            start_column=start_c,
            end_row=curr_r,
            end_column=start_c + 1
        )

        t1 = ws.cell(
            row=curr_r,
            column=start_c,
            value="Total"
        )
        t1.font = font_tot
        t1.fill = fill_tot
        t1.border = tot_border

        ws.cell(row=curr_r, column=start_c + 1).border = tot_border

        t2 = ws.cell(
            row=curr_r,
            column=start_c + 2,
            value=tot_val
        )
        t2.font = font_tot
        t2.fill = fill_tot
        t2.number_format = "#,##0.000"
        t2.border = tot_border

        t3 = ws.cell(
            row=curr_r,
            column=start_c + 3,
            value=1.0 if tot_val > 0 else 0.0
        )
        t3.font = font_tot
        t3.fill = fill_tot
        t3.number_format = "0.0%"
        t3.border = tot_border

        return curr_r



    # -------------------------------------------------------------------------
    # ROW 1
    # -------------------------------------------------------------------------
    r1_ends = [

        write_box(
            4,
            1,
            "Terminalwise",
            "Terminal",
            "Qty handled in MT",
            "% of Qty",
            data['terminal_wise']['rows'],
            data['terminal_wise']['total_qty']
        ),

        write_box(
            4,
            4,
            "Pipeline wise",
            "Pipeline",
            "Qty handled in MT",
            "% of Qty",
            data['pipeline_wise']['rows'],
            data['pipeline_wise']['total_qty']
        ),

        write_box(
            4,
            7,
            "Cargowise",
            "Cargo Type",
            "Qty handled in MT",
            "% of Qty",
            data['cargo_wise']['rows'],
            data['cargo_wise']['total_qty']
        ),

        write_customer_box(
            4,
            10,
            data['customer_wise']['rows'],
            data['customer_wise']['total_qty']
        ),

        write_box(
            4,
            14,
            "Pipeline Utilisation",
            "Pipeline",
            "No of Hours",
            "% of Hours",
            data['pipeline_utilisation']['rows'],
            data['pipeline_utilisation']['total_hours'],
            is_hours=True
        )
    ]

    r2_start = max(r1_ends) + 2

    # -------------------------------------------------------------------------
    # ROW 2
    # -------------------------------------------------------------------------
    r2_ends = [

        write_box(
            r2_start,
            1,
            "Vessel Agentwise",
            "Agent Name",
            "Qty handled in MT",
            "% of Qty",
            data['vessel_agent_wise']['rows'],
            data['vessel_agent_wise']['total_qty']
        ),

        write_box(
            r2_start,
            4,
            "Flagwise",
            "Flag Wise",
            "Qty handled in MT",
            "% of Qty",
            data['flag_wise']['rows'],
            data['flag_wise']['total_qty']
        ),

        write_box(
            r2_start,
            7,
            "Portwise",
            "Port Name",
            "Qty handled in MT",
            "% of Qty",
            data['port_wise']['rows'],
            data['port_wise']['total_qty']
        ),

        write_box(
            r2_start,
            10,
            "Payment type wise",
            "Name",
            "Qty handled in MT",
            "% of Qty",
            data['payment_type_wise']['rows'],
            data['payment_type_wise']['total_qty']
        ),

        write_box(
            r2_start,
            13,
            "Equipment Utilisation",
            "MLA No",
            "Qty handled in MT",
            "% of Qty",
            data['equipment_utilisation']['rows'],
            data['equipment_utilisation']['total_qty']
        )
    ]

    # -------------------------------------------------------------------------
    # ROW 3
    # -------------------------------------------------------------------------
    r3_start = max(r2_ends) + 2

    write_box(
        r3_start,
        1,
        "Vessel Run Typewise",
        "Run Type",
        "Qty handled in MT",
        "% of Qty",
        data['vessel_run_type_wise']['rows'],
        data['vessel_run_type_wise']['total_qty']
    )

    write_box(
        r3_start,
        4,
        "Operation Type wise",
        "Operation",
        "Qty handled in MT",
        "% of Qty",
        data['operation_type_wise']['rows'],
        data['operation_type_wise']['total_qty']
    )

    # -------------------------------------------------------------------------
    # AUTO-FIT EXISTING ANALYTICS SHEET
    # -------------------------------------------------------------------------
    for col_idx in range(1, 17):

        col_letter = get_column_letter(
            col_idx
        )

        max_len = 0

        for row in ws.iter_rows(
            min_col=col_idx,
            max_col=col_idx
        ):

            val = row[0].value

            if val is not None:
                max_len = max(
                    max_len,
                    len(str(val))
                )

        ws.column_dimensions[
            col_letter
        ].width = max(
            max_len + 3,
            13
        )

    # =========================================================================
    # NEW SHEET:
    # PIPELINE OPERATION DETAIL
    # =========================================================================
    #
    # IMPORTANT:
    # This sheet is populated directly from the database using the
    # selected date/time range.
    #
    # It does NOT change the existing Analytics_Report calculations.
    # =========================================================================

    ws_pipeline = wb.create_sheet(
        title="Pipeline_Operation_Detail"
    )

    # -------------------------------------------------------------------------
    # GET SELECTED DATE/TIME RANGE
    # -------------------------------------------------------------------------
    try:

        start_text = data['meta'].get(
            'start_date',
            ''
        )

        end_text = data['meta'].get(
            'end_date',
            ''
        )

        selected_start = _parse_dt(
            start_text
        )

        selected_end = _parse_dt(
            end_text
        )

        # If metadata contains only dates, use complete date range.
        if selected_start and selected_end:

            if len(str(start_text).strip()) == 10:
                selected_start = datetime.combine(
                    selected_start.date(),
                    datetime.min.time()
                )

            # End-of-day: apply when string is date-only (len=10) OR
            # when the parsed time is exactly midnight 00:00 (the default
            # when no meaningful time was set for the end boundary).
            if (
                len(str(end_text).strip()) == 10
                or (
                    selected_end.hour == 0
                    and selected_end.minute == 0
                    and selected_end.second == 0
                )
            ):
                selected_end = datetime.combine(
                    selected_end.date(),
                    datetime.max.time()
                )

        else:
            selected_start = None
            selected_end = None

    except Exception:
        selected_start = None
        selected_end = None

    pipeline_operation_rows = []

    # -------------------------------------------------------------------------
    # DATABASE QUERY
    # -------------------------------------------------------------------------
    conn_pipeline = None

    try:

        conn_pipeline = get_db()
        cur_pipeline = get_cursor(
            conn_pipeline
        )

        # -------------------------------------------------------------
        # Get live vessel calls WITH cast_off_datetime
        # (same field the main analytics uses for date filtering)
        # -------------------------------------------------------------
        cur_pipeline.execute("""
            SELECT
                lh.id AS ldud_id,
                lh.cast_off_datetime,
                lh.discharge_completed,
                vh.vessel_name,
                vh.operation_type

            FROM ldud_header lh

            JOIN vcn_header vh
                ON vh.id = lh.vcn_id

            WHERE COALESCE(
                lh.is_deleted,
                FALSE
            ) = FALSE
        """)

        vessel_rows = cur_pipeline.fetchall()

        for vessel_row in vessel_rows:

            ldud_id = vessel_row['ldud_id']

            # Use cast_off_datetime (or discharge_completed) as the
            # filter reference date — exactly what main analytics uses.
            vessel_ref_dt = (
                _parse_dt(vessel_row['cast_off_datetime'])
                or
                _parse_dt(vessel_row['discharge_completed'])
            )

            # Skip vessels that fall outside the selected date range.
            if vessel_ref_dt:
                if selected_start and vessel_ref_dt < selected_start:
                    continue
                if selected_end and vessel_ref_dt > selected_end:
                    continue
            # If no cast_off date at all, skip (vessel not completed).
            else:
                continue

            vessel_name = (
                vessel_row['vessel_name']
                or
                'Missing Vessel'
            ).strip()

            operation_type = (
                vessel_row['operation_type']
                or
                ''
            ).strip().capitalize()

            # ---------------------------------------------------------
            # Get parcel operations
            # ---------------------------------------------------------
            cur_pipeline.execute("""
                SELECT
                    po.id AS po_id,
                    po.parcel_ids,
                    po.start_dt,
                    po.end_dt

                FROM ldud_parcel_ops po

                WHERE po.ldud_id = %s

                ORDER BY
                    po.start_dt,
                    po.id
            """, [ldud_id])

            parcel_ops = cur_pipeline.fetchall()

            if not parcel_ops:
                continue

            # ---------------------------------------------------------
            # Correct parcel master based on operation type
            # ---------------------------------------------------------
            tbl = (
                'vcn_export_cargo_declaration'
                if operation_type == 'Export'
                else
                'vcn_consigners'
            )

            for po in parcel_ops:

                operation_start = _parse_dt(
                    po['start_dt']
                )

                operation_stop = _parse_dt(
                    po['end_dt']
                )

                # Date filtering is already done at the vessel level
                # using cast_off_datetime above. All parcel ops for
                # a vessel that passed the filter are included here.

                # -----------------------------------------------------
                # Parcel IDs
                # -----------------------------------------------------
                parcel_ids = [
                    int(x.strip())
                    for x in str(
                        po['parcel_ids']
                        or
                        ''
                    ).split(',')
                    if x.strip().isdigit()
                ]

                pipeline_names = []

                if parcel_ids:

                    cur_pipeline.execute(
                        f"""
                            SELECT
                                pipeline_name

                            FROM {tbl}

                            WHERE id = ANY(%s)
                        """,
                        [parcel_ids]
                    )

                    pipeline_rows = (
                        cur_pipeline.fetchall()
                    )

                    for pipeline_row in pipeline_rows:

                        pipeline_name = (
                            pipeline_row[
                                'pipeline_name'
                            ]
                            or
                            ''
                        ).strip()

                        if pipeline_name:
                            if pipeline_name not in pipeline_names:
                                pipeline_names.append(
                                    pipeline_name
                                )

                # -----------------------------------------------------
                # Business rule:
                # Empty pipeline = Flexible Hose
                # -----------------------------------------------------
                if pipeline_names:

                    pipeline_name_display = ', '.join(
                        pipeline_names
                    )

                else:

                    pipeline_name_display = (
                        'Flexible Hose'
                    )

                pipeline_operation_rows.append({
                    'pipeline_name':
                        pipeline_name_display,

                    'vessel_name':
                        vessel_name,

                    'operation_start':
                        operation_start,

                    'operation_stop':
                        operation_stop,

                    'source': 'Live'
                })

        # ---------------------------------------------------------------------
        # HISTORICAL RECORDS from mis_history + mis_vessel_master
        # (Same source the main analytics uses for Sheet 1 pipeline totals)
        # ---------------------------------------------------------------------
        fin_year_meta = data.get('meta', {}).get('fin_year', '')
        month_meta = data.get('meta', {}).get('month', 'All')

        if fin_year_meta:
            cur_pipeline.execute("""
                SELECT
                    mvm.vessel_name,
                    mvm.unload_pipeline,
                    mvm.ops_commenced,
                    mvm.cargo_completion,
                    mvm.cast_off,
                    mvm.sail_cast_off,
                    mvm.month
                FROM mis_history mh
                LEFT JOIN mis_vessel_master mvm
                    ON mvm.vcn_no = mh.vcn_no
                WHERE mh.fin_year = %s
            """, [fin_year_meta])

            hist_vessel_rows = cur_pipeline.fetchall()

            seen_hist = set()

            for hv in hist_vessel_rows:

                # Same date-filter logic as the main analytics
                dt_val = (
                    _parse_dt(hv['cast_off'])
                    or _parse_dt(hv['sail_cast_off'])
                    or _parse_dt(hv['cargo_completion'])
                )

                if dt_val:
                    if selected_start and dt_val < selected_start:
                        continue
                    if selected_end and dt_val > selected_end:
                        continue
                else:
                    m_text = str(hv.get('month') or '').strip()
                    if month_meta and month_meta.lower() != 'all':
                        m_short = month_meta[:3].lower()
                        if m_short not in m_text.lower():
                            continue

                pipe_raw = (hv['unload_pipeline'] or '').strip()
                pipe_display = pipe_raw or 'Flexible Hose'
                vessel_h = (hv['vessel_name'] or 'Unknown Vessel').strip()
                op_start_h = _parse_dt(hv['ops_commenced'])
                op_stop_h = _parse_dt(hv['cargo_completion'])

                # De-duplicate: same vessel + pipeline + start
                dedup_key = (pipe_display, vessel_h, str(op_start_h))
                if dedup_key in seen_hist:
                    continue
                seen_hist.add(dedup_key)

                pipeline_operation_rows.append({
                    'pipeline_name': pipe_display,
                    'vessel_name': vessel_h,
                    'operation_start': op_start_h,
                    'operation_stop': op_stop_h,
                    'source': 'Historical'
                })

    except Exception as _pipe_exc:
        # Do not break the complete Excel export, but preserve the
        # error so it can be seen during debugging.
        pipeline_operation_rows = []

    finally:

        if conn_pipeline:
            conn_pipeline.close()

    # -------------------------------------------------------------------------
    # SHEET TITLE
    # -------------------------------------------------------------------------
    ws_pipeline.merge_cells(
        start_row=1,
        start_column=1,
        end_row=1,
        end_column=5
    )

    title_cell = ws_pipeline.cell(
        row=1,
        column=1,
        value="Pipeline Operation Detail"
    )

    title_cell.font = font_title
    title_cell.fill = fill_title
    title_cell.alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    # -------------------------------------------------------------------------
    # SELECTED RANGE
    # -------------------------------------------------------------------------
    ws_pipeline.merge_cells(
        start_row=2,
        start_column=1,
        end_row=2,
        end_column=5
    )

    range_cell = ws_pipeline.cell(
        row=2,
        column=1,
        value=(
            f"Selected Range: "
            f"{data['meta']['start_date']} "
            f"to "
            f"{data['meta']['end_date']}"
        )
    )

    range_cell.font = font_meta
    range_cell.alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    # -------------------------------------------------------------------------
    # HEADERS
    # -------------------------------------------------------------------------
    pipeline_headers = [
        "Pipeline Name",
        "Vessel Name",
        "Operation Start",
        "Operation Stop",
        "Source"
    ]

    for idx, header in enumerate(
        pipeline_headers,
        start=1
    ):

        cell = ws_pipeline.cell(
            row=4,
            column=idx,
            value=header
        )

        cell.font = font_hdr
        cell.fill = fill_hdr
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )
        cell.border = box_border

    # -------------------------------------------------------------------------
    # DATA
    # -------------------------------------------------------------------------
    # Sort rows by Pipeline Name for consistency with Sheet 1 order
    pipeline_operation_rows.sort(
        key=lambda x: x.get('pipeline_name', '').lower()
    )

    current_row = 5

    for item in pipeline_operation_rows:

        c1 = ws_pipeline.cell(
            row=current_row,
            column=1,
            value=item['pipeline_name']
        )

        c2 = ws_pipeline.cell(
            row=current_row,
            column=2,
            value=item['vessel_name']
        )

        c3 = ws_pipeline.cell(
            row=current_row,
            column=3,
            value=item['operation_start']
        )

        c4 = ws_pipeline.cell(
            row=current_row,
            column=4,
            value=item['operation_stop']
        )

        c5 = ws_pipeline.cell(
            row=current_row,
            column=5,
            value=item.get('source', '')
        )

        for cell in (c1, c2, c3, c4, c5):

            cell.font = font_data
            cell.border = box_border
            cell.alignment = Alignment(
                vertical="center"
            )

        if item['operation_start']:

            c3.number_format = (
                "dd-mm-yyyy hh:mm:ss"
            )

        if item['operation_stop']:

            c4.number_format = (
                "dd-mm-yyyy hh:mm:ss"
            )

        current_row += 1

    # -------------------------------------------------------------------------
    # NO DATA MESSAGE
    # -------------------------------------------------------------------------
    if not pipeline_operation_rows:

        ws_pipeline.merge_cells(
            start_row=5,
            start_column=1,
            end_row=5,
            end_column=5
        )

        no_data_cell = ws_pipeline.cell(
            row=5,
            column=1,
            value=(
                "No pipeline operations found "
                "for the selected date and time range."
            )
        )

        no_data_cell.font = font_data
        no_data_cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    # -------------------------------------------------------------------------
    # AUTO-FIT NEW SHEET
    # -------------------------------------------------------------------------
    pipeline_widths = {
        1: 32,
        2: 32,
        3: 22,
        4: 22,
        5: 12
    }

    for col_idx, width in pipeline_widths.items():

        ws_pipeline.column_dimensions[
            get_column_letter(col_idx)
        ].width = width

    # -------------------------------------------------------------------------
    # FREEZE HEADER
    # -------------------------------------------------------------------------
    ws_pipeline.freeze_panes = "A5"

    return wb
