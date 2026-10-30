import io
import math
import re
from datetime import datetime, date, timedelta, time
from decimal import Decimal
from functools import wraps

from flask import (
    Blueprint,
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

MONTH_NAMES = ["Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"]

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


def _format_dt(dt):
    if not dt:
        return ''
    if isinstance(dt, datetime):
        return dt.strftime('%Y-%m-%d %H:%M')
    parsed = _parse_dt(dt)
    return parsed.strftime('%Y-%m-%d %H:%M') if parsed else str(dt)


def _format_date_only(dt):
    if not dt:
        return ''
    if isinstance(dt, datetime):
        return dt.strftime('%Y-%m-%d')
    parsed = _parse_dt(dt)
    return parsed.strftime('%Y-%m-%d') if parsed else str(dt)[:10]


def _dt_diff_days(end_dt, start_dt):
    if not end_dt or not start_dt:
        return None
    d_end = _parse_dt(end_dt)
    d_start = _parse_dt(start_dt)
    if not d_end or not d_start or d_end < d_start:
        return None
    return (d_end - d_start).total_seconds() / 86400.0


def _dt_diff_hrs(end_dt, start_dt):
    days = _dt_diff_days(end_dt, start_dt)
    return (days * 24.0) if days is not None else None


def _safe_float(val, default=None):
    if val is None:
        return default
    try:
        f = float(val)
        return f if not math.isnan(f) else default
    except (ValueError, TypeError):
        return default


def _date_to_fin_year_and_month(dt):
    if not dt:
        return '', ''
    d = _parse_dt(dt)
    if not d:
        return '', ''
    if d.month >= 4:
        fy = f"{d.year}-{(d.year + 1) % 100:02d}"
    else:
        fy = f"{d.year - 1}-{d.year % 100:02d}"
    m_str = d.strftime('%b-%y')
    return fy, m_str


def _load_port_master_mapping(cur):
    """
    Dynamically loads port mappings from the database:
    1. Historical distinct port mappings from mis_vessel_master
    2. Authoritative live Port Master table (VPM01 - port_master), taking precedence.
    No hardcoded port dictionary — completely dynamic from database master data.
    Returns (port_code_map, port_name_map, known_codes).
    """
    port_code_map = {}
    port_name_map = {}
    known_codes = set()

    # 1. Historical distinct ports from mis_vessel_master
    try:
        cur.execute("""
            SELECT DISTINCT port_of_loading, port_code 
            FROM mis_vessel_master 
            WHERE port_code IS NOT NULL AND TRIM(port_code) != ''
        """)
        for m_p in cur.fetchall():
            p_name = (m_p.get('port_of_loading') or '').strip()
            p_code = (m_p.get('port_code') or '').strip().upper()
            if p_code:
                known_codes.add(p_code)
                port_code_map[p_code] = p_code
                if p_name and p_code not in port_name_map:
                    port_name_map[p_code] = p_name
            if p_name and p_code:
                norm_name = p_name.upper()
                port_code_map[norm_name] = p_code
                first_token = re.split(r'[,/(-]', norm_name)[0].strip()
                if first_token and first_token not in port_code_map:
                    port_code_map[first_token] = p_code
    except Exception:
        pass

    # 2. Authoritative live Port Master table (VPM01 - port_master)
    try:
        cur.execute("""
            SELECT name, port_code 
            FROM port_master 
            WHERE (port_code IS NOT NULL AND TRIM(port_code) != '')
               OR (name IS NOT NULL AND TRIM(name) != '')
        """)
        for pm in cur.fetchall():
            p_name = (pm.get('name') or '').strip()
            p_code = (pm.get('port_code') or '').strip().upper()
            if p_code:
                known_codes.add(p_code)
                port_code_map[p_code] = p_code
                if p_name:
                    port_name_map[p_code] = p_name
            if p_name and p_code:
                norm_name = p_name.upper()
                port_code_map[norm_name] = p_code
                first_token = re.split(r'[,/(-]', norm_name)[0].strip()
                port_code_map[first_token] = p_code
    except Exception:
        pass

    return port_code_map, port_name_map, known_codes


def _resolve_port_info(port_loading, existing_port_code, port_code_map, port_name_map, known_codes):
    """
    Resolves official port_code and port_of_loading from port_master.
    """
    p_code = (existing_port_code or '').strip().upper()
    p_load = (port_loading or '').strip()

    # If code is already recognized and valid
    if p_code and p_code in known_codes:
        if not p_load and p_code in port_name_map:
            p_load = port_name_map[p_code]
        return p_code, p_load

    # If port_of_loading is provided
    if p_load:
        norm_load = p_load.upper()
        # 1. Exact match
        if norm_load in port_code_map:
            return port_code_map[norm_load], p_load

        # 2. Match primary part before comma / separator
        first_token = re.split(r'[,/(-]', norm_load)[0].strip()
        if first_token and first_token in port_code_map:
            return port_code_map[first_token], p_load

        # 3. Substring match (min length 4)
        for k, code in port_code_map.items():
            if len(k) >= 4 and (k in norm_load or norm_load in k):
                return code, p_load

    return p_code, p_load


def _load_vessel_cargo_mapping(cur):
    """
    Dynamically loads cargo categorization from vessel_cargo (VCG01 - Vessel Cargo Master)
    and backfills from mis_vessel_master.
    No hardcoded mappings — 100% dynamic based on database records.
    Returns cargo_map: norm_name -> {new_cat, category1, category}.
    """
    cargo_map = {}

    def _clean(val, default=''):
        if not val:
            return default
        s = str(val).strip()
        if s.upper() == 'POL':
            return 'POL'
        if s.upper() in ('POL-BLACK', 'POL BLACK'):
            return 'POL Black'
        if s.upper() in ('FARM LIQUIDS', 'FARM LIQUID'):
            return 'Other Liquid'
        return s.title()

    # 1. Backfill from mis_vessel_master (historical ground truth)
    try:
        cur.execute("""
            SELECT DISTINCT cargo, new_cat, category1, category
            FROM mis_vessel_master
            WHERE cargo IS NOT NULL AND TRIM(cargo) != ''
        """)
        for m in cur.fetchall():
            mc = (m.get('cargo') or '').strip().upper()
            if mc:
                cargo_map[mc] = {
                    'new_cat': m.get('new_cat') or '',
                    'category1': m.get('category1') or '',
                    'category': m.get('category') or ''
                }
    except Exception:
        pass

    # 2. Authoritative live Vessel Cargo Master (vessel_cargo / VCG01) — overrides and takes priority
    try:
        cur.execute("""
            SELECT 
                cargo_name, cargo_code, cargo_type, 
                cargo_category, cargo_category_2, cargo_sub_category, cargo_sub_category_2
            FROM vessel_cargo
            WHERE (cargo_name IS NOT NULL AND TRIM(cargo_name) != '')
               OR (cargo_code IS NOT NULL AND TRIM(cargo_code) != '')
        """)
        for r in cur.fetchall():
            cn = (r.get('cargo_name') or '').strip().upper()
            cc = (r.get('cargo_code') or '').strip().upper()

            sub2 = (r.get('cargo_sub_category_2') or '').strip()
            cat2 = (r.get('cargo_category_2') or '').strip()
            sub1 = (r.get('cargo_sub_category') or '').strip()
            cat = (r.get('cargo_category') or '').strip()
            ctype = (r.get('cargo_type') or '').strip()

            new_cat_val = ''
            for test in (sub2, cat2, cat, ctype):
                tu = test.upper()
                if 'CHEM' in tu:
                    new_cat_val = 'Chemical'
                    break
                elif 'EDIBLE' in tu:
                    new_cat_val = 'Edible Oil'
                    break
                elif 'POL' in tu:
                    new_cat_val = 'POL'
                    break
                elif 'FARM' in tu or 'OTHER' in tu:
                    new_cat_val = 'Other Liquid'
                    break
            if not new_cat_val and sub2:
                new_cat_val = _clean(sub2)
            elif not new_cat_val and cat:
                new_cat_val = _clean(cat)

            cat1_val = _clean(cat2 or sub1 or sub2 or 'Other')
            cat_val = _clean(cat or sub2 or 'Other Liquid')

            existing = cargo_map.get(cn) or cargo_map.get(cc) or {}
            final_info = {
                'new_cat': new_cat_val or existing.get('new_cat') or 'Other Liquid',
                'category1': cat1_val if (cat2 or sub1) else (existing.get('category1') or cat1_val),
                'category': cat_val if cat else (existing.get('category') or cat_val)
            }

            if cn:
                cargo_map[cn] = final_info
            if cc:
                cargo_map[cc] = final_info
    except Exception:
        pass

    return cargo_map


def _resolve_cargo_categories(cargo_str, cargo_map):
    """
    Resolves (new_cat, category1, category) from vessel_cargo master.
    Handles multiple cargos if separated by comma or slash.
    """
    if not cargo_str or not str(cargo_str).strip():
        return '', '', ''

    raw_str = str(cargo_str).strip()
    parts = [p.strip() for p in re.split(r'[,/]', raw_str) if p.strip()]
    if not parts:
        parts = [raw_str]

    res_new_cat = []
    res_cat1 = []
    res_cat = []

    for p in parts:
        pu = p.upper()
        match = cargo_map.get(pu)
        if not match:
            cleaned = re.sub(r'\[.*?\]|\(.*?\)', '', pu).strip()
            match = cargo_map.get(cleaned)
        if not match:
            for k, v in cargo_map.items():
                if len(k) >= 4 and (k in pu or pu in k):
                    match = v
                    break

        if match:
            if match.get('new_cat') and match['new_cat'] not in res_new_cat:
                res_new_cat.append(match['new_cat'])
            if match.get('category1') and match['category1'] not in res_cat1:
                res_cat1.append(match['category1'])
            if match.get('category') and match['category'] not in res_cat:
                res_cat.append(match['category'])

    final_new_cat = ', '.join(res_new_cat) if res_new_cat else 'Other Liquid'
    final_cat1 = ', '.join(res_cat1) if res_cat1 else 'Other'
    final_cat = ', '.join(res_cat) if res_cat else 'Other Liquid'

    return final_new_cat, final_cat1, final_cat


# ──────────────────────────────────────────────────────────────────
#  Data Fetching: Historical Master + Live Operations
# ──────────────────────────────────────────────────────────────────

def fetch_master_report_records():
    """
    Returns unified list of vessel call records across historical master and live operational data.
    Ensures zero duplicate joins and authoritative metrics.
    """
    conn = get_db()
    records = []
    seen_vcns = set()

    try:
        cur = get_cursor(conn)
        port_code_map, port_name_map, known_codes = _load_port_master_mapping(cur)
        cargo_map = _load_vessel_cargo_mapping(cur)

        # -------------------------------------------------------------
        # 1. Historical Dataset: mis_vessel_master (+ parcel details from mis_history)
        # -------------------------------------------------------------
        cur.execute("""
            SELECT 
                m.id,
                m.sr_no,
                m.fin_year,
                m.month,
                m.berth_no,
                m.vcn_no,
                m.vessel_name,
                m.overseas_coastal,
                m.foreign_indian,
                m.imo_no,
                m.flag,
                m.bhc,
                m.port_code,
                m.port_of_loading,
                m.grt,
                m.draft,
                m.loa,
                m.import_export,
                m.agent,
                m.unload_pipeline,
                m.consigner,
                m.unloading_terminal,
                m.new_cat,
                m.category1,
                m.category,
                m.cargo,
                m.nor,
                m.anchorage_time,
                m.pilot_pickup,
                m.first_line,
                m.alongside,
                m.ops_commenced,
                m.cargo_completion,
                m.sail_cast_off,
                m.cast_off,
                m.pilot_board_departure,
                m.pilot_disembarked,
                m.quantity,
                m.flow_rate,
                m.remarks,
                m.pre_berthing_waiting,
                m.waiting_port,
                m.waiting_non_port,
                m.stay_at_berth,
                m.arrive_to_comm,
                m.working_time,
                m.non_working_total,
                m.non_working_port,
                m.non_working_non_port,
                m.inward_movement,
                m.outward_movement,
                h.bl_quantity,
                h.importer_name
            FROM mis_vessel_master m
            LEFT JOIN (
                SELECT 
                    vcn_no,
                    SUM(quantity) AS bl_quantity,
                    string_agg(DISTINCT customer, ', ') AS importer_name
                FROM mis_history
                GROUP BY vcn_no
            ) h ON h.vcn_no = m.vcn_no
            ORDER BY m.sr_no NULLS LAST, m.id ASC
        """)
        m_rows = cur.fetchall()

        for r in m_rows:
            vcn = (r.get('vcn_no') or '').strip().upper()
            if vcn:
                seen_vcns.add(vcn)

            # Timings
            nor = _parse_dt(r.get('nor'))
            anchorage = _parse_dt(r.get('anchorage_time')) or nor
            pilot_pick = _parse_dt(r.get('pilot_pickup'))
            first_line = _parse_dt(r.get('first_line'))
            alongside = _parse_dt(r.get('alongside'))
            ops_comm = _parse_dt(r.get('ops_commenced'))
            cargo_comp = _parse_dt(r.get('cargo_completion'))
            sail_cast = _parse_dt(r.get('sail_cast_off')) or _parse_dt(r.get('cast_off'))
            cast_off = _parse_dt(r.get('cast_off')) or sail_cast
            pilot_dept = _parse_dt(r.get('pilot_board_departure'))
            pilot_dis = _parse_dt(r.get('pilot_disembarked'))

            # Status determination
            status = 'Unberth' if sail_cast else ('Berth' if alongside else 'Waiting')

            # Reference date for filtering
            ref_dt = sail_cast or alongside or ops_comm or nor or anchorage

            # Quantities
            qty = _safe_float(r.get('quantity'), 0.0)
            bl_qty = _safe_float(r.get('bl_quantity'), qty)
            discharged_till_now = qty if status == 'Unberth' else qty # historical calls are completed
            balance_qty = max(0.0, qty - discharged_till_now)

            # Flow rates
            working_time_days = _safe_float(r.get('working_time'))
            if working_time_days is None and cargo_comp and ops_comm and cargo_comp > ops_comm:
                working_time_days = (cargo_comp - ops_comm).total_seconds() / 86400.0

            working_time_hrs = (working_time_days * 24.0) if working_time_days is not None else None

            flow_rate = _safe_float(r.get('flow_rate'))
            if (flow_rate is None or flow_rate <= 0) and working_time_hrs and working_time_hrs > 0:
                flow_rate = round(qty / working_time_hrs, 2)

            present_flow_rate = flow_rate

            # Balance time & Expected completion
            exp_flow_rate = flow_rate
            balance_time_hrs = (balance_qty / exp_flow_rate) if (exp_flow_rate and exp_flow_rate > 0) else None
            expected_completion = None

            # Operational durations
            pbw_days = _safe_float(r.get('pre_berthing_waiting'))
            if pbw_days is None and alongside and nor and alongside > nor:
                pbw_days = (alongside - nor).total_seconds() / 86400.0
            pbw_hrs = (pbw_days * 24.0) if pbw_days is not None else None

            waiting_port_days = _safe_float(r.get('waiting_port'), pbw_days or 0.0)
            waiting_port_hrs = (waiting_port_days * 24.0) if waiting_port_days is not None else None

            waiting_np_days = _safe_float(r.get('waiting_non_port'), 0.0)
            waiting_np_hrs = (waiting_np_days * 24.0) if waiting_np_days is not None else None

            stay_berth_days = _safe_float(r.get('stay_at_berth'))
            if stay_berth_days is None and sail_cast and alongside and sail_cast > alongside:
                stay_berth_days = (sail_cast - alongside).total_seconds() / 86400.0
            stay_berth_hrs = (stay_berth_days * 24.0) if stay_berth_days is not None else None

            arr_comm_days = _safe_float(r.get('arrive_to_comm'))
            if arr_comm_days is None and ops_comm and alongside and ops_comm > alongside:
                arr_comm_days = (ops_comm - alongside).total_seconds() / 86400.0
            arr_comm_hrs = (arr_comm_days * 24.0) if arr_comm_days is not None else None

            nw_total_days = _safe_float(r.get('non_working_total'))
            if nw_total_days is None and stay_berth_days is not None and working_time_days is not None:
                nw_total_days = max(0.0, stay_berth_days - working_time_days)
            nw_total_hrs = (nw_total_days * 24.0) if nw_total_days is not None else None

            nw_port_days = _safe_float(r.get('non_working_port'), nw_total_days or 0.0)
            nw_port_hrs = (nw_port_days * 24.0) if nw_port_days is not None else None

            nw_np_days = _safe_float(r.get('non_working_non_port'), 0.0)
            nw_np_hrs = (nw_np_days * 24.0) if nw_np_days is not None else None

            inward_days = _safe_float(r.get('inward_movement'))
            if inward_days is None and alongside and pilot_pick and alongside > pilot_pick:
                inward_days = (alongside - pilot_pick).total_seconds() / 86400.0
            inward_hrs = (inward_days * 24.0) if inward_days is not None else None

            outward_days = _safe_float(r.get('outward_movement'))
            if outward_days is None and pilot_dis and sail_cast and pilot_dis > sail_cast:
                outward_days = (pilot_dis - sail_cast).total_seconds() / 86400.0
            outward_hrs = (outward_days * 24.0) if outward_days is not None else None

            comp_unberth_days = _dt_diff_days(sail_cast, cargo_comp)
            comp_unberth_hrs = (comp_unberth_days * 24.0) if comp_unberth_days is not None else None

            tat_days = _dt_diff_days(pilot_dis, pilot_pick)
            tat_hrs = (tat_days * 24.0) if tat_days is not None else None

            productivity_day = round(flow_rate * 24.0, 2) if flow_rate else None

            # Financial year & month
            fy = (r.get('fin_year') or '').strip()
            month_val = (r.get('month') or '').strip()
            if not fy or not month_val:
                calc_fy, calc_m = _date_to_fin_year_and_month(ref_dt)
                fy = fy or calc_fy
                month_val = month_val or calc_m

            # Dimensions
            berth_no = r.get('berth_no') or ''
            dim_code = f"{berth_no}{status}"
            dim_via_fi_overseas_imo = f"{vcn} / {r.get('foreign_indian') or ''} / {r.get('overseas_coastal') or ''} / {r.get('imo_no') or ''}"
            dim_loa_bhc_dft = f"{r.get('loa') or ''} / {r.get('bhc') or ''} / {r.get('draft') or ''}"
            dim_agt_tf_pl = f"{r.get('agent') or ''} / {r.get('unloading_terminal') or ''} / {r.get('unload_pipeline') or ''}"

            sr_no = r.get('sr_no')

            hist_pc, hist_pol = _resolve_port_info(
                port_loading=r.get('port_of_loading'),
                existing_port_code=r.get('port_code'),
                port_code_map=port_code_map,
                port_name_map=port_name_map,
                known_codes=known_codes
            )

            hist_cargo = r.get('cargo') or ''
            if not (r.get('new_cat') and r.get('category1') and r.get('category')):
                f_nc, f_c1, f_c = _resolve_cargo_categories(hist_cargo, cargo_map)
                hist_new_cat = r.get('new_cat') or f_nc
                hist_cat1 = r.get('category1') or f_c1
                hist_cat = r.get('category') or f_c
            else:
                hist_new_cat = r.get('new_cat') or ''
                hist_cat1 = r.get('category1') or ''
                hist_cat = r.get('category') or ''

            records.append({
                'source': 'Master',
                'id': f"M-{r['id']}",
                'code': dim_code,
                'status': status,
                'sr_no': sr_no,
                'berth_no': berth_no,
                'vcn_no': vcn,
                'vessel_name': r.get('vessel_name') or '',
                'cat_1': None, 'cat_2': None, 'cat_3': None, 'cat_4': None, 'cat_5': None, 'cat_6': None,
                'cat_7': sr_no,
                'overseas_coastal': r.get('overseas_coastal') or '',
                'foreign_indian': r.get('foreign_indian') or '',
                'imo_no': r.get('imo_no') or '',
                'flag': r.get('flag') or '',
                'bhc': r.get('bhc') or '',
                'port_code': hist_pc,
                'port_of_loading': hist_pol,
                'grt': _safe_float(r.get('grt')),
                'draft': _safe_float(r.get('draft')),
                'loa': _safe_float(r.get('loa')),
                'import_export': r.get('import_export') or '',
                'agent': r.get('agent') or '',
                'unload_pipeline': r.get('unload_pipeline') or '',
                'consigner': r.get('consigner') or '',
                'unloading_terminal': r.get('unloading_terminal') or '',
                'new_cat': hist_new_cat,
                'category1': hist_cat1,
                'category': hist_cat,
                'cargo': hist_cargo,
                'bl_quantity': bl_qty,
                'operational_quantity': qty,
                'quantity': qty,
                'discharged_till_now': discharged_till_now,
                'till_time': _format_dt(cargo_comp or sail_cast),
                'expected_flow_rate': exp_flow_rate,
                'present_flow_rate': present_flow_rate,
                'flow_rate': flow_rate,
                'last_24h_qty': None,
                'balance_quantity': balance_qty,
                'balance_time_hrs': balance_time_hrs,
                'expected_completion': _format_dt(expected_completion),
                'cargo_status': 'Completed' if status == 'Unberth' else 'In Progress',
                'remarks': r.get('remarks') or '',
                # Timings
                'nor': _format_dt(nor),
                'anchorage_time': _format_dt(anchorage),
                'pilot_pickup': _format_dt(pilot_pick),
                'first_line': _format_dt(first_line),
                'alongside': _format_dt(alongside),
                'ops_commenced': _format_dt(ops_comm),
                'cargo_completion': _format_dt(cargo_comp),
                'sail_cast_off': _format_dt(sail_cast),
                'cast_off': _format_dt(cast_off),
                'pilot_board_departure': _format_dt(pilot_dept),
                'pilot_disembarked': _format_dt(pilot_dis),
                'last_updated': _format_dt(sail_cast or alongside),
                'current_status': status,
                # Dimensions
                'fin_year': fy,
                'month': month_val,
                'dim_via_fi_overseas_imo': dim_via_fi_overseas_imo,
                'dim_loa_bhc_dft': dim_loa_bhc_dft,
                'dim_agt_tf_pl': dim_agt_tf_pl,
                # Durations (Days & Hours)
                'pre_berthing_waiting_days': round(pbw_days, 2) if pbw_days is not None else None,
                'pre_berthing_waiting_hrs': round(pbw_hrs, 2) if pbw_hrs is not None else None,
                'waiting_port_days': round(waiting_port_days, 2) if waiting_port_days is not None else None,
                'waiting_port_hrs': round(waiting_port_hrs, 2) if waiting_port_hrs is not None else None,
                'waiting_non_port_days': round(waiting_np_days, 2) if waiting_np_days is not None else None,
                'waiting_non_port_hrs': round(waiting_np_hrs, 2) if waiting_np_hrs is not None else None,
                'stay_at_berth_days': round(stay_berth_days, 2) if stay_berth_days is not None else None,
                'stay_at_berth_hrs': round(stay_berth_hrs, 2) if stay_berth_hrs is not None else None,
                'arrive_to_comm_days': round(arr_comm_days, 2) if arr_comm_days is not None else None,
                'arrive_to_comm_hrs': round(arr_comm_hrs, 2) if arr_comm_hrs is not None else None,
                'working_time_days': round(working_time_days, 2) if working_time_days is not None else None,
                'working_time_hrs': round(working_time_hrs, 2) if working_time_hrs is not None else None,
                'non_working_total_days': round(nw_total_days, 2) if nw_total_days is not None else None,
                'non_working_total_hrs': round(nw_total_hrs, 2) if nw_total_hrs is not None else None,
                'non_working_port_days': round(nw_port_days, 2) if nw_port_days is not None else None,
                'non_working_port_hrs': round(nw_port_hrs, 2) if nw_port_hrs is not None else None,
                'non_working_non_port_days': round(nw_np_days, 2) if nw_np_days is not None else None,
                'non_working_non_port_hrs': round(nw_np_hrs, 2) if nw_np_hrs is not None else None,
                'inward_movement_days': round(inward_days, 2) if inward_days is not None else None,
                'inward_movement_hrs': round(inward_hrs, 2) if inward_hrs is not None else None,
                'outward_movement_days': round(outward_days, 2) if outward_days is not None else None,
                'outward_movement_hrs': round(outward_hrs, 2) if outward_hrs is not None else None,
                'comp_to_unberth_days': round(comp_unberth_days, 2) if comp_unberth_days is not None else None,
                'comp_to_unberth_hrs': round(comp_unberth_hrs, 2) if comp_unberth_hrs is not None else None,
                'productivity_per_day': productivity_day,
                'vessel_count': 1,
                'cargo_qty': qty,
                'turnaround_time_days': round(tat_days, 2) if tat_days is not None else None,
                'turnaround_time_hrs': round(tat_hrs, 2) if tat_hrs is not None else None,
                'turnaround_non_port_days': 0.0,
                'turnaround_non_port_hrs': 0.0,
                'ref_date': _format_date_only(ref_dt),
            })

        # -------------------------------------------------------------
        # 2. Live Operational Dataset: vcn_header + ldud_header + ldud_parcel_ops + lueu_parcel_log
        # -------------------------------------------------------------
        cur.execute("""
            WITH parcel_logs_agg AS (
                SELECT 
                    po.ldud_id,
                    SUM(CASE WHEN COALESCE(lg.is_shortclose, FALSE) = FALSE THEN lg.quantity ELSE 0 END) AS discharged_qty,
                    MAX(lg.entry_date || ' ' || lg.to_time) AS latest_log_time,
                    SUM(CASE 
                        WHEN lg.entry_date IS NOT NULL 
                         AND lg.entry_date::text >= (CURRENT_DATE - INTERVAL '1 day')::text
                        THEN lg.quantity ELSE 0 
                    END) AS last_24h_qty
                FROM ldud_parcel_ops po
                JOIN lueu_parcel_log lg ON lg.parcel_op_id = po.id
                WHERE lg.is_deleted IS NOT TRUE
                GROUP BY po.ldud_id
            ),
            parcels_agg AS (
                SELECT 
                    po.ldud_id,
                    MIN(NULLIF(TRIM(po.start_dt), '')) AS ops_commenced,
                    MAX(NULLIF(TRIM(po.end_dt), '')) AS cargo_completion,
                    SUM(po.quantity) AS operational_qty,
                    AVG(po.expected_flow_rate) AS exp_flow_rate,
                    string_agg(DISTINCT po.cargo_name, ', ') AS cargo_names,
                    string_agg(DISTINCT po.terminal_name, ', ') AS terminal_names
                FROM ldud_parcel_ops po
                GROUP BY po.ldud_id
            ),
            vcn_parcels AS (
                SELECT 
                    vcn_id,
                    SUM(NULLIF(regexp_replace(COALESCE(quantity, ''), '[^0-9.]', '', 'g'), '')::numeric) AS bl_qty,
                    string_agg(DISTINCT consigner_name, ', ') AS consigner_names,
                    string_agg(DISTINCT unload_terminal, ', ') AS vcn_terminals,
                    string_agg(DISTINCT pipeline_name, ', ') AS vcn_pipelines
                FROM (
                    SELECT vcn_id, quantity, consigner_name, unload_terminal, pipeline_name FROM vcn_consigners
                    UNION ALL
                    SELECT vcn_id, quantity, consigner_name, unload_terminal, pipeline_name FROM vcn_export_cargo_declaration
                ) vp
                GROUP BY vcn_id
            ),
            delays_agg AS (
                SELECT 
                    vd.vcn_id,
                    SUM(CASE 
                        WHEN LOWER(COALESCE(pdt.responsibility, '')) = 'agent' 
                          OR LOWER(COALESCE(pdt.type, '')) LIKE '%non-port%' 
                        THEN EXTRACT(EPOCH FROM (vd.delay_end::timestamp - vd.delay_start::timestamp)) / 3600.0
                        ELSE 0 
                    END) AS non_port_delay_hrs,
                    SUM(CASE 
                        WHEN LOWER(COALESCE(pdt.responsibility, 'port')) = 'port' 
                         AND (pdt.type IS NULL OR LOWER(pdt.type) NOT LIKE '%non-port%')
                        THEN EXTRACT(EPOCH FROM (vd.delay_end::timestamp - vd.delay_start::timestamp)) / 3600.0
                        ELSE 0 
                    END) AS port_delay_hrs
                FROM vcn_delays vd
                LEFT JOIN port_delay_types pdt ON LOWER(TRIM(pdt.name)) = LOWER(TRIM(vd.delay_name))
                WHERE vd.delay_start IS NOT NULL AND vd.delay_end IS NOT NULL
                GROUP BY vd.vcn_id
            )
            SELECT 
                vh.id AS vcn_id,
                COALESCE(vh.via_number, vh.vcn_doc_num) AS vcn_no,
                vh.vcn_doc_num,
                vh.doc_date,
                vh.vessel_name,
                vh.vessel_agent_name,
                vh.importer_exporter_name,
                vh.operation_type,
                vh.cargo_type,
                vh.load_port,
                vh.loa,
                vh.draft,
                vh.berth_name,
                vh.vessel_run_type,
                lh.id AS ldud_id,
                lh.doc_num AS ldud_doc_num,
                lh.doc_status AS ldud_status,
                lh.anchored_datetime,
                lh.arrival_inner_anchorage,
                lh.arrival_outer_anchorage,
                lh.nor_tendered,
                lh.discharge_commenced,
                lh.discharge_completed,
                lh.initial_draft_survey_quantity,
                lh.alongside_datetime,
                lh.cast_off_datetime,
                lh.pilot_board_departure,
                lh.pilot_disembarked,
                lh.first_line,
                lh.pilot_pickup_time,
                ves.imo_num,
                ves.nationality AS flag,
                ves.gt AS grt,
                ves.loa AS ves_loa,
                ves.vessel_category_name,
                pa.ops_commenced AS pa_ops_commenced,
                pa.cargo_completion AS pa_cargo_completion,
                pa.operational_qty,
                pa.exp_flow_rate,
                pa.cargo_names,
                pa.terminal_names,
                pla.discharged_qty,
                pla.latest_log_time,
                pla.last_24h_qty,
                vp.bl_qty,
                vp.consigner_names,
                vp.vcn_terminals,
                vp.vcn_pipelines,
                da.non_port_delay_hrs,
                da.port_delay_hrs
            FROM vcn_header vh
            LEFT JOIN ldud_header lh ON (lh.vcn_id = vh.id OR lh.vcn_doc_num = vh.vcn_doc_num) AND lh.is_deleted IS NOT TRUE
            LEFT JOIN vessels ves ON ves.doc_num = split_part(COALESCE(vh.vessel_master_doc, ''), '/', 1)
            LEFT JOIN parcels_agg pa ON pa.ldud_id = lh.id
            LEFT JOIN parcel_logs_agg pla ON pla.ldud_id = lh.id
            LEFT JOIN vcn_parcels vp ON vp.vcn_id = vh.id
            LEFT JOIN delays_agg da ON da.vcn_id = vh.id
            ORDER BY vh.id DESC
        """)
        live_rows = cur.fetchall()

        next_sr_no = (len(m_rows) + 1)
        for r in live_rows:
            vcn = (r.get('vcn_no') or '').strip().upper()
            if vcn and vcn in seen_vcns:
                # Skip duplicate if already present in historical dataset
                continue

            # Timings
            nor = _parse_dt(r.get('nor_tendered'))
            anchorage = _parse_dt(r.get('arrival_inner_anchorage')) or _parse_dt(r.get('anchored_datetime')) or _parse_dt(r.get('arrival_outer_anchorage')) or nor
            pilot_pick = _parse_dt(r.get('pilot_pickup_time'))
            first_line = _parse_dt(r.get('first_line'))
            alongside = _parse_dt(r.get('alongside_datetime'))
            ops_comm = _parse_dt(r.get('pa_ops_commenced')) or _parse_dt(r.get('discharge_commenced'))
            cargo_comp = _parse_dt(r.get('pa_cargo_completion')) or _parse_dt(r.get('discharge_completed'))
            sail_cast = _parse_dt(r.get('cast_off_datetime'))
            cast_off = sail_cast
            pilot_dept = _parse_dt(r.get('pilot_board_departure'))
            pilot_dis = _parse_dt(r.get('pilot_disembarked'))

            # Status determination
            status = 'Unberth' if sail_cast else ('Berth' if alongside else 'Waiting')

            # Reference date
            ref_dt = sail_cast or alongside or ops_comm or nor or _parse_dt(r.get('doc_date'))

            # Quantities
            bl_qty = _safe_float(r.get('bl_qty'), 0.0)
            op_qty = _safe_float(r.get('operational_qty')) or _safe_float(r.get('initial_draft_survey_quantity'), bl_qty)
            discharged_qty = _safe_float(r.get('discharged_qty'), 0.0)
            last_24h = _safe_float(r.get('last_24h_qty'), 0.0)

            # If vessel completed, discharged quantity is at least the operational quantity
            if status == 'Unberth' and discharged_qty <= 0:
                discharged_qty = op_qty

            balance_qty = max(0.0, op_qty - discharged_qty)

            # Flow rates
            working_time_days = _dt_diff_days(cargo_comp, ops_comm)
            working_time_hrs = (working_time_days * 24.0) if working_time_days is not None else None

            # Current operational elapsed working hours
            now_dt = datetime.now()
            till_dt = cargo_comp if status == 'Unberth' else (_parse_dt(r.get('latest_log_time')) or now_dt)

            flow_rate = None
            if working_time_hrs and working_time_hrs > 0 and discharged_qty > 0:
                flow_rate = round(discharged_qty / working_time_hrs, 2)

            present_flow_rate = None
            if ops_comm and till_dt and till_dt > ops_comm:
                elapsed_hrs = (till_dt - ops_comm).total_seconds() / 3600.0
                if elapsed_hrs > 0 and discharged_qty > 0:
                    present_flow_rate = round(discharged_qty / elapsed_hrs, 2)

            exp_flow_rate = _safe_float(r.get('exp_flow_rate')) or flow_rate or 350.0

            balance_time_hrs = None
            if balance_qty > 0 and exp_flow_rate and exp_flow_rate > 0:
                balance_time_hrs = round(balance_qty / exp_flow_rate, 2)

            expected_completion = None
            if balance_time_hrs is not None and till_dt:
                expected_completion = till_dt + timedelta(hours=balance_time_hrs)

            # Operational Durations
            pbw_days = _dt_diff_days(alongside, nor)
            pbw_hrs = (pbw_days * 24.0) if pbw_days is not None else None

            waiting_np_hrs = _safe_float(r.get('non_port_delay_hrs'), 0.0)
            waiting_np_days = waiting_np_hrs / 24.0

            waiting_port_hrs = max(0.0, (pbw_hrs - waiting_np_hrs)) if pbw_hrs is not None else 0.0
            waiting_port_days = waiting_port_hrs / 24.0

            stay_berth_days = _dt_diff_days(sail_cast, alongside)
            stay_berth_hrs = (stay_berth_days * 24.0) if stay_berth_days is not None else None

            arr_comm_days = _dt_diff_days(ops_comm, alongside)
            arr_comm_hrs = (arr_comm_days * 24.0) if arr_comm_days is not None else None

            nw_total_days = max(0.0, (stay_berth_days - working_time_days)) if (stay_berth_days is not None and working_time_days is not None) else None
            nw_total_hrs = (nw_total_days * 24.0) if nw_total_days is not None else None

            nw_port_hrs = _safe_float(r.get('port_delay_hrs'), nw_total_hrs or 0.0)
            nw_port_days = nw_port_hrs / 24.0

            nw_np_hrs = max(0.0, (nw_total_hrs - nw_port_hrs)) if nw_total_hrs is not None else 0.0
            nw_np_days = nw_np_hrs / 24.0

            inward_days = _dt_diff_days(alongside, pilot_pick)
            inward_hrs = (inward_days * 24.0) if inward_days is not None else None

            outward_days = _dt_diff_days(pilot_dis, sail_cast)
            outward_hrs = (outward_days * 24.0) if outward_days is not None else None

            comp_unberth_days = _dt_diff_days(sail_cast, cargo_comp)
            comp_unberth_hrs = (comp_unberth_days * 24.0) if comp_unberth_days is not None else None

            tat_days = _dt_diff_days(pilot_dis, pilot_pick)
            tat_hrs = (tat_days * 24.0) if tat_days is not None else None

            effective_rate = flow_rate or present_flow_rate or exp_flow_rate
            productivity_day = round(effective_rate * 24.0, 2) if effective_rate else None

            # Financial Year & Month
            calc_fy, calc_m = _date_to_fin_year_and_month(ref_dt)
            fy = calc_fy
            month_val = calc_m

            # Dimensions
            berth_no = r.get('berth_name') or ''
            dim_code = f"{berth_no}{status}"
            loa_val = _safe_float(r.get('loa')) or _safe_float(r.get('ves_loa'))
            dft_val = _safe_float(r.get('draft'))
            imo_val = r.get('imo_num') or ''
            flag_val = r.get('flag') or ''
            agent_val = r.get('vessel_agent_name') or ''
            terminal_val = r.get('terminal_names') or r.get('vcn_terminals') or berth_no
            pipeline_val = r.get('vcn_pipelines') or ''
            consigner_val = r.get('consigner_names') or r.get('importer_exporter_name') or ''
            cargo_val = r.get('cargo_names') or r.get('cargo_type') or ''
            live_new_cat, live_cat1, live_cat = _resolve_cargo_categories(cargo_val, cargo_map)
            vessel_run = r.get('vessel_run_type') or 'Overseas'

            dim_via_fi_overseas_imo = f"{vcn} / F / {vessel_run} / {imo_val}"
            dim_loa_bhc_dft = f"{loa_val or ''} /  / {dft_val or ''}"
            dim_agt_tf_pl = f"{agent_val} / {terminal_val} / {pipeline_val}"

            live_load_port = r.get('load_port') or (r.get('discharge_port') if (r.get('operation_type') or '').lower() == 'export' else '')
            live_pc, live_pol = _resolve_port_info(
                port_loading=live_load_port,
                existing_port_code='',
                port_code_map=port_code_map,
                port_name_map=port_name_map,
                known_codes=known_codes
            )

            records.append({
                'source': 'Actual',
                'id': f"A-{r['vcn_id']}",
                'code': dim_code,
                'status': status,
                'sr_no': next_sr_no,
                'berth_no': berth_no,
                'vcn_no': vcn,
                'vessel_name': r.get('vessel_name') or '',
                'cat_1': None, 'cat_2': None, 'cat_3': None, 'cat_4': None, 'cat_5': None, 'cat_6': None,
                'cat_7': next_sr_no,
                'overseas_coastal': vessel_run,
                'foreign_indian': 'F',
                'imo_no': imo_val,
                'flag': flag_val,
                'bhc': '',
                'port_code': live_pc,
                'port_of_loading': live_pol,
                'grt': _safe_float(r.get('grt')),
                'draft': dft_val,
                'loa': loa_val,
                'import_export': r.get('operation_type') or 'Import',
                'agent': agent_val,
                'unload_pipeline': pipeline_val,
                'consigner': consigner_val,
                'unloading_terminal': terminal_val,
                'new_cat': live_new_cat,
                'category1': live_cat1,
                'category': live_cat,
                'cargo': cargo_val,
                'bl_quantity': bl_qty,
                'operational_quantity': op_qty,
                'quantity': op_qty,
                'discharged_till_now': discharged_qty,
                'till_time': _format_dt(till_dt),
                'expected_flow_rate': exp_flow_rate,
                'present_flow_rate': present_flow_rate,
                'flow_rate': flow_rate,
                'last_24h_qty': last_24h,
                'balance_quantity': balance_qty,
                'balance_time_hrs': balance_time_hrs,
                'expected_completion': _format_dt(expected_completion),
                'cargo_status': 'Completed' if status == 'Unberth' else ('In Progress' if status == 'Berth' else 'Pending'),
                'remarks': r.get('ldud_doc_num') or '',
                # Timings
                'nor': _format_dt(nor),
                'anchorage_time': _format_dt(anchorage),
                'pilot_pickup': _format_dt(pilot_pick),
                'first_line': _format_dt(first_line),
                'alongside': _format_dt(alongside),
                'ops_commenced': _format_dt(ops_comm),
                'cargo_completion': _format_dt(cargo_comp),
                'sail_cast_off': _format_dt(sail_cast),
                'cast_off': _format_dt(cast_off),
                'pilot_board_departure': _format_dt(pilot_dept),
                'pilot_disembarked': _format_dt(pilot_dis),
                'last_updated': _format_dt(till_dt),
                'current_status': status,
                # Dimensions
                'fin_year': fy,
                'month': month_val,
                'dim_via_fi_overseas_imo': dim_via_fi_overseas_imo,
                'dim_loa_bhc_dft': dim_loa_bhc_dft,
                'dim_agt_tf_pl': dim_agt_tf_pl,
                # Durations (Days & Hours)
                'pre_berthing_waiting_days': round(pbw_days, 2) if pbw_days is not None else None,
                'pre_berthing_waiting_hrs': round(pbw_hrs, 2) if pbw_hrs is not None else None,
                'waiting_port_days': round(waiting_port_days, 2) if waiting_port_days is not None else None,
                'waiting_port_hrs': round(waiting_port_hrs, 2) if waiting_port_hrs is not None else None,
                'waiting_non_port_days': round(waiting_np_days, 2) if waiting_np_days is not None else None,
                'waiting_non_port_hrs': round(waiting_np_hrs, 2) if waiting_np_hrs is not None else None,
                'stay_at_berth_days': round(stay_berth_days, 2) if stay_berth_days is not None else None,
                'stay_at_berth_hrs': round(stay_berth_hrs, 2) if stay_berth_hrs is not None else None,
                'arrive_to_comm_days': round(arr_comm_days, 2) if arr_comm_days is not None else None,
                'arrive_to_comm_hrs': round(arr_comm_hrs, 2) if arr_comm_hrs is not None else None,
                'working_time_days': round(working_time_days, 2) if working_time_days is not None else None,
                'working_time_hrs': round(working_time_hrs, 2) if working_time_hrs is not None else None,
                'non_working_total_days': round(nw_total_days, 2) if nw_total_days is not None else None,
                'non_working_total_hrs': round(nw_total_hrs, 2) if nw_total_hrs is not None else None,
                'non_working_port_days': round(nw_port_days, 2) if nw_port_days is not None else None,
                'non_working_port_hrs': round(nw_port_hrs, 2) if nw_port_hrs is not None else None,
                'non_working_non_port_days': round(nw_np_days, 2) if nw_np_days is not None else None,
                'non_working_non_port_hrs': round(nw_np_hrs, 2) if nw_np_hrs is not None else None,
                'inward_movement_days': round(inward_days, 2) if inward_days is not None else None,
                'inward_movement_hrs': round(inward_hrs, 2) if inward_hrs is not None else None,
                'outward_movement_days': round(outward_days, 2) if outward_days is not None else None,
                'outward_movement_hrs': round(outward_hrs, 2) if outward_hrs is not None else None,
                'comp_to_unberth_days': round(comp_unberth_days, 2) if comp_unberth_days is not None else None,
                'comp_to_unberth_hrs': round(comp_unberth_hrs, 2) if comp_unberth_hrs is not None else None,
                'productivity_per_day': productivity_day,
                'vessel_count': 1,
                'cargo_qty': op_qty,
                'turnaround_time_days': round(tat_days, 2) if tat_days is not None else None,
                'turnaround_time_hrs': round(tat_hrs, 2) if tat_hrs is not None else None,
                'turnaround_non_port_days': 0.0,
                'turnaround_non_port_hrs': 0.0,
                'ref_date': _format_date_only(ref_dt),
            })
            next_sr_no += 1

    finally:
        conn.close()

    return records


def filter_and_aggregate_records(records, args):
    """Apply multi-criteria filters, calculate summary KPIs, and return filtered rows."""
    date_from = args.get('date_from', '').strip()
    date_to = args.get('date_to', '').strip()
    fin_year = (args.get('fin_year') or args.get('financial_year') or 'ALL').strip()
    month = args.get('month', 'ALL').strip()
    berth = args.get('berth', 'ALL').strip()
    cargo = args.get('cargo', 'ALL').strip()
    agent = args.get('agent', 'ALL').strip()
    terminal = args.get('terminal', 'ALL').strip()
    import_export = args.get('import_export', 'ALL').strip()
    status = args.get('status', 'ALL').strip()
    search = args.get('search', '').strip().lower()

    filtered = []
    for r in records:
        if date_from and r['ref_date'] and r['ref_date'] < date_from:
            continue
        if date_to and r['ref_date'] and r['ref_date'] > date_to:
            continue
        if fin_year != 'ALL' and r['fin_year'] != fin_year:
            continue
        if month != 'ALL' and r['month'].lower() != month.lower():
            continue
        if berth != 'ALL' and r['berth_no'].lower() != berth.lower():
            continue
        if cargo != 'ALL' and cargo.lower() not in (r['cargo'] or '').lower():
            continue
        if agent != 'ALL' and agent.lower() not in (r['agent'] or '').lower():
            continue
        if terminal != 'ALL' and terminal.lower() not in (r['unloading_terminal'] or '').lower():
            continue
        if import_export != 'ALL' and r['import_export'].lower() != import_export.lower():
            continue
        if status != 'ALL' and r['status'].lower() != status.lower():
            continue

        if search:
            match_str = f"{r['vessel_name']} {r['vcn_no']} {r['agent']} {r['cargo']} {r['berth_no']} {r['consigner']} {r['code']} {r['port_code']} {r['port_of_loading']} {r['new_cat']} {r['category1']} {r['category']}".lower()
            if search not in match_str:
                continue

        filtered.append(r)

    # Sort records
    sort_field = args.get('sort_field', 'sr_no')
    sort_dir = args.get('sort_dir', 'asc').lower()
    reverse = (sort_dir == 'desc')

    def sort_key(row):
        val = row.get(sort_field)
        if val is None:
            return -1 if reverse else float('inf')
        return val

    try:
        filtered.sort(key=sort_key, reverse=reverse)
    except Exception:
        pass

    # Calculate summary KPIs
    total_qty = sum(r['quantity'] or 0.0 for r in filtered)
    import_qty = sum(r['quantity'] or 0.0 for r in filtered if r['import_export'].lower() == 'import')
    export_qty = sum(r['quantity'] or 0.0 for r in filtered if r['import_export'].lower() == 'export')
    total_balance = sum(r['balance_quantity'] or 0.0 for r in filtered)

    flow_rates = [r['flow_rate'] for r in filtered if r['flow_rate']]
    avg_flow_rate = round(sum(flow_rates) / len(flow_rates), 2) if flow_rates else 0.0

    tats = [r['turnaround_time_hrs'] for r in filtered if r['turnaround_time_hrs']]
    avg_tat_hrs = round(sum(tats) / len(tats), 2) if tats else 0.0

    pbws = [r['pre_berthing_waiting_hrs'] for r in filtered if r['pre_berthing_waiting_hrs']]
    avg_pbw_hrs = round(sum(pbws) / len(pbws), 2) if pbws else 0.0

    stays = [r['stay_at_berth_hrs'] for r in filtered if r['stay_at_berth_hrs']]
    avg_stay_hrs = round(sum(stays) / len(stays), 2) if stays else 0.0

    summary = {
        'total_vessels': len(filtered),
        'total_quantity': round(total_qty, 3),
        'import_quantity': round(import_qty, 3),
        'export_quantity': round(export_qty, 3),
        'total_balance': round(total_balance, 3),
        'avg_flow_rate': avg_flow_rate,
        'avg_turnaround_time_hrs': avg_tat_hrs,
        'avg_pre_berthing_waiting_hrs': avg_pbw_hrs,
        'avg_stay_at_berth_hrs': avg_stay_hrs,
    }

    return filtered, summary


# ──────────────────────────────────────────────────────────────────
#  Routes & APIs
# ──────────────────────────────────────────────────────────────────

@bp.route('/module/RP01/master-report/')
@login_required
def master_report_page():
    perms = get_perms()
    if not perms.get('can_read'):
        return render_template('no_access.html'), 403
    return render_template('master_report/master_report.html', username=session.get('username'), permissions=perms)


@bp.route('/api/module/RP01/master-report/meta')
@login_required
def master_report_meta():
    records = fetch_master_report_records()
    months_by_fy = {}
    for r in records:
        fy = r.get('fin_year')
        m = r.get('month')
        if fy and m:
            months_by_fy.setdefault(fy, set()).add(m)

    def month_sort_key(m_str):
        parts = m_str.split('-')
        m_name = parts[0]
        y_val = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 0
        m_idx = MONTH_NAMES.index(m_name) if m_name in MONTH_NAMES else 99
        return (y_val, m_idx)

    fin_years = sorted(list({r['fin_year'] for r in records if r['fin_year']}), reverse=True)
    all_months = sorted(list({r['month'] for r in records if r['month']}), key=month_sort_key)
    sorted_months_by_fy = {
        fy: sorted(list(mset), key=month_sort_key)
        for fy, mset in months_by_fy.items()
    }

    return jsonify({
        'fin_years': fin_years,
        'months': all_months,
        'months_by_fy': sorted_months_by_fy,
    })


@bp.route('/api/module/RP01/master-report/data')
@login_required
def master_report_data():
    records = fetch_master_report_records()
    filtered, summary = filter_and_aggregate_records(records, request.args)

    page_str = request.args.get('page', '1')
    size_str = request.args.get('size', '50')

    total_rows = len(filtered)

    if size_str.lower() == 'all':
        paged_rows = filtered
        page = 1
        last_page = 1
    else:
        try:
            page = max(1, int(page_str))
            size = max(1, int(size_str))
        except ValueError:
            page = 1
            size = 50

        start_idx = (page - 1) * size
        end_idx = start_idx + size
        paged_rows = filtered[start_idx:end_idx]
        last_page = max(1, math.ceil(total_rows / size))

    return jsonify({
        'data': paged_rows,
        'summary': summary,
        'total': total_rows,
        'page': page,
        'last_page': last_page,
    })


@bp.route('/api/module/RP01/master-report/export')
@login_required
def master_report_export():
    records = fetch_master_report_records()
    filtered, summary = filter_and_aggregate_records(records, request.args)

    wb = Workbook()
    ws = wb.active
    ws.title = "Master Report"

    # Color Palette & Styles
    header_font_title = Font(name="Calibri", size=14, bold=True, color="1F4E78")
    section_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    col_font = Font(name="Calibri", size=9, bold=True, color="000000")
    data_font = Font(name="Calibri", size=9)

    border_thin = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    fill_sec_basic = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_sec_vessel = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    fill_sec_cargo = PatternFill(start_color="1B75BC", end_color="1B75BC", fill_type="solid")
    fill_sec_timing = PatternFill(start_color="1E824C", end_color="1E824C", fill_type="solid")
    fill_sec_rate = PatternFill(start_color="D35400", end_color="D35400", fill_type="solid")
    fill_sec_dim = PatternFill(start_color="8E44AD", end_color="8E44AD", fill_type="solid")
    fill_sec_perf = PatternFill(start_color="27AE60", end_color="27AE60", fill_type="solid")
    fill_hdr = PatternFill(start_color="F2F4F7", end_color="F2F4F7", fill_type="solid")

    # Row 1: Report Title
    ws.merge_cells("A1:K1")
    ws["A1"] = "JAWAHARLAL NEHRU PORT AUTHORITY - MASTER REPORT"
    ws["A1"].font = header_font_title

    # Section Headers (Row 2) and Column Headers (Row 3)
    columns_def = [
        # Basic (1-6)
        ("Basic Details", fill_sec_basic, "Code", "code"),
        ("Basic Details", fill_sec_basic, "Status", "status"),
        ("Basic Details", fill_sec_basic, "LOA", "loa"),
        ("Basic Details", fill_sec_basic, "Berth No", "berth_no"),
        ("Basic Details", fill_sec_basic, "Via", "vcn_no"),
        ("Basic Details", fill_sec_basic, "Vessel Name", "vessel_name"),
        # Vessel Details (7-15)
        ("Vessel Details", fill_sec_vessel, "Overseas", "overseas_coastal"),
        ("Vessel Details", fill_sec_vessel, "F/I", "foreign_indian"),
        ("Vessel Details", fill_sec_vessel, "IMO NO", "imo_no"),
        ("Vessel Details", fill_sec_vessel, "Flag", "flag"),
        ("Vessel Details", fill_sec_vessel, "BHC", "bhc"),
        ("Vessel Details", fill_sec_vessel, "Port Code", "port_code"),
        ("Vessel Details", fill_sec_vessel, "Port of Loading", "port_of_loading"),
        ("Vessel Details", fill_sec_vessel, "GRT", "grt"),
        ("Vessel Details", fill_sec_vessel, "Draft (m)", "draft"),
        # Cargo Details (16-25)
        ("Importer and Cargo Details", fill_sec_cargo, "Import/Export", "import_export"),
        ("Importer and Cargo Details", fill_sec_cargo, "Agent", "agent"),
        ("Importer and Cargo Details", fill_sec_cargo, "Unload Pipeline", "unload_pipeline"),
        ("Importer and Cargo Details", fill_sec_cargo, "Consigner", "consigner"),
        ("Importer and Cargo Details", fill_sec_cargo, "Unloading Terminal", "unloading_terminal"),
        ("Importer and Cargo Details", fill_sec_cargo, "New Cat", "new_cat"),
        ("Importer and Cargo Details", fill_sec_cargo, "Category-1", "category1"),
        ("Importer and Cargo Details", fill_sec_cargo, "Category", "category"),
        ("Importer and Cargo Details", fill_sec_cargo, "Cargo", "cargo"),
        ("Importer and Cargo Details", fill_sec_cargo, "Quantity (MT)", "quantity"),
        # Vessel Timing Details (27-37)
        ("Vessel Timing Details", fill_sec_timing, "NOR", "nor"),
        ("Vessel Timing Details", fill_sec_timing, "Anchorage Time", "anchorage_time"),
        ("Vessel Timing Details", fill_sec_timing, "Pilot Pick Up", "pilot_pickup"),
        ("Vessel Timing Details", fill_sec_timing, "First Line", "first_line"),
        ("Vessel Timing Details", fill_sec_timing, "Alongside", "alongside"),
        ("Vessel Timing Details", fill_sec_timing, "Ops Commenced", "ops_commenced"),
        ("Vessel Timing Details", fill_sec_timing, "Cargo Completion", "cargo_completion"),
        ("Vessel Timing Details", fill_sec_timing, "Sail Cast Off", "sail_cast_off"),
        ("Vessel Timing Details", fill_sec_timing, "Cast Off", "cast_off"),
        ("Vessel Timing Details", fill_sec_timing, "Pilot Board Departure", "pilot_board_departure"),
        ("Vessel Timing Details", fill_sec_timing, "Pilot Disembarked", "pilot_disembarked"),
        # Quantity, Rate & Forecast (38-46)
        ("Qty & Rate Forecast", fill_sec_rate, "Discharged Till Now (MT)", "discharged_till_now"),
        ("Qty & Rate Forecast", fill_sec_rate, "Till Time", "till_time"),
        ("Qty & Rate Forecast", fill_sec_rate, "Expected Flow Rate", "expected_flow_rate"),
        ("Qty & Rate Forecast", fill_sec_rate, "Flow Rate (MT/hr)", "flow_rate"),
        ("Qty & Rate Forecast", fill_sec_rate, "Present Flow Rate", "present_flow_rate"),
        ("Qty & Rate Forecast", fill_sec_rate, "Balance (MT)", "balance_quantity"),
        ("Qty & Rate Forecast", fill_sec_rate, "Balance Time (hrs)", "balance_time_hrs"),
        ("Qty & Rate Forecast", fill_sec_rate, "Expected Completion", "expected_completion"),
        ("Qty & Rate Forecast", fill_sec_rate, "Last 24 hrs Qty", "last_24h_qty"),
        # Classification & Dimensions (47-51)
        ("Reporting Dimensions", fill_sec_dim, "FY", "fin_year"),
        ("Reporting Dimensions", fill_sec_dim, "Month", "month"),
        ("Reporting Dimensions", fill_sec_dim, "VIA/F/OVERSEAS/IMO", "dim_via_fi_overseas_imo"),
        ("Reporting Dimensions", fill_sec_dim, "LOA/BHC/DFT", "dim_loa_bhc_dft"),
        ("Reporting Dimensions", fill_sec_dim, "AGT / TF / PL", "dim_agt_tf_pl"),
        # Turnaround & Operational Performance (52-67)
        ("Turnaround & Operational Performance", fill_sec_perf, "Pre-Berthing Waiting (Days)", "pre_berthing_waiting_days"),
        ("Turnaround & Operational Performance", fill_sec_perf, "Pre-Berthing Waiting (Hrs)", "pre_berthing_waiting_hrs"),
        ("Turnaround & Operational Performance", fill_sec_perf, "Port Delay (Hrs)", "waiting_port_hrs"),
        ("Turnaround & Operational Performance", fill_sec_perf, "Non-Port Delay (Hrs)", "waiting_non_port_hrs"),
        ("Turnaround & Operational Performance", fill_sec_perf, "Stay at Berth (Days)", "stay_at_berth_days"),
        ("Turnaround & Operational Performance", fill_sec_perf, "Stay at Berth (Hrs)", "stay_at_berth_hrs"),
        ("Turnaround & Operational Performance", fill_sec_perf, "Arrive to Comm (Hrs)", "arrive_to_comm_hrs"),
        ("Turnaround & Operational Performance", fill_sec_perf, "Working Time (Days)", "working_time_days"),
        ("Turnaround & Operational Performance", fill_sec_perf, "Working Time (Hrs)", "working_time_hrs"),
        ("Turnaround & Operational Performance", fill_sec_perf, "Non-Working Time (Hrs)", "non_working_total_hrs"),
        ("Turnaround & Operational Performance", fill_sec_perf, "Inward Movement (Hrs)", "inward_movement_hrs"),
        ("Turnaround & Operational Performance", fill_sec_perf, "Outward Movement (Hrs)", "outward_movement_hrs"),
        ("Turnaround & Operational Performance", fill_sec_perf, "Comp to Unberth (Hrs)", "comp_to_unberth_hrs"),
        ("Turnaround & Operational Performance", fill_sec_perf, "Productivity per Day (MT)", "productivity_per_day"),
        ("Turnaround & Operational Performance", fill_sec_perf, "Turnaround Time (Days)", "turnaround_time_days"),
        ("Turnaround & Operational Performance", fill_sec_perf, "Turnaround Time (Hrs)", "turnaround_time_hrs"),
    ]

    # Write Section & Column Headers
    curr_sec = None
    sec_start_col = 1
    for col_idx, (sec_name, sec_fill, col_title, _) in enumerate(columns_def, start=1):
        # Header cell
        c_cell = ws.cell(row=3, column=col_idx, value=col_title)
        c_cell.font = col_font
        c_cell.fill = fill_hdr
        c_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c_cell.border = border_thin

        # Section merging
        if curr_sec != sec_name:
            if curr_sec is not None and col_idx - 1 >= sec_start_col:
                ws.merge_cells(start_row=2, start_column=sec_start_col, end_row=2, end_column=col_idx - 1)
                s_cell = ws.cell(row=2, column=sec_start_col)
                s_cell.value = curr_sec
                s_cell.font = section_font
                s_cell.fill = prev_sec_fill
                s_cell.alignment = Alignment(horizontal="center", vertical="center")
            curr_sec = sec_name
            sec_start_col = col_idx
            prev_sec_fill = sec_fill

    # Merge last section
    if curr_sec is not None and len(columns_def) >= sec_start_col:
        ws.merge_cells(start_row=2, start_column=sec_start_col, end_row=2, end_column=len(columns_def))
        s_cell = ws.cell(row=2, column=sec_start_col)
        s_cell.value = curr_sec
        s_cell.font = section_font
        s_cell.fill = prev_sec_fill
        s_cell.alignment = Alignment(horizontal="center", vertical="center")

    # Data Rows
    row_idx = 4
    for r in filtered:
        for col_idx, (_, _, _, field_key) in enumerate(columns_def, start=1):
            val = r.get(field_key)
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = data_font
            cell.border = border_thin
            if isinstance(val, (int, float)):
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
        row_idx += 1

    # Auto-adjust column widths
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max(len(str(cell.value or '')) for cell in col[:15])
        ws.column_dimensions[col_letter].width = max(11, min(max_len + 3, 30))

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)

    filename = f"Master_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(
        bio,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename,
    )
