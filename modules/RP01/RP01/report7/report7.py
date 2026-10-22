"""
Report-7 — Performance in Terms of Vessels Handled - All Vessels
Flask Blueprint version. Reads from mis_vessel_master (historical) and falls back
to live operational pipeline (lueu_parcel_log + ldud_header + vcn_header) for post-June open periods.

Port: JAWAHARLAL NEHRU PORT AUTHORITY
Columns: Liquid Bulk | All

Section Breakdown:
  1. No. of Vessels Sailed- (Total No.)
     a) At Berth- (No.)
        (i) Overseas
        (ii) Coastal
     b) At Stream- (No.) (Inn Anchorage)
  2. Total Cargo Traffic Handled - (000 Tonnes)
     a) At Berth - (000 Tonnes)
     b) At Stream (Inn Anchorage) - (000 Tonnes)
  3. Pre Berthing Waiting Time(Hrs)-Total
     a) Pre Berthing Waiting Time(Hrs)-Port a/c
     b) Pre Berthing Waiting Time(Hrs)-Non Port a/c
  4. Working Time (Hrs)  [Formula: (first start parcel - last end parcel)]
  5. N.W Time at working berth(Hrs)-Total  [Formula: (Alongside - Castoff in hrs) - Sr No 4]
     a) N.W Time at working berth(Hrs)-Port a/c
     b) N.W Time at working berth(Hrs)-N.P a/c
  6. Time at Non working berth(Hrs)-Total
     a) Time at Non working berth(Hrs)-Port a/c
     b) Time at Non working berth(Hrs)-N.P a/c
  7. Navigation Time (Hrs)-Total
     a) Inward Movement (Hrs)  [Pilot pick up to alongside]
     b) Outward Movement (Hrs) [Pilot disembark to castoff]
  8. Shifting Time
  9. Turn Round Time (Hrs)-Total=(3+4+5+6)
     a) Turn Round Time (Hrs)-Port a/c=(3(a)+4+5(a)+6(a)+7+8)
     b) Turn Round Time (Hrs)-N.P a/c=(3(b)+5(b)+6(b))
"""

import io
import math
import traceback
from functools import wraps
from datetime import datetime, date

import pandas as pd
from flask import jsonify, request, render_template, send_file, session, redirect, url_for
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

from database import get_db, get_cursor
from .. import bp


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated


MONTH_NAMES = ["Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"]


def fy_start_year(fin_year: str) -> int:
    try:
        return int(fin_year.split("-")[0])
    except Exception:
        return datetime.now().year


def month_options_for(fin_year: str):
    start_y = fy_start_year(fin_year)
    opts = []
    for idx, mn in enumerate(MONTH_NAMES):
        yy = start_y if idx < 9 else start_y + 1
        opts.append({"idx": idx, "label": f"{mn}-{str(yy % 100).zfill(2)}"})
    return opts


def _parse_dt(s):
    if not s:
        return None
    if isinstance(s, datetime):
        return s
    if isinstance(s, date):
        return datetime(s.year, s.month, s.day)
    clean = str(s).strip().replace('T', ' ')
    for length in (19, 16, 10):
        sub = clean[:length]
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


def _merge_and_sum_intervals(intervals):
    """
    Given a list of (start_datetime, end_datetime) tuples,
    merges overlapping and contiguous time intervals to avoid duplicate counting,
    and returns the total unique duration in hours.
    """
    valid = [(s, e) for s, e in intervals if s and e and e > s]
    if not valid:
        return 0.0

    valid.sort(key=lambda x: (x[0], x[1]))

    merged = []
    cur_start, cur_end = valid[0]

    for s, e in valid[1:]:
        if s <= cur_end:
            if e > cur_end:
                cur_end = e
        else:
            merged.append((cur_start, cur_end))
            cur_start, cur_end = s, e

    merged.append((cur_start, cur_end))
    total_seconds = sum((e - s).total_seconds() for s, e in merged)
    return total_seconds / 3600.0


def _is_milestone_duplicate(existing_delays, ms_start, ms_end, ms_type="berth"):
    """
    Milestone deduplication logic matching vessel_delay.py:
    1. Specialized overlap check for Berth Not Available / Port Congestion:
       If an original delay has 'berth' & ('notavail', 'unavail', 'occup') or 'congestion',
       and its interval overlaps the milestone interval, mark as duplicate.
    2. Universal exact-match check:
       If an original delay matches milestone start and end times within 1 hour,
       mark as duplicate.
    """
    if not ms_start or not ms_end or not existing_delays:
        return False
    for existing in existing_delays:
        ex_start = existing.get("start")
        ex_end = existing.get("end")
        ex_name = str(existing.get("delay_name") or existing.get("name") or "").lower().replace(" ", "")

        if ex_start and ex_end:
            # 1. Specialized overlap check for Berth Not Available / Port Congestion
            if ms_type == "berth":
                has_berth = "berth" in ex_name
                has_congestion = "congestion" in ex_name
                is_unavailable = any(x in ex_name for x in ["notavail", "unavail", "notavaileb", "occup"])
                if (has_berth and is_unavailable) or has_congestion:
                    if ex_start <= ms_end and ms_start <= ex_end:
                        return True

            # 2. Universal exact-match check for ANY milestone (within 1 hour)
            start_diff = abs((ex_start - ms_start).total_seconds())
            end_diff = abs((ex_end - ms_end).total_seconds())
            if start_diff <= 3600 and end_diff <= 3600:
                return True
    return False



def _parse_delay_hours(h_str):
    """Convert delay duration to decimal hours."""
    if h_str is None or str(h_str).strip() == "":
        return 0.0

    s = str(h_str).strip()

    try:
        # HH:MM or HH:MM:SS
        if ":" in s:
            parts = s.split(":")
            hours = int(parts[0])
            minutes = int(parts[1])
            seconds = float(parts[2]) if len(parts) > 2 else 0.0
            return hours + minutes / 60.0 + seconds / 3600.0

        # Vessel Delay Report format: HH.MM, where digits
        # after the dot represent minutes, not decimal fractions.
        if "." in s:
            hours_part, minutes_part = s.split(".", 1)
            hours = int(hours_part)
            minutes = int(minutes_part)

            if 0 <= minutes < 60:
                return hours + minutes / 60.0

        return float(s)

    except (ValueError, TypeError):
        return 0.0




def _load_live_pipeline_rows():
    conn = get_db()

    try:
        cur = get_cursor(conn)
        cur.execute("""
            SELECT
                ld.id AS ldud_id,
                ld.cast_off_datetime,
                ld.alongside_datetime,
                ld.anchored_datetime,
                ld.nor_tendered,
                ld.discharge_commenced,
                ld.discharge_completed,
                ld.pilot_pickup_time,
                ld.pilot_disembarked,
                h.id AS vcn_id,
                COALESCE(
                    h.vcn_doc_num,
                    h.via_number,
                    ld.vcn_doc_num,
                    ''
                ) AS vcn_doc_num,
                h.vessel_name,
                h.berth_name,
                h.cargo_type,
                h.operation_type,
                SUM(l.quantity) AS quantity,
                MIN(po.start_dt) AS first_parcel_start,
                MAX(po.end_dt) AS last_parcel_end,
                COALESCE((
                    SELECT SUM(
                        EXTRACT(EPOCH FROM (
                            vd.delay_end::timestamp
                            - vd.delay_start::timestamp
                        )) / 3600.0
                    )
                    FROM vcn_delays vd
                    JOIN port_delay_types pdt
                      ON LOWER(TRIM(pdt.name))
                       = LOWER(TRIM(vd.delay_name))
                    WHERE vd.vcn_id = h.id
                      AND vd.delay_start IS NOT NULL
                      AND vd.delay_end IS NOT NULL
                      AND pdt.type = 'Pre-Berthing Delays'
                      AND pdt.responsibility = 'Agent'
                ), 0.0) AS non_port_preberthing_delay_hrs
            FROM lueu_parcel_log l
            JOIN ldud_parcel_ops po
              ON po.id = l.parcel_op_id
            JOIN ldud_header ld
              ON ld.id = po.ldud_id
            JOIN vcn_header h
              ON h.id = ld.vcn_id
            WHERE COALESCE(l.is_deleted, FALSE) = FALSE
              AND COALESCE(ld.is_deleted, FALSE) = FALSE
              AND COALESCE(l.is_shortclose, FALSE) = FALSE
            GROUP BY
                ld.id,
                ld.cast_off_datetime,
                ld.alongside_datetime,
                ld.anchored_datetime,
                ld.nor_tendered,
                ld.discharge_commenced,
                ld.discharge_completed,
                ld.pilot_pickup_time,
                ld.pilot_disembarked,
                h.id,
                COALESCE(
                    h.vcn_doc_num,
                    h.via_number,
                    ld.vcn_doc_num,
                    ''
                ),
                h.vessel_name,
                h.berth_name,
                h.cargo_type,
                h.operation_type
        """)
        live_raw = cur.fetchall()
    finally:
        conn.close()

    # Load the same displayed delay rows used by Vessel Delay Report.
    from ..vessel_delay.vessel_delay import fetch_vessel_delay_data

    vd_data = fetch_vessel_delay_data('ALL', 'ALL')
    vd_rows = vd_data.get('rows', [])

    # Index Port delay durations by VCN and report month.
    port_hours_by_vessel_month = {}

    for delay_row in vd_rows:
        if (
            str(delay_row.get('delay_account') or '')
            .strip()
            .lower() != 'port'
        ):
            continue

        vcn_no = str(
            delay_row.get('vcn_no') or ''
        ).strip().upper()

        vessel_name = str(
            delay_row.get('vessel_name') or ''
        ).strip().upper()

        month_label = str(
            delay_row.get('month') or ''
        ).strip()

        if not month_label:
            continue

        delay_hours = _parse_delay_hours(
            delay_row.get('hours')
        )

        # Prefer the VCN identifier to avoid mixing vessels
        if vcn_no:
            port_hours_by_vessel_month.setdefault(
                (vcn_no, month_label), []
            ).append(delay_hours)

        if vessel_name:
            port_hours_by_vessel_month.setdefault(
                (vessel_name, month_label), []
            ).append(delay_hours)

    live_rows = []

    for r in live_raw:
        dt = (
            _parse_dt(r.get('cast_off_datetime'))
            or _parse_dt(r.get('alongside_datetime'))
        )

        if not dt:
            continue

        fy_start = dt.year if dt.month >= 4 else dt.year - 1
        fy = f"{fy_start}-{(fy_start + 1) % 100:02d}"

        idx = dt.month - 4 if dt.month >= 4 else dt.month + 8
        mn_label = (
            f"{MONTH_NAMES[idx]}-{str(dt.year)[-2:]}"
        )

        along = _parse_dt(r.get('alongside_datetime'))
        cast = _parse_dt(r.get('cast_off_datetime'))
        p_pick = _parse_dt(r.get('pilot_pickup_time'))
        p_dis = _parse_dt(r.get('pilot_disembarked'))

        first_parcel_start = _parse_dt(
            r.get('first_parcel_start')
        )
        last_parcel_end = _parse_dt(
            r.get('last_parcel_end')
        )

        # Working Time: first parcel start to last parcel end.
        wt_hrs = (
            (last_parcel_end - first_parcel_start).total_seconds()
            / 3600.0
            if (
                first_parcel_start
                and last_parcel_end
                and last_parcel_end > first_parcel_start
            )
            else 0.0
        )

        # Port A/C: use displayed Vessel Delay Report rows
        # for this vessel and this report month only.
        v_doc = str(
            r.get('vcn_doc_num') or ''
        ).strip().upper()

        v_name = str(
            r.get('vessel_name') or ''
        ).strip().upper()

        v_id = str(r.get('vcn_id') or '').strip()

        # Try the document number first.
        hrs_list = port_hours_by_vessel_month.get(
            (v_doc, mn_label)
        ) if v_doc else None

        # Try the VCN number format used in the delay report.
        if hrs_list is None and v_id:
            hrs_list = port_hours_by_vessel_month.get(
                (f"VCN-{v_id}", mn_label)
            )

        # Use vessel name only if no VCN match exists.
        if hrs_list is None and v_name:
            hrs_list = port_hours_by_vessel_month.get(
                (v_name, mn_label)
            )

        waiting_port_hrs = sum(hrs_list or [])

        # Keep the existing Non-Port calculation unchanged.
        waiting_non_port_hrs = float(
            r.get('non_port_preberthing_delay_hrs') or 0.0
        )

        pbw_hrs = waiting_port_hrs + waiting_non_port_hrs

        # Total Berth Stay: Cast-off minus Alongside.
        sab_hrs = (
            (cast - along).total_seconds() / 3600.0
            if cast and along and cast > along
            else 0.0
        )

        # Inward movement: Pilot Pickup to Alongside.
        inward_hrs = (
            (along - p_pick).total_seconds() / 3600.0
            if along and p_pick and along > p_pick
            else 0.0
        )

        # Outward movement: Cast-off to Pilot Disembarked.
        outward_hrs = (
            (p_dis - cast).total_seconds() / 3600.0
            if cast and p_dis and p_dis > cast
            else 0.0
        )

        live_rows.append({
            'fin_year': fy,
            'month': mn_label,
            'berth_no': r.get('berth_name'),
            'cargo': r.get('cargo_type'),
            'category': 'LIQUID',
            'new_cat': 'LIQUID',
            'category1': 'LIQUID',
            'unloading_terminal': r.get('berth_name'),
            'quantity': float(r.get('quantity') or 0),
            'pre_berthing_waiting': pbw_hrs / 24.0,
            'waiting_port': waiting_port_hrs / 24.0,
            'waiting_non_port': waiting_non_port_hrs / 24.0,
            'stay_at_berth': sab_hrs / 24.0,
            'working_time': wt_hrs / 24.0,
            'inward_movement': inward_hrs / 24.0,
            'outward_movement': outward_hrs / 24.0,
            'overseas_coastal': 'Overseas',
            'vcn_no': v_id,
            'ops_commenced': first_parcel_start,
            'cargo_completion': last_parcel_end,
            'alongside': along,
            'cast_off': cast,
            'pilot_pickup': p_pick,
            'pilot_disembarked': p_dis
        })

    return live_rows



def calculate_section_metrics(rlist):
    if not rlist:
        return {
            "vsl_sailed_total": 0,
            "vsl_berth_total": 0,
            "vsl_overseas": 0,
            "vsl_coastal": 0,
            "vsl_stream_total": 0,
            "traffic_total": 0.0,
            "traffic_berth": 0.0,
            "traffic_stream": 0.0,
            "pbw_total": 0.0,
            "pbw_port": 0.0,
            "pbw_non_port": 0.0,
            "working_time": 0.0,
            "nw_working_total": 0.0,
            "nw_working_port": 0.0,
            "nw_working_non_port": 0.0,
            "nw_non_working_total": 0.0,
            "nw_non_working_port": 0.0,
            "nw_non_working_non_port": 0.0,
            "nav_total": 0.0,
            "inward_movement": 0.0,
            "outward_movement": 0.0,
            "shifting_time": 0.0,
            "trt_total": 0.0,
            "trt_port": 0.0,
            "trt_non_port": 0.0
        }

    # 1. Vessels Sailed
    vsl_sailed_total = len(rlist)
    vsl_stream_total = sum(1 for r in rlist if 'ANCH' in str(r.get('berth_no') or '').upper() or 'STREAM' in str(r.get('berth_no') or '').upper())
    vsl_berth_total = vsl_sailed_total - vsl_stream_total
    vsl_coastal = sum(1 for r in rlist if str(r.get('overseas_coastal') or '').strip().lower() == 'coastal' and 'ANCH' not in str(r.get('berth_no') or '').upper())
    vsl_overseas = vsl_berth_total - vsl_coastal

    # 2. Traffic Handled (000 Tonnes)
    traffic_total = sum(float(r.get('quantity') or 0) for r in rlist) / 1000.0
    traffic_stream = sum(float(r.get('quantity') or 0) for r in rlist if 'ANCH' in str(r.get('berth_no') or '').upper() or 'STREAM' in str(r.get('berth_no') or '').upper()) / 1000.0
    traffic_berth = traffic_total - traffic_stream

    # 3. Pre Berthing Waiting Time (Hrs)
    # Excel logic:
    #   Total = Port a/c + Non-Port a/c
    pbw_port = sum(float(r.get('waiting_port') or 0) for r in rlist) * 24.0
    pbw_non_port = sum(float(r.get('waiting_non_port') or 0) for r in rlist) * 24.0
    pbw_total = round(pbw_port + pbw_non_port, 2)
    pbw_port = round(pbw_port, 2)
    pbw_non_port = round(pbw_non_port, 2)

    # 4. Working Time (in Days and Hrs)
    wt_days_sum = 0.0
    wt_hrs_sum = 0.0
    for r in rlist:
        ops_start = _parse_dt(r.get('ops_commenced'))
        ops_end = _parse_dt(r.get('cargo_completion'))
        if ops_start and ops_end and ops_end > ops_start:
            wt_hrs_sum += (ops_end - ops_start).total_seconds() / 3600.0
            wt_days_sum += (ops_end - ops_start).total_seconds() / 86400.0
        else:
            wt_hrs_sum += float(r.get('working_time') or 0) * 24.0
            wt_days_sum += float(r.get('working_time') or 0)

    # Excel logic: Working Time is reported in HOURS, not days.
    working_time = round(wt_hrs_sum, 2)

    # 5. Stay at berth in hours (using stay_at_berth column if present, else Cast Off - Alongside timestamp)
    sab_hrs_sum = 0.0
    for r in rlist:
        if r.get('stay_at_berth') is not None:
            sab_hrs_sum += float(r.get('stay_at_berth') or 0) * 24.0
        else:
            along = _parse_dt(r.get('alongside'))
            cast = _parse_dt(r.get('cast_off'))
            if along and cast and cast > along:
                sab_hrs_sum += (cast - along).total_seconds() / 3600.0

    # N.W Time at working berth (Hrs) = Stay at berth (in Hrs) - Working Time (in Hrs)
    nw_working_total = round(sab_hrs_sum - wt_hrs_sum, 2)
    if nw_working_total < 0:
        nw_working_total = 0.0
    nw_working_port = nw_working_total
    nw_working_non_port = 0.0

    # 6. Time at Non working berth
    nw_non_working_total = 0.0
    nw_non_working_port = 0.0
    nw_non_working_non_port = 0.0

    # 7. Navigation Time (Hrs)
    in_ts_sum = 0.0
    out_ts_sum = 0.0
    for r in rlist:
        along = _parse_dt(r.get('alongside'))
        cast = _parse_dt(r.get('cast_off'))
        p_pick = _parse_dt(r.get('pilot_pickup'))
        p_dis = _parse_dt(r.get('pilot_disembarked'))

        if p_pick and along and along > p_pick:
            in_ts_sum += (along - p_pick).total_seconds() / 3600.0
        else:
            in_ts_sum += float(r.get('inward_movement') or 0) * 24.0

        if p_dis and cast and p_dis > cast:
            out_ts_sum += (p_dis - cast).total_seconds() / 3600.0
        else:
            out_ts_sum += float(r.get('outward_movement') or 0) * 24.0

    inward_movement = round(in_ts_sum, 2)
    outward_movement = round(out_ts_sum, 2)
    nav_total = round(inward_movement + outward_movement, 2)

    # 8. Shifting Time
    shifting_time = 0.0

    # 9. Turn Round Time (Hrs)
    # Total = (3 + 4 + 5 + 6)
    # Port  = (3(a) + 4 + 5(a) + 6(a) + 7 + 8)
    # N.P   = (3(b) + 5(b) + 6(b))
    # Excel logic: Total Turn Round Time = 3 + 4 + 5 + 6
    trt_total = round(
        pbw_total
        + working_time
        + nw_working_total
        + nw_non_working_total,
        2
    )
    trt_port = round(pbw_port + working_time + nw_working_port + nw_non_working_port + nav_total + shifting_time, 2)
    trt_non_port = round(pbw_non_port + nw_working_non_port + nw_non_working_non_port, 2)

    return {
        "vsl_sailed_total": vsl_sailed_total,
        "vsl_berth_total": vsl_berth_total,
        "vsl_overseas": vsl_overseas,
        "vsl_coastal": vsl_coastal,
        "vsl_stream_total": vsl_stream_total,
        "traffic_total": round(traffic_total, 3),
        "traffic_berth": round(traffic_berth, 3),
        "traffic_stream": round(traffic_stream, 3),
        "pbw_total": pbw_total,
        "pbw_port": pbw_port,
        "pbw_non_port": pbw_non_port,
        "working_time": working_time,
        "nw_working_total": nw_working_total,
        "nw_working_port": nw_working_port,
        "nw_working_non_port": nw_working_non_port,
        "nw_non_working_total": nw_non_working_total,
        "nw_non_working_port": nw_non_working_port,
        "nw_non_working_non_port": nw_non_working_non_port,
        "nav_total": nav_total,
        "inward_movement": inward_movement,
        "outward_movement": outward_movement,
        "shifting_time": shifting_time,
        "trt_total": trt_total,
        "trt_port": trt_port,
        "trt_non_port": trt_non_port
    }


def build_report7_data(fin_year: str, selected_month: str = "Jun-26"):
    conn = get_db()
    try:
        cur = get_cursor(conn)
        cur.execute("""
            SELECT fin_year, month, berth_no, cargo, category, new_cat, category1, unloading_terminal,
                   quantity, pre_berthing_waiting, waiting_port, waiting_non_port,
                   stay_at_berth, working_time, non_working_port, inward_movement, outward_movement,
                   overseas_coastal, cast_off, vcn_no, ops_commenced, cargo_completion, alongside,
                   pilot_pickup, pilot_disembarked, nor
            FROM mis_vessel_master
            WHERE fin_year = %s
        """, [fin_year])
        mv_rows = cur.fetchall()

        cur.execute("SELECT DISTINCT fin_year FROM mis_vessel_master WHERE fin_year IS NOT NULL ORDER BY fin_year DESC")
        avail_years = [r['fin_year'] for r in cur.fetchall()]
    finally:
        conn.close()

    live_pipeline_rows = _load_live_pipeline_rows()
    live_months = set(
        str(r.get('month') or '').strip()
        for r in live_pipeline_rows
        if r.get('month') and r.get('fin_year') == fin_year
    )

    # Use live data for months where live pipeline data is available;
    # use mis_vessel_master only for historical months not present in live data
    historical_rows = [
        r for r in mv_rows
        if str(r.get('month') or '').strip() not in live_months
    ]
    current_live_rows = [
        r for r in live_pipeline_rows
        if r['fin_year'] == fin_year
    ]

    all_dataset_rows = historical_rows + current_live_rows

    live_fys = set(r['fin_year'] for r in live_pipeline_rows)
    for fy in live_fys:
        if fy not in avail_years:
            avail_years.append(fy)
    avail_years.sort(reverse=True)

    m_opts = month_options_for(fin_year)
    m_labels = [opt["label"] for opt in m_opts]

    # Filter by selected_month if specified and not 'ALL'
    if selected_month and selected_month != 'ALL':
        filtered_rows = [r for r in all_dataset_rows if str(r.get('month') or '').strip() == selected_month]
    else:
        filtered_rows = all_dataset_rows

    # Only include vessels where cast off is complete (exclude partial data)
    filtered_rows = [r for r in filtered_rows if r.get('cast_off')]

    def is_liquid(r):
        b = str(r.get('berth_no') or '').strip().upper()
        c = str(r.get('cargo') or '').strip().upper()
        cat = str(r.get('category') or '').strip().upper()
        ncat = str(r.get('new_cat') or '').strip().upper()
        term = str(r.get('unloading_terminal') or '').strip().upper()
        return ('LB-' in b) or ('LIQUID' in cat) or ('LIQUID' in ncat) or ('OIL' in c) or ('ACID' in c) or ('CHEMICAL' in cat)

    liquid_rows = [r for r in filtered_rows if is_liquid(r)]

    liquid_metrics = calculate_section_metrics(liquid_rows)
    all_metrics = calculate_section_metrics(filtered_rows)

    return {
        "fin_year": fin_year,
        "selected_month": selected_month,
        "available_years": avail_years if avail_years else [fin_year],
        "month_labels": m_labels,
        "liquid_bulk": liquid_metrics,
        "all_vessels": all_metrics
    }


def _get_current_month():
    now = datetime.now()
    idx = now.month - 4 if now.month >= 4 else now.month + 8
    return f"{MONTH_NAMES[idx]}-{str(now.year)[-2:]}"


def _get_current_fy():
    now = datetime.now()
    fy_start = now.year if now.month >= 4 else now.year - 1
    return f"{fy_start}-{(fy_start + 1) % 100:02d}"


@bp.route('/module/RP01/report7/')
@login_required
def report7_page():
    return render_template('report7/report7.html', username=session.get('username'))


@bp.route('/api/module/RP01/report7/data')
@login_required
def report7_data():
    fin_year = request.args.get('year', _get_current_fy()).strip()
    month = request.args.get('month', _get_current_month()).strip()
    try:
        data = build_report7_data(fin_year, month)
        return jsonify(data)
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@bp.route('/api/module/RP01/report7/export')
@login_required
def report7_export():
    fin_year = request.args.get('year', _get_current_fy()).strip()
    month = request.args.get('month', _get_current_month()).strip()
    try:
        data = build_report7_data(fin_year, month)

        wb = Workbook()
        ws = wb.active
        ws.title = "Physical Perf"
        ws.views.sheetView[0].showGridLines = True

        thin_border = Border(
            left=Side(style='thin', color='A0A0A0'),
            right=Side(style='thin', color='A0A0A0'),
            top=Side(style='thin', color='A0A0A0'),
            bottom=Side(style='thin', color='A0A0A0')
        )
        bold_font = Font(name='Calibri', size=11, bold=True)
        norm_font = Font(name='Calibri', size=10)

        title_font = Font(name='Calibri', size=14, bold=True)
        subhead_font = Font(name='Calibri', size=11, bold=True)

        center_align = Alignment(horizontal='center', vertical='center')
        left_align = Alignment(horizontal='left', vertical='center')
        right_align = Alignment(horizontal='right', vertical='center')

        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        # Row 1: Main Header
        ws.merge_cells("A1:C1")
        title_cell = ws["A1"]
        title_cell.value = "Performance in Terms of Vessels Handled -All Vessels"
        title_cell.font = title_font
        title_cell.alignment = left_align

        # Row 2: Sub Header
        ws.merge_cells("A2:C2")
        sub_cell = ws["A2"]
        sub_cell.value = "Port: JAWAHARLAL NEHRU PORT AUTHORITY"
        sub_cell.font = subhead_font
        sub_cell.alignment = left_align

        # Row 4: Column Headers
        ws.cell(row=4, column=1, value="Item").font = bold_font
        ws.cell(row=4, column=1).alignment = center_align
        ws.cell(row=4, column=1).border = thin_border

        ws.cell(row=4, column=2, value="Liquid\nBulk").font = bold_font
        ws.cell(row=4, column=2).alignment = center_align
        ws.cell(row=4, column=2).border = thin_border

        ws.cell(row=4, column=3, value="All").font = bold_font
        ws.cell(row=4, column=3).alignment = center_align
        ws.cell(row=4, column=3).border = thin_border

        rows_def = [
            ("1. No. of Vessels Sailed- (Total No.)", "vsl_sailed_total", "0", True, True),
            ("    a) At Berth- (No.)", "vsl_berth_total", "0", False, True),
            ("        (i) Overseas", "vsl_overseas", "0", False, True),
            ("        (ii) Coastal", "vsl_coastal", "0", False, True),
            ("    b) At Stream- (No.) (Inn Anchorage)", "vsl_stream_total", "0", False, False),
            ("2. Total Cargo Traffic Handled - (000 Tonnes)", "traffic_total", "#,##0.000", True, True),
            ("    a) At Berth - (000 Tonnes)", "traffic_berth", "#,##0.000", False, True),
            ("    b) At Stream (Inn Anchorage) - (000 Tonnes)", "traffic_stream", "#,##0.000", False, False),
            ("3. Pre Berthing Waiting Time(Hrs)-Total", "pbw_total", "#,##0.00", True, True),
            ("    a) Pre Berthing Waiting Time(Hrs)-Port a/c", "pbw_port", "#,##0.00", False, True),
            ("    b) Pre Berthing Waiting Time(Hrs)-Non Port a/c", "pbw_non_port", "#,##0.00", False, True),
            ("4. Working Time (Hrs)", "working_time", "#,##0.00", True, True),
            ("5. N.W Time at working berth(Hrs)-Total", "nw_working_total", "#,##0.00", True, True),
            ("    a) N.W Time at working berth(Hrs)-Port a/c", "nw_working_port", "#,##0.00", False, True),
            ("    b) N.W Time at working berth(Hrs)-N.P a/c", "nw_working_non_port", "#,##0.00", False, False),
            ("6. Time at Non working berth(Hrs)-Total", "nw_non_working_total", "#,##0.00", True, False),
            ("    a) Time at Non working berth(Hrs)-Port a/c", "nw_non_working_port", "#,##0.00", False, False),
            ("    b) Time at Non working berth(Hrs)-N.P a/c", "nw_non_working_non_port", "#,##0.00", False, False),
            ("7. Navigation Time (Hrs)-Total", "nav_total", "#,##0.00", True, True),
            ("    a) Inward Movement (Hrs)", "inward_movement", "#,##0.00", False, True),
            ("    b) Outward Movement (Hrs)", "outward_movement", "#,##0.00", False, True),
            ("8. Shifting Time", "shifting_time", "0", False, False),
            ("9. Turn Round Time (Hrs)-Total=(3+4+5+6)", "trt_total", "#,##0.00", True, False),
            ("    a) Turn Round Time (Hrs)-Port a/c=(3(a)+4+5(a)+6(a)+7+8)", "trt_port", "#,##0.00", False, False),
            ("    b) Turn Round Time (Hrs)-N.P a/c=(3(b)+5(b)+6(b))", "trt_non_port", "#,##0.00", False, False),
        ]

        curr_r = 5
        for label, key, num_fmt, is_bold, is_yellow in rows_def:
            item_cell = ws.cell(row=curr_r, column=1, value=label)
            item_cell.font = bold_font if is_bold else norm_font
            item_cell.alignment = left_align
            item_cell.border = thin_border
            if is_yellow:
                item_cell.fill = yellow_fill

            liq_val = data["liquid_bulk"].get(key, 0)
            liq_cell = ws.cell(row=curr_r, column=2, value=liq_val)
            liq_cell.font = bold_font if is_bold else norm_font
            liq_cell.alignment = right_align
            liq_cell.number_format = num_fmt
            liq_cell.border = thin_border
            if is_yellow:
                liq_cell.fill = yellow_fill

            all_val = data["all_vessels"].get(key, 0)
            all_cell = ws.cell(row=curr_r, column=3, value=all_val)
            all_cell.font = bold_font if is_bold else norm_font
            all_cell.alignment = right_align
            all_cell.number_format = num_fmt
            all_cell.border = thin_border

            curr_r += 1


        ws.column_dimensions["A"].width = 56
        ws.column_dimensions["B"].width = 16
        ws.column_dimensions["C"].width = 16

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        filename = f"7. Physical Performance- PDMP-BT {month}.xlsx"
        return send_file(
            buf,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500
