"""
Vessel Delay Report - RP01
Displays operational vessel delay records based on Vessel Cast Off Date Time.
Columns: VCN No, vessel Name, Month, Cast Off Time, Delay name, Delay type, Delay account, Start time, End time, Hours
"""

import io
import traceback
from datetime import datetime, date
from decimal import Decimal
from functools import wraps

from flask import jsonify, request, render_template, send_file, session, redirect, url_for
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

from database import get_db, get_cursor
from .. import bp


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated


MONTH_LABELS = [
    "April", "May", "June", "July", "August", "September",
    "October", "November", "December", "January", "February", "March"
]


def _parse_datetime(val):
    if not val:
        return None
    if isinstance(val, datetime):
        return val
    if isinstance(val, date):
        return datetime(val.year, val.month, val.day)
    s = str(val).strip().replace('T', ' ')
    for length in (19, 16, 10, 8):
        sub = s[:length]
        for fmt in (
            "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
            "%d-%m-%Y %H:%M:%S", "%d-%m-%Y %H:%M", "%d-%m-%Y",
            "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y",
            "%d%m%Y %H:%M:%S", "%d%m%Y %H:%M", "%d%m%Y"
        ):
            try:
                return datetime.strptime(sub, fmt)
            except Exception:
                pass
    return None


def _format_datetime(dt):
    if not dt:
        return ""
    if isinstance(dt, datetime):
        return dt.strftime("%d-%m-%Y %H:%M")
    parsed = _parse_datetime(dt)
    return parsed.strftime("%d-%m-%Y %H:%M") if parsed else str(dt)


def _format_month(dt):
    if not dt:
        return ""
    if isinstance(dt, datetime):
        return dt.strftime("%b-%y")
    parsed = _parse_datetime(dt)
    return parsed.strftime("%b-%y") if parsed else ""


def _format_duration_hh_mm(start_or_sec, end_dt=None):
    if start_or_sec is None:
        return ""
    if end_dt is not None:
        if not start_or_sec:
            return ""
        diff_sec = (end_dt - start_or_sec).total_seconds()
        if diff_sec < 0:
            return ""
    else:
        try:
            diff_sec = float(start_or_sec)
            if diff_sec < 0:
                return ""
        except (ValueError, TypeError):
            return ""
    total_minutes = int(round(diff_sec / 60.0))
    hours = total_minutes // 60
    mins = total_minutes % 60
    return f"{hours:02d}:{mins:02d}"


def _date_to_fin_year_and_idx(dt):
    if dt.month >= 4:
        fy = f"{dt.year}-{str(dt.year + 1)[-2:]}"
        idx = dt.month - 4
    else:
        fy = f"{dt.year - 1}-{str(dt.year)[-2:]}"
        idx = dt.month + 8
    return fy, idx


def _current_fin_year_and_idx():
    return _date_to_fin_year_and_idx(datetime.now())


def fetch_vessel_delay_data(year_filter=None, month_filter=None):
    curr_fy, curr_m_idx = _current_fin_year_and_idx()

    if year_filter is None:
        year_filter = curr_fy
    if month_filter is None:
        month_filter = str(curr_m_idx)

    conn = get_db()
    try:
        cur = get_cursor(conn)

        # 1. Discover available financial years dynamically strictly from cast_off_datetime
        cur.execute("""
            SELECT DISTINCT lh.cast_off_datetime AS ref_date
            FROM ldud_header lh
            LEFT JOIN vcn_header vh ON (lh.vcn_id = vh.id OR lh.vcn_doc_num = vh.vcn_doc_num)
            LEFT JOIN vcn_delays vd ON vd.vcn_id = vh.id
            WHERE lh.cast_off_datetime IS NOT NULL
              AND TRIM(lh.cast_off_datetime) != ''
              AND lh.is_deleted IS NOT TRUE
              AND (
                  (vd.delay_name IS NOT NULL AND TRIM(vd.delay_name) != '')
                  OR (vd.delay_start IS NOT NULL AND TRIM(vd.delay_start) != '')
                  OR ((COALESCE(NULLIF(TRIM(lh.anchored_datetime), ''), NULLIF(TRIM(lh.nor_tendered), '')) IS NOT NULL)
                      AND (NULLIF(TRIM(lh.pilot_pickup_time), '') IS NOT NULL))
                  OR ((NULLIF(TRIM(lh.pilot_pickup_time), '') IS NOT NULL)
                      AND (NULLIF(TRIM(lh.alongside_datetime), '') IS NOT NULL))
                  OR ((NULLIF(TRIM(lh.cast_off_datetime), '') IS NOT NULL)
                      AND (NULLIF(TRIM(lh.pilot_disembarked), '') IS NOT NULL))
              )
        """)
        fys = set()
        fys.add(curr_fy)
        for r in cur.fetchall():
            d = _parse_datetime(r['ref_date'])
            if d:
                fy, _ = _date_to_fin_year_and_idx(d)
                fys.add(fy)

        avail_years = sorted(list(fys), reverse=True)

        raw_rows = []

        # =========================================================================
        # 2a. ORIGINAL DELAYS from vcn_delays (shown first)
        # =========================================================================
        cur.execute("""
            SELECT
                vd.id,
                vd.vcn_id,
                COALESCE(vh.vcn_doc_num, vh.via_number, lh.vcn_doc_num, '') AS vcn_no,
                COALESCE(vh.vessel_name, lh.vessel_name, '') AS vessel_name,
                lh.cast_off_datetime,
                vd.delay_name,
                vd.delay_start,
                vd.delay_end,
                COALESCE(vdt.type, pdt.delay_type, pdt.type, '') AS delay_type,
                COALESCE(NULLIF(TRIM(pdt.responsibility), ''), 'Port') AS delay_account
            FROM vcn_delays vd
            JOIN vcn_header vh ON vd.vcn_id = vh.id
            JOIN ldud_header lh ON (lh.vcn_id = vh.id OR lh.vcn_doc_num = vh.vcn_doc_num)
                 AND lh.is_deleted IS NOT TRUE
                 AND lh.cast_off_datetime IS NOT NULL
                 AND TRIM(lh.cast_off_datetime) != ''
            LEFT JOIN vessel_delay_types vdt ON LOWER(TRIM(vd.delay_name)) = LOWER(TRIM(vdt.name))
            LEFT JOIN port_delay_types pdt ON LOWER(TRIM(vd.delay_name)) = LOWER(TRIM(pdt.name))
            WHERE ((vd.delay_name IS NOT NULL AND TRIM(vd.delay_name) != '')
               OR (vd.delay_start IS NOT NULL AND TRIM(vd.delay_start) != ''))
            ORDER BY vd.id ASC
        """)
        for r in cur.fetchall():
            d_cast_off = _parse_datetime(r['cast_off_datetime'])
            if not d_cast_off:
                continue

            d_start = _parse_datetime(r['delay_start'])
            d_end = _parse_datetime(r['delay_end'])

            # Report basis: Strictly Vessel Cast Off Date Time
            ref_dt = d_cast_off

            if not ref_dt:
                continue

            fy, m_idx = _date_to_fin_year_and_idx(ref_dt)
            m_str = ref_dt.strftime("%b-%y")

            hrs_val = ""
            if d_start and d_end:
                hrs_val = _format_duration_hh_mm(d_start, d_end)

            raw_rows.append({
                "sort_dt": ref_dt,
                "delay_dt": d_start or ref_dt,
                "source_priority": 1,
                "fin_year": fy,
                "month_idx": m_idx,
                "vcn_no": r['vcn_no'] or (f"VCN-{r['vcn_id']}" if r['vcn_id'] else ""),
                "vessel_name": r['vessel_name'],
                "month": m_str,
                "delay_name": r['delay_name'] or "",
                "delay_type": (r['delay_type'] or '').strip(),
                "delay_account": (r['delay_account'] or 'Port').strip(),
                "start_time": _format_datetime(d_start),
                "end_time": _format_datetime(d_end),
                "hours": hrs_val
            })

        # =========================================================================
        # 2b. MILESTONE DELAYS from ldud_header (shown after original delays)
        #     - Berth Not availeb: Anchorage/NOR -> Pilot Pickup
        #     - Pilot pick up- along side: Pilot Pickup -> Alongside
        #     - cast of time - Pilot disemberd: Cast Off -> Pilot Disembarked
        # =========================================================================
        cur.execute("""
            SELECT
                lh.id AS ldud_id,
                lh.vcn_id,
                COALESCE(vh.vcn_doc_num, vh.via_number, lh.vcn_doc_num, '') AS vcn_no,
                COALESCE(vh.vessel_name, lh.vessel_name, '') AS vessel_name,
                lh.cast_off_datetime,
                lh.anchored_datetime,
                lh.nor_tendered,
                lh.pilot_pickup_time,
                lh.alongside_datetime,
                lh.pilot_disembarked
            FROM ldud_header lh
            LEFT JOIN vcn_header vh ON (lh.vcn_id = vh.id OR lh.vcn_doc_num = vh.vcn_doc_num)
            WHERE lh.is_deleted IS NOT TRUE
              AND lh.cast_off_datetime IS NOT NULL
              AND TRIM(lh.cast_off_datetime) != ''
            ORDER BY lh.id ASC
        """)
        for r in cur.fetchall():
            d_cast_off = _parse_datetime(r['cast_off_datetime'])
            if not d_cast_off:
                continue

            ref_dt = d_cast_off
            fy, m_idx = _date_to_fin_year_and_idx(ref_dt)
            m_str = ref_dt.strftime("%b-%y")
            v_vcn = r['vcn_no'] or (f"VCN-{r['vcn_id']}" if r['vcn_id'] else "")
            v_vessel = r['vessel_name'] or ""

            d_anchored = _parse_datetime(r['anchored_datetime']) or _parse_datetime(r['nor_tendered'])
            d_pilot = _parse_datetime(r['pilot_pickup_time'])
            d_alongside = _parse_datetime(r['alongside_datetime'])
            d_pilot_dis = _parse_datetime(r['pilot_disembarked'])

            # 1) Berth Not availeb (Anchorage / NOR -> Pilot Pickup)
            if d_anchored and d_pilot:
                diff_sec = (d_pilot - d_anchored).total_seconds()
                if diff_sec > 0:
                    raw_rows.append({
                        "sort_dt": ref_dt,
                        "delay_dt": d_anchored,
                        "source_priority": 2,
                        "fin_year": fy,
                        "month_idx": m_idx,
                        "vcn_no": v_vcn,
                        "vessel_name": v_vessel,
                        "month": m_str,
                        "delay_name": "Berth Not availeb",
                        "delay_type": "Port",
                        "delay_account": "Port",
                        "start_time": _format_datetime(d_anchored),
                        "end_time": _format_datetime(d_pilot),
                        "hours": _format_duration_hh_mm(diff_sec)
                    })

            # 2) Pilot pick up- along side (Pilot Pickup -> Alongside)
            if d_pilot and d_alongside:
                diff_sec = (d_alongside - d_pilot).total_seconds()
                if diff_sec > 0:
                    raw_rows.append({
                        "sort_dt": ref_dt,
                        "delay_dt": d_pilot,
                        "source_priority": 3,
                        "fin_year": fy,
                        "month_idx": m_idx,
                        "vcn_no": v_vcn,
                        "vessel_name": v_vessel,
                        "month": m_str,
                        "delay_name": "Pilot pick up- along side",
                        "delay_type": "Pilot",
                        "delay_account": "Port",
                        "start_time": _format_datetime(d_pilot),
                        "end_time": _format_datetime(d_alongside),
                        "hours": _format_duration_hh_mm(diff_sec)
                    })

            # 3) cast of time - Pilot disemberd (Cast Off -> Pilot Disembarked)
            if d_cast_off and d_pilot_dis:
                diff_sec = (d_pilot_dis - d_cast_off).total_seconds()
                if diff_sec > 0:
                    raw_rows.append({
                        "sort_dt": ref_dt,
                        "delay_dt": d_cast_off,
                        "source_priority": 4,
                        "fin_year": fy,
                        "month_idx": m_idx,
                        "vcn_no": v_vcn,
                        "vessel_name": v_vessel,
                        "month": m_str,
                        "delay_name": "cast of time - Pilot disemberd",
                        "delay_type": "Pilot",
                        "delay_account": "Port",
                        "start_time": _format_datetime(d_cast_off),
                        "end_time": _format_datetime(d_pilot_dis),
                        "hours": _format_duration_hh_mm(diff_sec)
                    })

        # Deduplicate identical records if any
        seen_entries = set()
        deduped_rows = []
        for row in raw_rows:
            entry_key = (row["vcn_no"], row["start_time"], row["end_time"], row["delay_name"])
            if entry_key in seen_entries:
                continue
            seen_entries.add(entry_key)
            deduped_rows.append(row)

        # =========================================================================
        # 3. FILTERING: Dynamic filtering by selected Year and Month (based on Cast Off)
        # =========================================================================
        filtered_rows = []
        for row in deduped_rows:
            # Year filter
            if year_filter and year_filter != "ALL":
                if row["fin_year"] != year_filter:
                    continue

            # Month filter
            if month_filter and month_filter != "ALL":
                try:
                    if int(row["month_idx"]) != int(month_filter):
                        continue
                except (ValueError, TypeError):
                    pass

            filtered_rows.append(row)

        # Sort: Original delays first (source_priority=1), followed by milestone delays (2, 3, 4)
        filtered_rows.sort(
            key=lambda x: (
                x["sort_dt"] or datetime.min,
                x["vcn_no"],
                x.get("source_priority", 1),
                x.get("delay_dt") or datetime.min
            )
        )

        clean_rows = []
        for idx, r in enumerate(filtered_rows, start=1):
            clean_rows.append({
                "sr_no": idx,
                "vcn_no": r["vcn_no"],
                "vessel_name": r["vessel_name"],
                "month": r["month"],
                "delay_name": r["delay_name"],
                "delay_type": r["delay_type"],
                "delay_account": r["delay_account"],
                "start_time": r["start_time"],
                "end_time": r["end_time"],
                "hours": r["hours"]
            })

        month_options = [{"idx": "ALL", "label": "ALL"}] + [
            {"idx": i, "label": MONTH_LABELS[i]} for i in range(12)
        ]

        return {
            "available_years": avail_years,
            "months": month_options,
            "current_fin_year": curr_fy,
            "current_month_idx": curr_m_idx,
            "current_month_label": MONTH_LABELS[curr_m_idx],
            "selected_year": year_filter,
            "selected_month": month_filter,
            "rows": clean_rows,
            "total_count": len(clean_rows)
        }

    finally:
        conn.close()


# ================== ROUTES ==================

@bp.route('/module/RP01/vessel-delay/')
@login_required
def vessel_delay_page():
    return render_template('vessel_delay/vessel_delay.html', username=session.get('username'))


@bp.route('/api/module/RP01/vessel-delay/data')
@login_required
def vessel_delay_data():
    curr_fy, curr_m_idx = _current_fin_year_and_idx()
    year = request.args.get('year', curr_fy).strip()
    month = request.args.get('month', str(curr_m_idx)).strip()
    try:
        data = fetch_vessel_delay_data(year, month)
        return jsonify(data)
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@bp.route('/api/module/RP01/vessel-delay/export')
@login_required
def vessel_delay_export():
    curr_fy, curr_m_idx = _current_fin_year_and_idx()
    year = request.args.get('year', curr_fy).strip()
    month = request.args.get('month', str(curr_m_idx)).strip()

    try:
        data = fetch_vessel_delay_data(year, month)
        rows = data.get("rows", [])

        wb = Workbook()
        ws = wb.active
        ws.title = "Vessel Delay Report"
        ws.views.sheetView[0].showGridLines = True

        # Styles
        font_title = Font(name="Calibri", size=13, bold=True, color="000000")
        header_font = Font(name="Calibri", size=10, bold=True, color="000000")
        data_font = Font(name="Calibri", size=10, color="000000")

        header_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

        thin_side = Side(border_style="thin", color="000000")
        cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        center_align = Alignment(horizontal="center", vertical="center")
        left_align = Alignment(horizontal="left", vertical="center")
        right_align = Alignment(horizontal="right", vertical="center")

        # Row 1: Title "Vessel delay report" centered
        ws.row_dimensions[1].height = 24
        ws.merge_cells("A1:I1")
        title_cell = ws.cell(row=1, column=1, value="Vessel delay report")
        title_cell.font = font_title
        title_cell.alignment = center_align

        # Columns layout
        headers = [
            ("VCN No", "vcn_no", center_align),
            ("vessel Name", "vessel_name", left_align),
            ("Month", "month", center_align),
            ("Delay name", "delay_name", left_align),
            ("Delay type", "delay_type", center_align),
            ("Delay account", "delay_account", center_align),
            ("Start time", "start_time", center_align),
            ("End time", "end_time", center_align),
            ("Hours", "hours", right_align),
        ]

        ws.row_dimensions[2].height = 24

        # Header Row (Row 2)
        for col_idx, (col_title, _, align) in enumerate(headers, start=1):
            cell = ws.cell(row=2, column=col_idx, value=col_title)
            cell.font = header_font
            cell.alignment = align
            cell.border = cell_border
            cell.fill = header_fill

        # Data Rows (Row 3+)
        curr_row = 3
        for r in rows:
            ws.row_dimensions[curr_row].height = 20
            for col_idx, (_, key, align) in enumerate(headers, start=1):
                val = r.get(key, "")
                cell = ws.cell(row=curr_row, column=col_idx)
                cell.font = data_font
                cell.alignment = align
                cell.border = cell_border

                if key == "hours":
                    if val != "" and val is not None:
                        cell.value = str(val)
                        cell.number_format = "@"
                    else:
                        cell.value = ""
                else:
                    cell.value = val or ""
            curr_row += 1

        # Column widths
        ws.column_dimensions["A"].width = 14  # VCN No
        ws.column_dimensions["B"].width = 24  # vessel Name
        ws.column_dimensions["C"].width = 12  # Month
        ws.column_dimensions["D"].width = 30  # Delay name
        ws.column_dimensions["E"].width = 15  # Delay type
        ws.column_dimensions["F"].width = 15  # Delay account
        ws.column_dimensions["G"].width = 20  # Start time
        ws.column_dimensions["H"].width = 20  # End time
        ws.column_dimensions["I"].width = 12  # Hours

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        month_label = MONTH_LABELS[int(month)] if str(month).isdigit() and int(month) < len(MONTH_LABELS) else month
        filename = f"Vessel_Delay_Report_{year}_{month_label}.xlsx"

        return send_file(
            buf,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500
