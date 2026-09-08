"""
Vessel Timing Details Report - RP01
Displays operational timing milestones for vessels based on ldud_header,
ldud_parcel_ops, and lueu_parcel_log.
"""

import io
import traceback
from datetime import datetime, date
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


def _format_datetime(val):
    if not val:
        return ""
    if isinstance(val, (datetime, date)):
        return val.strftime("%d-%m-%Y %H:%M")
    s = str(val).strip().replace('T', ' ')
    for length in (19, 16, 10):
        sub = s[:length]
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d", "%d-%m-%Y %H:%M", "%d-%m-%Y"):
            try:
                dt = datetime.strptime(sub, fmt)
                if length == 10:
                    return dt.strftime("%d-%m-%Y")
                return dt.strftime("%d-%m-%Y %H:%M")
            except Exception:
                pass
    return s


def _parse_filter_date(val):
    if not val:
        return None
    if isinstance(val, (datetime, date)):
        return val
    s = str(val).strip().replace('T', ' ')
    for length in (19, 16, 10):
        sub = s[:length]
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d", "%d-%m-%Y %H:%M", "%d-%m-%Y"):
            try:
                return datetime.strptime(sub, fmt)
            except Exception:
                pass
    return None


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


def fetch_vessel_timing_data(year_filter=None, month_filter=None):
    curr_fy, curr_m_idx = _current_fin_year_and_idx()

    if year_filter is None:
        year_filter = curr_fy
    if month_filter is None:
        month_filter = str(curr_m_idx)

    conn = get_db()
    try:
        cur = get_cursor(conn)

        # Available financial years from ldud_header cast_off_datetime
        cur.execute("""
            SELECT DISTINCT cast_off_datetime AS ref_date
            FROM ldud_header
            WHERE is_deleted IS NOT TRUE
              AND cast_off_datetime IS NOT NULL
              AND NULLIF(TRIM(cast_off_datetime), '') IS NOT NULL
        """)
        fys = set()
        recorded_months_by_fy = {}
        for r in cur.fetchall():
            d = _parse_filter_date(r['ref_date'])
            if d:
                fy, m_idx = _date_to_fin_year_and_idx(d)
                fys.add(fy)
                recorded_months_by_fy.setdefault(fy, set()).add(m_idx)

        fys.add(curr_fy)
        avail_years = sorted(list(fys), reverse=True)

        # Main query: resolve each parcel's timing and quantity (with short close deducted),
        # then aggregate across all parcels for the vessel to get first parcel start to last parcel end
        query = """
            WITH parcel_logs_per_op AS (
                SELECT 
                    l.parcel_op_id,
                    MIN(l.entry_date || ' ' || l.from_time) AS log_start,
                    MAX(CASE 
                        WHEN COALESCE(l.is_shortclose, FALSE) = FALSE 
                         AND LOWER(COALESCE(l.remarks, '')) NOT LIKE '%short%' 
                        THEN l.entry_date || ' ' || l.to_time 
                    END) AS log_end,
                    SUM(CASE 
                        WHEN COALESCE(l.is_shortclose, FALSE) = FALSE 
                         AND LOWER(COALESCE(l.remarks, '')) NOT LIKE '%short%' 
                        THEN l.quantity 
                        ELSE 0 
                    END) AS log_actual_qty,
                    SUM(CASE 
                        WHEN COALESCE(l.is_shortclose, FALSE) = TRUE 
                          OR LOWER(COALESCE(l.remarks, '')) LIKE '%short%' 
                        THEN l.quantity 
                        ELSE 0 
                    END) AS log_short_qty
                FROM lueu_parcel_log l
                WHERE l.is_deleted IS NOT TRUE
                GROUP BY l.parcel_op_id
            ),
            resolved_parcels AS (
                SELECT 
                    po.id,
                    po.ldud_id,
                    po.quantity AS po_qty,
                    COALESCE(pl.log_short_qty, 0) AS short_qty,
                    COALESCE(NULLIF(REPLACE(po.start_dt, 'T', ' '), ''), pl.log_start) AS parcel_start,
                    COALESCE(NULLIF(REPLACE(po.end_dt, 'T', ' '), ''), pl.log_end) AS parcel_end,
                    pl.log_actual_qty
                FROM ldud_parcel_ops po
                LEFT JOIN parcel_logs_per_op pl ON pl.parcel_op_id = po.id
            ),
            vessel_parcels_agg AS (
                SELECT 
                    rp.ldud_id,
                    MIN(rp.parcel_start) AS first_parcel_start,
                    MAX(rp.parcel_end) AS last_parcel_end,
                    SUM(rp.po_qty) AS total_po_qty,
                    SUM(rp.short_qty) AS total_short_qty,
                    SUM(rp.log_actual_qty) AS total_log_actual_qty
                FROM resolved_parcels rp
                GROUP BY rp.ldud_id
            )
            SELECT 
                ld.id,
                ld.vessel_name,
                CASE 
                    WHEN vpa.total_po_qty IS NOT NULL 
                        THEN GREATEST(0, vpa.total_po_qty - COALESCE(vpa.total_short_qty, 0))
                    WHEN vpa.total_log_actual_qty IS NOT NULL 
                        THEN vpa.total_log_actual_qty
                    ELSE GREATEST(0, COALESCE(ld.initial_draft_survey_quantity, 0) - COALESCE(vpa.total_short_qty, 0))
                END AS quantity,
                COALESCE(ld.arrival_inner_anchorage, ld.anchored_datetime, ld.arrival_outer_anchorage) AS arrival,
                ld.nor_tendered,
                ld.nor_accepted,
                ld.pilot_pickup_time,
                ld.first_line,
                ld.alongside_datetime,
                ld.custom_clearance,
                ld.agent_stevedore_onboard,
                COALESCE(vpa.first_parcel_start, ld.discharge_commenced) AS cargo_commenced,
                COALESCE(vpa.last_parcel_end, ld.discharge_completed) AS cargo_completed,
                ld.pilot_board_departure,
                ld.cast_off_datetime,
                ld.pilot_disembarked,
                ld.cast_off_datetime AS ref_date
            FROM ldud_header ld
            LEFT JOIN vessel_parcels_agg vpa ON vpa.ldud_id = ld.id
            WHERE ld.is_deleted IS NOT TRUE
              AND ld.cast_off_datetime IS NOT NULL
              AND NULLIF(TRIM(ld.cast_off_datetime), '') IS NOT NULL
            ORDER BY ld.cast_off_datetime ASC, ld.id ASC
        """
        cur.execute(query)
        raw_rows = cur.fetchall()
    finally:
        conn.close()

    filtered_rows = []
    for r in raw_rows:
        ref_dt = _parse_filter_date(r['ref_date'])
        row_fy, row_m_idx = _date_to_fin_year_and_idx(ref_dt) if ref_dt else (None, None)

        # Apply Year filter
        if year_filter and year_filter != "ALL":
            if row_fy != year_filter and (not ref_dt or str(ref_dt.year) != str(year_filter)):
                continue

        # Apply Month filter
        if month_filter and month_filter != "ALL":
            if row_m_idx is None:
                continue
            if str(month_filter).isdigit():
                if row_m_idx != int(month_filter):
                    continue
            else:
                m_name = MONTH_LABELS[row_m_idx] if row_m_idx is not None else ""
                if m_name.lower() != str(month_filter).strip().lower():
                    continue

        filtered_rows.append({
            "id": r["id"],
            "vessel_name": r["vessel_name"] or "",
            "quantity": float(r["quantity"]) if r["quantity"] is not None else 0.0,
            "arrival": _format_datetime(r["arrival"]),
            "nor": _format_datetime(r["nor_tendered"]),
            "nor_accepted": _format_datetime(r["nor_accepted"]),
            "pilot_pickup": _format_datetime(r["pilot_pickup_time"]),
            "first_line": _format_datetime(r["first_line"]),
            "alongside": _format_datetime(r["alongside_datetime"]),
            "customs_clearance": _format_datetime(r["custom_clearance"]),
            "agent_clearance": _format_datetime(r["agent_stevedore_onboard"]),
            "cargo_commenced": _format_datetime(r["cargo_commenced"]),
            "cargo_completed": _format_datetime(r["cargo_completed"]),
            "pilot_board_departure": _format_datetime(r["pilot_board_departure"]),
            "cast_off": _format_datetime(r["cast_off_datetime"]),
            "pilot_disembarked": _format_datetime(r["pilot_disembarked"]),
        })

    # Available months list for dropdown
    month_options = [{"idx": i, "label": MONTH_LABELS[i]} for i in range(12)]

    return {
        "available_years": avail_years,
        "months": month_options,
        "current_fin_year": curr_fy,
        "current_month_idx": curr_m_idx,
        "current_month_label": MONTH_LABELS[curr_m_idx],
        "selected_year": year_filter,
        "selected_month": month_filter,
        "rows": filtered_rows,
        "total_count": len(filtered_rows)
    }


# ================== ROUTES ==================

@bp.route('/module/RP01/vessel-timing/')
@login_required
def vessel_timing_page():
    return render_template('vessel_timing/vessel_timing.html', username=session.get('username'))


@bp.route('/api/module/RP01/vessel-timing/data')
@login_required
def vessel_timing_data():
    curr_fy, curr_m_idx = _current_fin_year_and_idx()
    year = request.args.get('year', curr_fy).strip()
    month = request.args.get('month', str(curr_m_idx)).strip()
    try:
        data = fetch_vessel_timing_data(year, month)
        return jsonify(data)
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@bp.route('/api/module/RP01/vessel-timing/export')
@login_required
def vessel_timing_export():
    curr_fy, curr_m_idx = _current_fin_year_and_idx()
    year = request.args.get('year', curr_fy).strip()
    month = request.args.get('month', str(curr_m_idx)).strip()
    try:
        data = fetch_vessel_timing_data(year, month)
        rows = data["rows"]

        wb = Workbook()
        ws = wb.active
        ws.title = "Vessel Timing Details"
        ws.views.sheetView[0].showGridLines = True

        thin_side = Side(style='thin', color='444444')
        cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        title_font = Font(name='Calibri', size=11, bold=True)
        header_font = Font(name='Calibri', size=10, bold=True)
        data_font = Font(name='Calibri', size=10)

        center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
        left_align = Alignment(horizontal='left', vertical='center')
        right_align = Alignment(horizontal='right', vertical='center')

        # Row 1: Merged Title Block matching Image 1
        ws.merge_cells("A1:O1")
        title_cell = ws["A1"]
        title_cell.value = "Vessel timing details"
        title_cell.font = title_font
        title_cell.alignment = center_align
        title_cell.fill = header_fill
        for col_idx in range(1, 16):
            ws.cell(row=1, column=col_idx).border = cell_border
            ws.cell(row=1, column=col_idx).fill = header_fill

        # Row 2: 15 Column Headers
        headers = [
            ("Vessel Name", "vessel_name", left_align),
            ("Qty", "quantity", right_align),
            ("Arrival", "arrival", center_align),
            ("NOR", "nor", center_align),
            ("NOR Accepted", "nor_accepted", center_align),
            ("Pilot pickup", "pilot_pickup", center_align),
            ("First line", "first_line", center_align),
            ("Along side", "alongside", center_align),
            ("Customs clearance", "customs_clearance", center_align),
            ("Agent clearance", "agent_clearance", center_align),
            ("Cargo commence time", "cargo_commenced", center_align),
            ("Cargo complete time", "cargo_completed", center_align),
            ("Pilot board departure", "pilot_board_departure", center_align),
            ("Cast off time", "cast_off", center_align),
            ("Pilot disembarked", "pilot_disembarked", center_align),
        ]

        ws.row_dimensions[1].height = 24
        ws.row_dimensions[2].height = 28

        for col_idx, (title, _, _) in enumerate(headers, start=1):
            cell = ws.cell(row=2, column=col_idx, value=title)
            cell.font = header_font
            cell.alignment = center_align
            cell.border = cell_border
            cell.fill = header_fill

        # Data Rows
        curr_row = 3
        for r in rows:
            ws.row_dimensions[curr_row].height = 20
            for col_idx, (_, key, align) in enumerate(headers, start=1):
                val = r.get(key, "")
                cell = ws.cell(row=curr_row, column=col_idx)
                cell.font = data_font
                cell.alignment = align
                cell.border = cell_border

                if key == "quantity":
                    cell.value = float(val) if val else 0.0
                    cell.number_format = "#,##0.00"
                else:
                    cell.value = val or ""
            curr_row += 1

        # Auto-adjust column widths
        for col_idx in range(1, 16):
            col_letter = get_column_letter(col_idx)
            max_len = max(len(str(ws.cell(row=r, column=col_idx).value or '')) for r in range(2, max(curr_row, 3)))
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
        ws.column_dimensions["A"].width = 22  # Vessel Name wider

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        month_label = MONTH_LABELS[int(month)] if str(month).isdigit() and int(month) < len(MONTH_LABELS) else month
        filename = f"Vessel_Timing_Details_{year}_{month_label}.xlsx"
        return send_file(
            buf,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500
