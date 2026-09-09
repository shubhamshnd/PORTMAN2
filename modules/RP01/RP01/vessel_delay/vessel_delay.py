"""
Vessel Delay Report - RP01
Displays operational vessel delay records completely dynamic from the database.
Columns: VCN No, vessel Name, Month, Delay name, Delay type, Delay account, Start time, End time, Hours
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
    for length in (19, 16, 10):
        sub = s[:length]
        for fmt in (
            "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
            "%d-%m-%Y %H:%M:%S", "%d-%m-%Y %H:%M", "%d-%m-%Y",
            "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y"
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

        # 1. Discover available financial years dynamically from vessel delay records
        cur.execute("""
            SELECT DISTINCT ref_date FROM (
                SELECT vd.delay_start AS ref_date FROM vcn_delays vd WHERE vd.delay_start IS NOT NULL AND TRIM(vd.delay_start) != ''
                UNION
                SELECT vh.doc_date AS ref_date FROM vcn_header vh WHERE vh.doc_date IS NOT NULL AND TRIM(vh.doc_date) != ''
            ) t
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
        # 2. DYNAMIC QUERY: vcn_delays joined with vcn_header and delay master tables
        # =========================================================================
        cur.execute("""
            SELECT
                vd.id,
                vd.vcn_id,
                COALESCE(vh.vcn_doc_num, vh.via_number, '') AS vcn_no,
                COALESCE(vh.vessel_name, '') AS vessel_name,
                vh.doc_date,
                vh.created_date,
                vd.delay_name,
                vd.delay_start,
                vd.delay_end,
                COALESCE(vdt.type, pdt.delay_type, pdt.type, '') AS delay_type,
                COALESCE(NULLIF(TRIM(pdt.responsibility), ''), 'Port') AS delay_account
            FROM vcn_delays vd
            LEFT JOIN vcn_header vh ON vd.vcn_id = vh.id
            LEFT JOIN vessel_delay_types vdt ON LOWER(TRIM(vd.delay_name)) = LOWER(TRIM(vdt.name))
            LEFT JOIN port_delay_types pdt ON LOWER(TRIM(vd.delay_name)) = LOWER(TRIM(pdt.name))
            WHERE (vd.delay_name IS NOT NULL AND TRIM(vd.delay_name) != '')
               OR (vd.delay_start IS NOT NULL AND TRIM(vd.delay_start) != '')
            ORDER BY vd.id ASC
        """)
        for r in cur.fetchall():
            d_start = _parse_datetime(r['delay_start'])
            d_end = _parse_datetime(r['delay_end'])
            ref_dt = d_start or _parse_datetime(r['doc_date']) or _parse_datetime(r['created_date'])

            if not ref_dt:
                continue

            fy, m_idx = _date_to_fin_year_and_idx(ref_dt)
            m_str = ref_dt.strftime("%b-%y")

            hrs_val = None
            if d_start and d_end:
                diff_sec = (d_end - d_start).total_seconds()
                hrs_val = round(max(diff_sec / 3600.0, 0.0), 2)

            raw_rows.append({
                "sort_dt": ref_dt,
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
                "hours": f"{hrs_val:.2f}" if hrs_val is not None else ""
            })

        # =========================================================================
        # 3. FILTERING: Dynamic filtering by selected Year and Month
        # =========================================================================
        filtered_rows = []
        for row in raw_rows:
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

        # Sort chronologically by sort_dt
        filtered_rows.sort(key=lambda x: (x["sort_dt"] or datetime.min, x["vcn_no"]))

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

        # Columns layout matching user mockup (Image 2)
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
                        try:
                            cell.value = float(val)
                            cell.number_format = "0.00"
                        except ValueError:
                            cell.value = val
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
