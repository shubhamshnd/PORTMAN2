"""Pro-forma invoice PDF — JJLTPL stationery, laid out to match the manual
document finance issues today (letterhead block, boxed body, no GST).

One renderer, used by both the on-screen preview and the mail attachment, so
what the customer receives is byte-identical to what was approved on screen.

# ponytail: fpdf2 rather than an HTML->PDF engine — the layout is a fixed
# boxed form, and every HTML engine worth using drags in 20 packages (or a
# browser binary) onto the prod box. Revisit only if the layout goes dynamic.
"""
import os
from decimal import Decimal, ROUND_HALF_UP

from fpdf import FPDF
from fpdf.enums import XPos, YPos

# Resolved off this file, not the working directory — a service started from
# anywhere else must still find the letterhead.
_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
HEADER_IMG = os.path.join(_ROOT, 'static', 'img', 'Invoice_address_header.png')
HEADER_W = 75.0          # mm; image is 348x116 px, height follows the ratio

# Printed only when the document carries a cargo-handling line — it is about
# that rate, and the toll-only pro forma does not carry it. Override per site
# with the FIN01 config key of the same name.
ESCALATION_NOTE = ('Note:- Differential will be charged if there is any escalation '
                   'on cargo handling rate')

# Stands in for the seal + Authorised Signatory block until finance supplies it.
# Override with the FIN01 config key of the same name.
SIGNATURE_NOTE = 'This is a system generated document and does not require a signature.'

MARGIN = 15.0
BODY_W = 180.0
COL_PART, COL_QTY, COL_RATE, COL_AMT = 90.0, 30.0, 30.0, 30.0
ROW_H = 5.0

_LATIN1 = {
    '–': '-', '—': '-', '‘': "'", '’': "'",
    '“': '"', '”': '"', '…': '...', '₹': 'Rs.',
    ' ': ' ',
}


def _l1(text):
    """Core PDF fonts are Latin-1; fold the typography we actually emit."""
    s = '' if text is None else str(text)
    for bad, good in _LATIN1.items():
        s = s.replace(bad, good)
    return s.encode('latin-1', 'replace').decode('latin-1')


def inr(n):
    """Indian digit grouping: 305843 -> '3,05,843' (paise only when nonzero)."""
    n = round(float(n or 0), 2)
    neg, n = n < 0, abs(n)
    rupees, paise = int(n), int(round((n - int(n)) * 100))
    s = str(rupees)
    if len(s) > 3:
        head, tail = s[:-3], s[-3:]
        parts = []
        while len(head) > 2:
            parts.insert(0, head[-2:])
            head = head[:-2]
        if head:
            parts.insert(0, head)
        s = ','.join(parts) + ',' + tail
    if paise:
        s += f'.{paise:02d}'
    return ('-' if neg else '') + s


def qty_fmt(q):
    return f'{float(q or 0):,.3f}'


class _PI(FPDF):
    def __init__(self):
        super().__init__(orientation='P', unit='mm', format='A4')
        self.set_margins(MARGIN, 12, MARGIN)
        self.set_auto_page_break(False)

    # --- small helpers over fpdf's cell API -----------------------------
    def row(self, cells, h=ROW_H, font=('helvetica', '', 9)):
        """cells = [(width, text, align, border), ...] laid out on one line."""
        self.set_font(*font)
        for i, (w, text, align, border) in enumerate(cells):
            last = i == len(cells) - 1
            self.cell(w, h, _l1(text), border=border, align=align,
                      new_x=XPos.LMARGIN if last else XPos.RIGHT,
                      new_y=YPos.NEXT if last else YPos.TOP)

    def body_row(self, label, qty, rate, amount, bold=False, underline=False,
                 indent=False, h=ROW_H):
        style = ('B' if bold else '') + ('U' if underline else '')
        self.set_font('helvetica', style, 9)
        self.cell(COL_PART, h, _l1(('     ' if indent else '') + (label or '')),
                  border='LR', align='L', new_x=XPos.RIGHT, new_y=YPos.TOP)
        self.set_font('helvetica', 'B' if bold else '', 9)
        for w, val in ((COL_QTY, qty), (COL_RATE, rate), (COL_AMT, amount)):
            self.cell(w, h, _l1(val or ''), border='LR', align='R',
                      new_x=XPos.RIGHT, new_y=YPos.TOP)
        self.ln(h)


def group_lines(lines):
    """Club the billable lines by service type instead of the per-parcel
    hopscotch the billables engine emits (P1/handling, P1/infra, P2/handling…).

    One row per service when every parcel shares the rate — the shape of the
    manual pro forma. When cargo-specific rates differ inside a service, the
    service becomes a heading with one row per cargo, so nothing is averaged.
    """
    order, groups = [], {}
    for l in lines:
        key = l.get('service_code') or l.get('service_name')
        if key not in groups:
            order.append(key)
            groups[key] = []
        groups[key].append(l)

    rows = []
    for key in order:
        members = groups[key]
        name = members[0].get('service_name') or key
        # GST is a property of the service, so it is constant across the group
        # and rides onto every row the group emits.
        gst = {k: members[0].get(k) for k in ('cgst_rate', 'sgst_rate', 'igst_rate')}
        rates = {round(float(l.get('rate') or 0), 4) for l in members}
        if len(rates) == 1:
            rate = float(members[0].get('rate') or 0)
            qty = round(sum(float(l.get('qty') or 0) for l in members), 3)
            rows.append({'label': name, 'indent': False, 'qty': qty,
                         'rate': rate, 'amount': round(qty * rate, 2), **gst})
            continue
        rows.append({'label': name, 'indent': False,
                     'qty': None, 'rate': None, 'amount': None, **gst})
        by_cargo, cargo_order = {}, []
        for l in members:
            ck = (l.get('cargo_name') or name, round(float(l.get('rate') or 0), 4))
            if ck not in by_cargo:
                cargo_order.append(ck)
                by_cargo[ck] = []
            by_cargo[ck].append(l)
        for ck in cargo_order:
            cargo, rate = ck
            qty = round(sum(float(x.get('qty') or 0) for x in by_cargo[ck]), 3)
            rows.append({'label': cargo, 'indent': True, 'qty': qty,
                         'rate': rate, 'amount': round(qty * rate, 2), **gst})
    return rows


def _rupees(amount):
    """Whole rupees, half up. The manual pro forma carries no paise on the tax
    lines (9% of 5,17,440 = 46,569.60 prints as 46,570), and the printed column
    has to add up for whoever checks it by hand."""
    return int(Decimal(str(amount)).quantize(Decimal('1'), rounding=ROUND_HALF_UP))


def tax_lines(rows, intra_state):
    """The 'ADD: C-GST 9%' / 'S-GST 9%' (or 'I-GST 18%') rows.

    One pair per distinct rate in play, so a pro forma mixing 18% cargo
    handling with 0%-rated toll shows the cargo tax and stays silent about the
    toll, rather than blending them into one wrong percentage.
    """
    components = ((('cgst_rate', 'C-GST'), ('sgst_rate', 'S-GST')) if intra_state
                  else (('igst_rate', 'I-GST'),))
    out = []
    for field, label in components:
        buckets = {}
        for r in rows:
            if r.get('amount') is None:
                continue
            pct = float(r.get(field) or 0)
            if pct:
                buckets[pct] = buckets.get(pct, 0.0) + r['amount']
        for pct in sorted(buckets):
            out.append({'label': f'ADD: {label} {pct:g}%',
                        'amount': _rupees(buckets[pct] * pct / 100)})
    return out


def render(ctx):
    """ctx: vessel_name, ref_no, date_str, customer{}, rows[], sac_codes,
    subtotal, amount_words, seller_gstin, seller_pan, payment_note.
    Returns the PDF as bytes."""
    pdf = _PI()
    pdf.add_page()

    # Letterhead — logo + registered address block, set to the right as on the
    # printed stationery. Natural aspect, never stretched.
    if os.path.exists(HEADER_IMG):
        pdf.image(HEADER_IMG, x=MARGIN + BODY_W - HEADER_W, y=12, w=HEADER_W)
        pdf.set_y(12 + HEADER_W * 116.0 / 348.0 + 10)
    else:
        pdf.set_y(30)

    pdf.row([(BODY_W, 'PRO-FORMA INVOICE', 'C', 1)], h=9,
            font=('helvetica', 'B', 12))

    pdf.row([(BODY_W * 0.55, f"Ref.No: {ctx['ref_no']}", 'L', 'LB'),
             (BODY_W * 0.45, f"DATE: {ctx['date_str']}", 'R', 'RB')],
            h=8, font=('helvetica', 'B', 9))

    # --- To block: variable height, so draw the text then box it ---------
    cust = ctx.get('customer') or {}
    to_lines = ['To,', cust.get('name') or '']
    for part in (cust.get('billing_address'), cust.get('city'), cust.get('pincode')):
        if part:
            to_lines.extend(str(part).splitlines())
    if cust.get('gstin'):
        to_lines.append(f"GSTIN:- {cust['gstin']}")
    box_top = pdf.get_y()
    pdf.set_font('helvetica', 'B', 9)
    pdf.set_xy(MARGIN + 1.5, box_top + 1.5)
    for line in to_lines:
        pdf.cell(BODY_W - 3, 4.4, _l1(line), border=0, align='L',
                 new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_x(MARGIN + 1.5)
    box_h = max(pdf.get_y() - box_top + 1.5, 34.0)
    pdf.rect(MARGIN, box_top, BODY_W, box_h)
    pdf.set_xy(MARGIN, box_top + box_h)

    # --- line table ------------------------------------------------------
    pdf.row([(COL_PART, 'Particulars', 'C', 1), (COL_QTY, 'Qty. in', 'C', 1),
             (COL_RATE, 'Rate', 'C', 1), (COL_AMT, 'Amount', 'C', 1)],
            h=5, font=('helvetica', 'B', 9))
    pdf.row([(COL_PART, '', 'C', 'LR'), (COL_QTY, 'MT', 'C', 1),
             (COL_RATE, '(Rs.)/MT', 'C', 1), (COL_AMT, '(Rs.)', 'C', 1)],
            h=5, font=('helvetica', 'B', 9))

    pdf.body_row(ctx.get('vessel_name') or '', '', '', '', bold=True, underline=True)
    for r in ctx['rows']:
        pdf.body_row(
            r['label'],
            qty_fmt(r['qty']) if r['qty'] is not None else '',
            f"{float(r['rate']):,.2f}" if r['rate'] is not None else '',
            inr(r['amount']) if r['amount'] is not None else '',
            bold=True, indent=r.get('indent'))

    pdf.body_row('', '', '', '')
    if ctx.get('sac_codes'):
        pdf.body_row(f"SAC Code- {ctx['sac_codes']}", '', '', '', bold=True)
    pdf.body_row('', '', '', '')

    pdf.row([(COL_PART, '', 'L', 'LTR'), (COL_QTY, 'Sub Total', 'C', 1),
             (COL_RATE, '', 'C', 1), (COL_AMT, inr(ctx['subtotal']), 'R', 1)],
            h=5.5, font=('helvetica', 'B', 9))

    # Tax sits between Sub Total and Total, labelled in the rate column.
    pdf.body_row('', '', '', '', h=6)
    for t in ctx.get('tax_rows') or []:
        pdf.body_row('', '', t['label'], inr(t['amount']), bold=True)
    pdf.body_row('', '', '', '', h=6)

    pdf.row([(COL_PART + COL_QTY + COL_RATE, 'Total :', 'R', 'LTR'),
             (COL_AMT, inr(ctx['total']), 'R', 1)],
            h=5.5, font=('helvetica', 'B', 9))

    pdf.row([(BODY_W, ctx['amount_words'], 'L', 'LRB')], h=6,
            font=('helvetica', 'B', 9))

    # --- signature / statutory footer ------------------------------------
    foot_top = pdf.get_y()
    if ctx.get('escalation_note'):
        pdf.set_xy(MARGIN + 1.5, foot_top + 2)
        pdf.set_font('helvetica', 'B', 8.5)
        pdf.multi_cell(BODY_W - 3, 4.4, _l1(ctx['escalation_note']), align='L')
    # ponytail: no seal or signature image yet — say so plainly rather than
    # printing an empty "Authorised Signatory" rule nobody has signed. Swap
    # this for the seal + signature block when finance supplies the artwork.
    pdf.set_xy(MARGIN, foot_top + 20)
    pdf.set_font('helvetica', 'B', 9)
    pdf.cell(BODY_W - 2, 5, _l1(ctx.get('signature_note') or SIGNATURE_NOTE),
             align='R', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(6)
    pdf.set_x(MARGIN + 1.5)
    pdf.cell(BODY_W - 3, 5, _l1(f"GSTIN:- {ctx['seller_gstin']}"),
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_x(MARGIN + 1.5)
    pdf.cell(BODY_W - 3, 5, _l1(f"PAN : {ctx['seller_pan']}"),
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(2)
    pdf.rect(MARGIN, foot_top, BODY_W, pdf.get_y() - foot_top)

    note_top = pdf.get_y()
    pdf.set_xy(MARGIN + 1.5, note_top + 1.5)
    pdf.set_font('helvetica', 'B', 8)
    pdf.multi_cell(BODY_W - 3, 4, _l1(ctx['payment_note']), align='L')
    pdf.rect(MARGIN, note_top, BODY_W, pdf.get_y() - note_top + 1.5)

    return bytes(pdf.output())


def demo():
    """Self-check against the real JJLTPL/PI/26-27/0487 pro forma: the clubbing,
    and the tax block reproducing its figures to the rupee."""
    # differing rates inside one service must split by cargo, never average
    split = group_lines([
        {'service_code': 'CHGL01', 'service_name': 'Cargo Handling Loading',
         'cargo_name': 'OIL A', 'qty': 100.0, 'rate': 20.0, 'cgst_rate': 9,
         'sgst_rate': 9, 'igst_rate': 18},
        {'service_code': 'CHGL01', 'service_name': 'Cargo Handling Loading',
         'cargo_name': 'OIL B', 'qty': 50.0, 'rate': 30.0, 'cgst_rate': 9,
         'sgst_rate': 9, 'igst_rate': 18},
    ])
    assert [r['label'] for r in split] == ['Cargo Handling Loading', 'OIL A', 'OIL B']
    assert split[0]['qty'] is None and split[1]['rate'] == 20.0
    assert round(sum(r['amount'] for r in split if r['amount']), 2) == 3500.0

    assert inr(71148) == '71,148' and inr(305843.5) == '3,05,843.50'

    # --- the real document: two parcels of one vessel, clubbed to two rows ---
    rows = group_lines([
        {'service_code': 'CHGL01', 'service_name': 'Cargo Handling Charges',
         'cargo_name': 'CARGO A', 'qty': 1000.0, 'rate': 252.00,
         'cgst_rate': 9, 'sgst_rate': 9, 'igst_rate': 18},
        {'service_code': 'INFM01', 'service_name': 'Infrastructure and Miscellaneous Charges',
         'cargo_name': 'CARGO A', 'qty': 1000.0, 'rate': 100.00,
         'cgst_rate': 9, 'sgst_rate': 9, 'igst_rate': 18},
        {'service_code': 'CHGL01', 'service_name': 'Cargo Handling Charges',
         'cargo_name': 'CARGO B', 'qty': 470.0, 'rate': 252.00,
         'cgst_rate': 9, 'sgst_rate': 9, 'igst_rate': 18},
        {'service_code': 'INFM01', 'service_name': 'Infrastructure and Miscellaneous Charges',
         'cargo_name': 'CARGO B', 'qty': 470.0, 'rate': 100.00,
         'cgst_rate': 9, 'sgst_rate': 9, 'igst_rate': 18},
    ])
    assert [r['label'] for r in rows] == ['Cargo Handling Charges',
                                          'Infrastructure and Miscellaneous Charges'], rows
    assert rows[0]['qty'] == 1470.0 and rows[0]['amount'] == 370440.0
    assert rows[1]['amount'] == 147000.0

    subtotal = round(sum(r['amount'] for r in rows), 2)
    assert subtotal == 517440.0
    tax = tax_lines(rows, intra_state=True)
    assert [t['label'] for t in tax] == ['ADD: C-GST 9%', 'ADD: S-GST 9%'], tax
    # 9% of 5,17,440 is 46,569.60 — the document shows 46,570, and 6,10,580
    assert [t['amount'] for t in tax] == [46570, 46570], tax
    total = round(subtotal + sum(t['amount'] for t in tax), 2)
    assert total == 610580.0, total

    # 0%-rated toll must not create a tax line, nor blend into the cargo rate
    mixed = tax_lines(rows + [{'label': 'Toll Charges', 'amount': 71148.0,
                               'cgst_rate': 0, 'sgst_rate': 0, 'igst_rate': 0}],
                      intra_state=True)
    assert [t['amount'] for t in mixed] == [46570, 46570], mixed
    # inter-state swaps the pair for a single IGST line at the combined rate
    assert [t['label'] for t in tax_lines(rows, intra_state=False)] == ['ADD: I-GST 18%']

    pdf = render({
        'vessel_name': 'MT HAFNIA HAWK', 'ref_no': 'JJLTPL/PI/26-27/0487',
        'date_str': '24.08.2026',
        'customer': {'name': 'MOTUMAL & CO', 'billing_address':
                     'A/C ADM AGRO INDUSTRIES KOTA & AKOLA PVT. LTD.\n'
                     '1 ST. FLOOR, 101, EMCA HOUSE,\n'
                     '289, SHAHID BHAGAT SINGH ROAD, FORT,',
                     'city': 'MUMBAI', 'pincode': '400 001.',
                     'gstin': '27AAFFM2481C1ZI'},
        'rows': rows, 'sac_codes': '996719', 'subtotal': subtotal,
        'tax_rows': tax, 'total': total,
        'amount_words': 'Rupees Six Lakh Ten Thousand Five Hundred Eighty Only.',
        'escalation_note': ESCALATION_NOTE,
        'seller_gstin': '27AAGCJ3665D1ZK', 'seller_pan': 'AAGCJ3665D',
        'payment_note': 'Note : Payment to be made through DD / Bankers Cheque/RTGS '
                        'drawn in favour of JSW JNPT LIQUID TERMINAL PRIVATE LIMITED, '
                        '(Axis Bank Ltd- Kalina Branch, Mumbai – 400098, Escrow '
                        'Account- 924020046923953, IFS CODE- UTIB0000776)',
    })
    assert pdf.startswith(b'%PDF') and len(pdf) > 2000, len(pdf)
    return pdf


if __name__ == '__main__':
    open('proforma_demo.pdf', 'wb').write(demo())
    print('ok — proforma_demo.pdf')
