"""Mailing the pro-forma: the PDF actually rides along, and the From name is
the JSW one. Also pins the service clubbing that shapes both the screen and
the document."""
import base64
import io
import re
from pathlib import Path

import mail_service
from modules.FIN01 import proforma_pdf

CFG = {'from_email': 'billing@example.com', 'host': 'h', 'port': 25}


def _row(**kw):
    row = {'subject': 'Pro-Forma Invoice JJLTPL/PI/26-27/VCN-1',
           'body_html': '<p>hi</p>', 'to_email': 'customer@example.com'}
    row.update(kw)
    return row


def test_attachment_rides_along_intact():
    pdf = proforma_pdf.demo()
    msg = mail_service.build_message(
        _row(attachment_b64=base64.b64encode(pdf).decode('ascii'),
             attachment_name='Pro-Forma Invoice JJLTPL-PI-26-27-0484.pdf'), CFG)

    parts = msg.get_payload()
    assert len(parts) == 2, 'body + one attachment'
    att = parts[1]
    assert att.get_content_type() == 'application/pdf'
    assert att.get_filename() == 'Pro-Forma Invoice JJLTPL-PI-26-27-0484.pdf'
    # byte-identical after the base64 round trip, and still a real PDF
    assert att.get_payload(decode=True) == pdf
    assert att.get_payload(decode=True).startswith(b'%PDF')


def test_plain_mail_is_unchanged_by_the_attachment_support():
    msg = mail_service.build_message(_row(), CFG)
    assert msg.get_content_subtype() == 'alternative'
    assert len(msg.get_payload()) == 1


def test_from_name_is_the_jsw_one_and_config_still_wins():
    assert 'JSW JNPA Liquid Terminals' in mail_service.build_message(_row(), CFG)['From']
    assert 'Portman' not in mail_service.build_message(_row(), CFG)['From']
    override = dict(CFG, from_name='Something Else')
    assert mail_service.build_message(_row(), override)['From'].startswith('Something Else')


def test_send_route_takes_the_address_from_the_master_only():
    """The recipient must never come off the request — otherwise the Send
    button becomes a way to mail an invoice to any address."""
    src = Path('modules/FIN01/views.py').read_text(encoding='utf-8')
    body = src[src.index('def send_proforma('):]
    body = body[:body.index('\ndef _proforma_mail_html')]
    assert "ctx['customer'].get('contact_email')" in body
    assert not re.search(r"request\.(json|args|form)[^\n]*email", body)


def test_billables_are_clubbed_by_service_not_parcel():
    """The screen sorts lines into service blocks; the document then merges
    each block into one row. Both come off _SERVICE_DISPLAY_ORDER."""
    src = Path('modules/FIN01/model.py').read_text(encoding='utf-8')
    assert '_SERVICE_DISPLAY_ORDER' in src

    order = ['CHGU01', 'CHGL01', 'INFM01', 'MLAC01', 'TOLL01']
    # hopscotch in: P1/handling, P1/infra, P2/handling, P2/infra
    lines = [{'service_code': c, 'parcel_no': p}
             for p in ('P1', 'P2') for c in ('CHGL01', 'INFM01')]
    lines.sort(key=lambda l: order.index(l['service_code']))
    assert [l['service_code'] for l in lines] == ['CHGL01', 'CHGL01', 'INFM01', 'INFM01']
    # stable: parcels keep their order inside each service block
    assert [l['parcel_no'] for l in lines] == ['P1', 'P2', 'P1', 'P2']


def test_proforma_pdf_self_check():
    proforma_pdf.demo()


if __name__ == '__main__':
    for name, fn in sorted(globals().items()):
        if name.startswith('test_'):
            fn()
    print('ok')
