"""Pro forma prints only the lines ticked on the billables screen (?l=...)."""
import re
from pathlib import Path

SRC = Path('modules/FIN01/views.py').read_text(encoding='utf-8')


def _filter(lines, picked):
    """Mirror of the ?l= filter in FIN01.proforma_invoice."""
    if not picked:
        return lines
    want = set(picked.split(','))
    return [l for l in lines
            if f"{l['cargo_source_type']}:{l['cargo_source_id']}:{l['service_type_id']}" in want]


def test_filter():
    lines = [{'cargo_source_type': 'VCN_IMPORT', 'cargo_source_id': 7, 'service_type_id': 1},
             {'cargo_source_type': 'VCN_IMPORT', 'cargo_source_id': 7, 'service_type_id': 2},
             {'cargo_source_type': 'VCN_EXPORT', 'cargo_source_id': 7, 'service_type_id': 1}]
    assert _filter(lines, None) == lines                          # no param -> all
    assert _filter(lines, 'VCN_IMPORT:7:2') == [lines[1]]         # one line
    # same id in the other source table must not leak in
    assert _filter(lines, 'VCN_EXPORT:7:1') == [lines[2]]
    assert _filter(lines, 'VCN_IMPORT:9:1') == []                 # -> 404 branch


def test_view_uses_same_key():
    body = SRC[SRC.index('def proforma_invoice'):]
    body = body[:body.index('\n@bp.route')]
    assert "request.args.get('l')" in body
    assert re.search(r"cargo_source_type.*cargo_source_id.*service_type_id", body, re.S)
    assert 'No lines selected' in body


if __name__ == '__main__':
    test_filter(); test_view_uses_same_key(); print('ok')
