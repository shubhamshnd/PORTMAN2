"""Remove bills and invoices, and put their cargo and services back on the
billable list. Built for the SAP cutover: generate test documents against the
SAP quality server, tear them down, then start clean in production.

  python remove_billing.py                          show what is there (dry run)
  python remove_billing.py --all                    dry run of a full teardown
  python remove_billing.py --all --apply            do it
  python remove_billing.py --invoice INV2026-0001 --apply
  python remove_billing.py --bill BILL0001 --apply
  python remove_billing.py --customer JUBILANT --apply
  python remove_billing.py --all --apply --include-posted   allow SAP-posted docs

Nothing is written without --apply. The run prints which SAP system the app is
pointed at first, so a production teardown cannot be done by accident.

Scope closure: a bill and the invoice it sits on are removed together. Picking
either one pulls in the other, because invoice_lines carries a bill_id and an
invoice whose bills are gone is unbillable wreckage. To keep the bills and only
undo the invoice, use Admin -> Uninvoice instead, which reverts them to
Approved for re-invoicing.

What it reverses that the in-app delete does not: `parcel_charge_billed`.
FIN01.delete_bill still calls the deprecated `_unmark_cargo_source_billed`
no-op, so deleting a bill through the UI leaves its parcels flagged billed
forever. This tool deletes the ledger rows the bill wrote -- and only those:
rows with a NULL bill_id are go-live cutover flags and are never touched.

Scope: bills, invoices, and the flags that mark their cargo and services
billed. Nothing else. Specifically NOT touched:
  * credit/debit notes  - their own documents; the run refuses while one
                          references an in-scope invoice
  * integration_logs    - the audit trail of what was sent to SAP
  * parcels, VCNs, LDUDs, service record content - only the billed flag moves
The SAP staging and outbound-queue rows for an invoice ARE removed: they are
that invoice own outbound plumbing, and a queued job for a deleted invoice
would post to SAP.

Document numbers are not sequences here; they are derived from MAX() over the
live tables, floored at the cutover seed (see FIN01.next_from_seed). Removing
documents therefore frees their numbers with no sequence reset, which is what
makes a QA run safe to discard.
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from database import get_db, get_cursor


def sap_target(cur):
    """The SAP system the app would post to right now."""
    try:
        cur.execute("""SELECT environment, base_url FROM sap_api_config
                       WHERE COALESCE(is_active, 0) = 1 ORDER BY id LIMIT 1""")
        row = cur.fetchone()
    except Exception:
        cur.connection.rollback()
        return None
    return dict(row) if row else None


def resolve_scope(cur, args):
    """Bill and invoice ids in scope, closed over the bill<->invoice link.

    Picking an invoice pulls in its bills and vice versa, repeatedly, so a
    bill shared by two invoices takes both with it rather than leaving one
    pointing at rows that no longer exist.
    """
    bills, invoices = set(), set()

    if args.all:
        cur.execute('SELECT id FROM bill_header')
        bills = {r['id'] for r in cur.fetchall()}
        cur.execute('SELECT id FROM invoice_header')
        invoices = {r['id'] for r in cur.fetchall()}
        return bills, invoices

    if args.invoice:
        cur.execute('SELECT id FROM invoice_header WHERE invoice_number = ANY(%s)', [args.invoice])
        invoices |= {r['id'] for r in cur.fetchall()}
    if args.bill:
        cur.execute('SELECT id FROM bill_header WHERE bill_number = ANY(%s)', [args.bill])
        bills |= {r['id'] for r in cur.fetchall()}
    if args.customer:
        cur.execute('SELECT id FROM bill_header WHERE customer_name = ANY(%s)', [args.customer])
        bills |= {r['id'] for r in cur.fetchall()}
        cur.execute('SELECT id FROM invoice_header WHERE customer_name = ANY(%s)', [args.customer])
        invoices |= {r['id'] for r in cur.fetchall()}

    while True:
        before = (len(bills), len(invoices))
        if invoices:
            cur.execute('SELECT bill_id FROM invoice_bill_mapping WHERE invoice_id = ANY(%s)',
                        [list(invoices)])
            bills |= {r['bill_id'] for r in cur.fetchall() if r['bill_id']}
        if bills:
            cur.execute('SELECT invoice_id FROM invoice_bill_mapping WHERE bill_id = ANY(%s)',
                        [list(bills)])
            invoices |= {r['invoice_id'] for r in cur.fetchall() if r['invoice_id']}
        if (len(bills), len(invoices)) == before:
            return bills, invoices


def posted_documents(cur, invoice_ids):
    """In-scope invoices that already carry a SAP document number."""
    if not invoice_ids:
        return []
    cur.execute("""SELECT invoice_number, sap_document_number
                   FROM invoice_header
                   WHERE id = ANY(%s) AND COALESCE(TRIM(sap_document_number), '') <> ''
                   ORDER BY invoice_number""", [list(invoice_ids)])
    return [dict(r) for r in cur.fetchall()]


def credit_notes(cur, invoice_ids):
    """Credit/debit notes raised against the in-scope invoices.

    These are separate financial documents with their own numbers and their own
    SAP postings, so they are not ours to delete as a side effect. The FK is ON
    DELETE NO ACTION, so an invoice cannot go while one exists — the run stops
    and says so rather than destroying it.
    """
    if not invoice_ids:
        return []
    cur.execute("""SELECT doc_number, doc_type, original_invoice_number,
                          sap_document_number
                   FROM fdcn_header WHERE original_invoice_id = ANY(%s)
                   ORDER BY doc_number""", [list(invoice_ids)])
    return [dict(r) for r in cur.fetchall()]


def survey(cur, bill_ids, invoice_ids):
    """Row counts each step would touch, in the order they are applied."""
    b, i = list(bill_ids), list(invoice_ids)
    steps = []

    def count(label, sql, params):
        if not params or not params[0]:
            steps.append((label, 0))
            return
        cur.execute(sql, params)
        steps.append((label, cur.fetchone()['n']))

    count('SAP staging rows', 'SELECT COUNT(*) n FROM invoice_sap_staging WHERE invoice_id = ANY(%s)', [i])
    count('SAP outbound queue jobs', 'SELECT COUNT(*) n FROM sap_outbound_queue WHERE invoice_id = ANY(%s)', [i])
    count('invoice lines', 'SELECT COUNT(*) n FROM invoice_lines WHERE invoice_id = ANY(%s)', [i])
    count('invoice-bill mappings', 'SELECT COUNT(*) n FROM invoice_bill_mapping WHERE invoice_id = ANY(%s)', [i])
    count('invoices', 'SELECT COUNT(*) n FROM invoice_header WHERE id = ANY(%s)', [i])
    count('parcel ledger rows freed', 'SELECT COUNT(*) n FROM parcel_charge_billed WHERE bill_id = ANY(%s)', [b])
    count('service records freed', 'SELECT COUNT(*) n FROM service_records WHERE bill_id = ANY(%s)', [b])
    count('bill lines', 'SELECT COUNT(*) n FROM bill_lines WHERE bill_id = ANY(%s)', [b])
    count('bill-vessel links', 'SELECT COUNT(*) n FROM bill_vessels WHERE bill_id = ANY(%s)', [b])
    count('bills', 'SELECT COUNT(*) n FROM bill_header WHERE id = ANY(%s)', [b])
    return steps


def remove(cur, bill_ids, invoice_ids):
    """Delete in FK order and reverse the billed flags. Caller commits."""
    b, i = list(bill_ids), list(invoice_ids)
    done = []

    def run(label, sql, params):
        if not params or not params[0]:
            return
        cur.execute(sql, params)
        done.append((label, cur.rowcount))

    # Invoices first: invoice_lines carries a bill_id, so the bills cannot go
    # until the invoice rows that point at them are gone.
    #
    # Credit/debit notes are deliberately NOT removed here - see credit_notes().
    # The staging and queue rows below are: they are this invoice own outbound
    # plumbing, and a pending job for a deleted invoice would post to SAP.
    run('SAP staging rows', 'DELETE FROM invoice_sap_staging WHERE invoice_id = ANY(%s)', [i])
    run('SAP outbound queue jobs', 'DELETE FROM sap_outbound_queue WHERE invoice_id = ANY(%s)', [i])
    run('invoice lines', 'DELETE FROM invoice_lines WHERE invoice_id = ANY(%s)', [i])
    run('invoice-bill mappings', 'DELETE FROM invoice_bill_mapping WHERE invoice_id = ANY(%s)', [i])
    run('invoices', 'DELETE FROM invoice_header WHERE id = ANY(%s)', [i])

    # bill_id IS NOT NULL guards the go-live cutover flags, which are exactly
    # the rows with no bill behind them and must survive a billing teardown.
    run('parcel ledger rows freed',
        'DELETE FROM parcel_charge_billed WHERE bill_id = ANY(%s) AND bill_id IS NOT NULL', [b])
    run('service records freed',
        'UPDATE service_records SET is_billed = 0, bill_id = NULL WHERE bill_id = ANY(%s)', [b])

    run('bill lines', 'DELETE FROM bill_lines WHERE bill_id = ANY(%s)', [b])
    run('bill-vessel links', 'DELETE FROM bill_vessels WHERE bill_id = ANY(%s)', [b])
    run('bills', 'DELETE FROM bill_header WHERE id = ANY(%s)', [b])
    return done


def main(argv=None):
    p = argparse.ArgumentParser(
        description='Remove bills and invoices and free their cargo/services.')
    p.add_argument('--all', action='store_true', help='every bill and invoice')
    p.add_argument('--invoice', action='append', metavar='NUMBER', help='invoice number (repeatable)')
    p.add_argument('--bill', action='append', metavar='NUMBER', help='bill number (repeatable)')
    p.add_argument('--customer', action='append', metavar='NAME', help='all documents of a party (repeatable)')
    p.add_argument('--apply', action='store_true', help='actually delete (default is a dry run)')
    p.add_argument('--include-posted', action='store_true',
                   help='allow removing invoices that already have a SAP document number')
    p.add_argument('--yes', action='store_true', help='skip the confirmation prompt')
    args = p.parse_args(argv)

    conn = get_db()
    cur = get_cursor(conn)
    try:
        target = sap_target(cur)
        print('\nSAP target :', (f"{target['environment'] or '?'}  {target['base_url'] or ''}".strip()
                                 if target else 'no active sap_api_config row'))

        if not (args.all or args.invoice or args.bill or args.customer):
            cur.execute('SELECT COUNT(*) n FROM bill_header')
            nb = cur.fetchone()['n']
            cur.execute('SELECT COUNT(*) n FROM invoice_header')
            ni = cur.fetchone()['n']
            print(f'\nIn the database: {nb} bill(s), {ni} invoice(s).')
            print('Nothing selected. Use --all, --invoice, --bill or --customer.')
            return 0

        bill_ids, invoice_ids = resolve_scope(cur, args)
        if not bill_ids and not invoice_ids:
            print('\nNothing matched that selection.')
            return 0

        print(f'\nScope: {len(bill_ids)} bill(s), {len(invoice_ids)} invoice(s)'
              '   (bills and their invoices are removed together)')

        posted = posted_documents(cur, invoice_ids)
        if posted:
            print(f'\n{len(posted)} invoice(s) already posted to SAP:')
            for d in posted[:10]:
                print(f"    {d['invoice_number']:24} SAP doc {d['sap_document_number']}")
            if len(posted) > 10:
                print(f'    ... and {len(posted) - 10} more')
            if not args.include_posted:
                print('\nRefusing: these carry a SAP document number. Removing them here does'
                      '\nNOT reverse anything in SAP -- that needs a reversal or credit note.'
                      '\nRe-run with --include-posted once SAP is settled.')
                return 1
            print('  --include-posted given: removing them anyway.')

        notes = credit_notes(cur, invoice_ids)
        if notes:
            print(f'\n{len(notes)} credit/debit note(s) reference these invoices:')
            for d in notes[:10]:
                sap = d['sap_document_number'] or 'not posted'
                print(f"    {d['doc_number']:22} {d['doc_type'] or '':3} against "
                      f"{d['original_invoice_number'] or '':20} SAP {sap}")
            if len(notes) > 10:
                print(f'    ... and {len(notes) - 10} more')
            print('\nRefusing: a note is its own document, with its own number and'
                  '\npossibly its own SAP posting, so this tool will not delete one as a'
                  '\nside effect. Remove the note(s) in FDCN01 first, then re-run.')
            return 1

        print('\nWould remove:' if not args.apply else '\nRemoving:')
        for label, n in survey(cur, bill_ids, invoice_ids):
            print(f'    {label:28}: {n}')

        if not args.apply:
            print('\nDry run -- nothing was changed. Add --apply to do it.')
            return 0

        if not args.yes:
            where = (target or {}).get('environment') or 'UNKNOWN'
            print(f'\nThis permanently deletes the documents above. SAP target: {where}')
            if input('Type YES to confirm: ').strip() != 'YES':
                print('Aborted.')
                return 1

        done = remove(cur, bill_ids, invoice_ids)
        conn.commit()
        print('\nDone:')
        for label, n in done:
            print(f'    {label:28}: {n}')
        print('\nCargo and services are billable again. Document numbers are derived'
              '\nfrom the live tables, so the freed numbers will be reused.')
        return 0
    except Exception:
        conn.rollback()
        print('\nFailed -- nothing was changed.')
        raise
    finally:
        conn.close()


if __name__ == '__main__':
    sys.exit(main())
