"""jnpa - mail attachments, and repair dangling finance_service_types.gst_rate_id

Two things, both needed before FIN01 can mail a pro-forma invoice that has the
right tax on it:

1. mail_queue carries one optional attachment (base64 in the row, so a retry
   three attempts later still has the bytes) — the pro-forma PDF.

2. A safety net for finance_service_types.gst_rate_id, which this project (not
   the live PORTMAN one) seeds from HARD-CODED gst_rates ids in
   jnpa37_seed_finance_services (GST_18 = 4, GST_0 = 1). Those assume the
   745e51f340e0 rate seed ran first and exactly once — and that seed has no
   ON CONFLICT guard, so on a database where the rates were re-seeded or edited
   through FGRM01 the ids can land on a different rate or on no row at all.

   The GST calculation in FIN01.save_bill_line is byte-identical to the live
   system's and is left exactly as it is. But it reads
   `SELECT ... FROM gst_rates WHERE id = <that id>` and, finding nothing, falls
   through without computing tax — so a dangling id shows up as a bill with no
   GST and no error. Repair by matching on the RATE rather than the id:
   cargo/infra services to whichever active row is 18%, toll to whichever is 0%.

   Only rows whose gst_rate_id is NULL or dangling are touched, so this is a
   no-op on a healthy database — a rate somebody deliberately set to something
   else is left alone.

Revision ID: jnpa65_mail_attachments_gst
Revises: jnpa64_parcel_removed_flag
Create Date: 2026-09-11
"""
from typing import Sequence, Union

from alembic import op

revision: str = 'jnpa65_mail_attachments_gst'
down_revision: Union[str, None] = 'jnpa64_parcel_removed_flag'
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None

# service_code -> the igst_rate its GST row must carry
_EXPECTED_RATE = {
    'CHGU01': 18, 'CHGL01': 18, 'INFM01': 18, 'MLAC01': 18, 'SHGW01': 18,
    'TOLL01': 0,
}


def upgrade() -> None:
    op.execute("""
        ALTER TABLE mail_queue
          ADD COLUMN IF NOT EXISTS attachment_name TEXT,
          ADD COLUMN IF NOT EXISTS attachment_b64  TEXT
    """)

    for code, igst in _EXPECTED_RATE.items():
        op.execute(f"""
            UPDATE finance_service_types s
               SET gst_rate_id = (
                     SELECT g.id FROM gst_rates g
                      WHERE g.is_active = 1 AND g.igst_rate = {igst}
                      ORDER BY g.id
                      LIMIT 1)
             WHERE s.service_code = '{code}'
               AND NOT EXISTS (
                     SELECT 1 FROM gst_rates g2 WHERE g2.id = s.gst_rate_id)
        """)


def downgrade() -> None:
    op.execute("""
        ALTER TABLE mail_queue
          DROP COLUMN IF EXISTS attachment_name,
          DROP COLUMN IF EXISTS attachment_b64
    """)
