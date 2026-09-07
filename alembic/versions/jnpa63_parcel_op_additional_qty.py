"""jnpa - additional (amended) quantity on ldud_parcel_ops

When a vessel discharges more than the declared BL, the extra tonnage is
already logged hour-by-hour in LUEU01 — it was simply discarded by the
completion cap, because the target came only from the VCN declared quantity.

Rather than editing the VCN (explicitly ruled out) or inventing a synthetic
log row (Short Close does that because its quantity was never handled), the
amendment is recorded once here and RAISES the target, so the real logged rows
start counting.

Revision ID: jnpa63_parcel_op_additional_qty
Revises: jnpa62_approval_log_proof
Create Date: 2026-09-07
"""
from typing import Sequence, Union
from alembic import op

revision: str = 'jnpa63_parcel_op_additional_qty'
down_revision: Union[str, None] = 'jnpa62_approval_log_proof'
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.execute('''
        ALTER TABLE ldud_parcel_ops
            ADD COLUMN IF NOT EXISTS additional_qty NUMERIC DEFAULT 0,
            ADD COLUMN IF NOT EXISTS additional_reason TEXT,
            ADD COLUMN IF NOT EXISTS additional_by TEXT,
            ADD COLUMN IF NOT EXISTS additional_date TEXT;
    ''')


def downgrade() -> None:
    op.execute('''
        ALTER TABLE ldud_parcel_ops
            DROP COLUMN IF EXISTS additional_qty,
            DROP COLUMN IF EXISTS additional_reason,
            DROP COLUMN IF EXISTS additional_by,
            DROP COLUMN IF EXISTS additional_date;
    ''')
