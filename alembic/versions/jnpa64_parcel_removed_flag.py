"""jnpa - removable (flagged, not deleted) VCN parcels

A parcel cancelled after declaration is flagged rather than deleted, so it
stops counting toward billing, closure and targets while staying visible in
every screen (greyed) and restorable.

The flag lives on the VCN parcel row rather than ldud_parcel_ops because
FIN01.get_billables selects billable parcels straight from these two tables —
ldud_parcel_ops only supplies the actual quantity. Flagging the op would leave
the parcel billable at its full DECLARED quantity, the opposite of the intent.
The operator still never opens VCN01: LUEU01 sets the flag.

Revision ID: jnpa64_parcel_removed_flag
Revises: jnpa63_parcel_op_additional_qty
Create Date: 2026-09-07
"""
from typing import Sequence, Union
from alembic import op

revision: str = 'jnpa64_parcel_removed_flag'
down_revision: Union[str, None] = 'jnpa63_parcel_op_additional_qty'
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None

_TABLES = ('vcn_consigners', 'vcn_export_cargo_declaration')


def upgrade() -> None:
    for t in _TABLES:
        op.execute(f'''
            ALTER TABLE {t}
                ADD COLUMN IF NOT EXISTS is_removed BOOLEAN DEFAULT FALSE,
                ADD COLUMN IF NOT EXISTS removed_reason TEXT,
                ADD COLUMN IF NOT EXISTS removed_by TEXT,
                ADD COLUMN IF NOT EXISTS removed_date TEXT;
        ''')


def downgrade() -> None:
    for t in _TABLES:
        op.execute(f'''
            ALTER TABLE {t}
                DROP COLUMN IF EXISTS is_removed,
                DROP COLUMN IF EXISTS removed_reason,
                DROP COLUMN IF EXISTS removed_by,
                DROP COLUMN IF EXISTS removed_date;
        ''')
