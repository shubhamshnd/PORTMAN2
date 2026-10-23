"""VCN01 delay remarks: free-text Remarks on both delay sub-tables
(Pre Berthing/Anchoring -> vcn_delays, Berthing -> vcn_berth_delays).

Revision ID: jnpa79_vcn_delay_remarks
Revises: jnpa78_sap_sandbox_log
Create Date: 2026-10-23
"""
from typing import Sequence, Union
from alembic import op

revision: str = 'jnpa79_vcn_delay_remarks'
down_revision: Union[str, None] = 'jnpa78_sap_sandbox_log'
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.execute('ALTER TABLE vcn_delays ADD COLUMN IF NOT EXISTS remarks TEXT;')
    op.execute('ALTER TABLE vcn_berth_delays ADD COLUMN IF NOT EXISTS remarks TEXT;')


def downgrade() -> None:
    op.execute('ALTER TABLE vcn_berth_delays DROP COLUMN IF EXISTS remarks;')
    op.execute('ALTER TABLE vcn_delays DROP COLUMN IF EXISTS remarks;')
