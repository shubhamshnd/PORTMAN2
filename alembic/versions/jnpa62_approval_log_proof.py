"""jnpa - proof image on approval_log

Reopening an approved/closed record is now admin-only and requires a photo as
evidence (MT Hodaka Galaxy, 2026-09-05). The image rides on the audit row
itself rather than a side table — reopens are rare and one row is one proof.

NOTE for callers: approval_log reads must stay explicit-column. A SELECT * now
drags BYTEA into JSON responses and breaks serialisation.

Revision ID: jnpa62_approval_log_proof
Revises: jnpa61_cutover_tables
Create Date: 2026-09-07
"""
from typing import Sequence, Union
from alembic import op

revision: str = 'jnpa62_approval_log_proof'
down_revision: Union[str, None] = 'jnpa61_cutover_tables'
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.execute('''
        ALTER TABLE approval_log
            ADD COLUMN IF NOT EXISTS proof_bytes BYTEA,
            ADD COLUMN IF NOT EXISTS proof_filename TEXT,
            ADD COLUMN IF NOT EXISTS proof_mime TEXT;
    ''')


def downgrade() -> None:
    op.execute('''
        ALTER TABLE approval_log
            DROP COLUMN IF EXISTS proof_bytes,
            DROP COLUMN IF EXISTS proof_filename,
            DROP COLUMN IF EXISTS proof_mime;
    ''')
