# VCN01 Operation-Type Guards — MT Hodaka Galaxy incident

> **Status (2026-09-07):** Part A Phase 1 and Part B Phases 6–8 are **implemented
> and tested** (`tests/test_ldud01_cargo_by_optype.py`,
> `tests/test_reopen_proof.py`, updated `tests/test_finance_parity_integration.py`).
> Part A Phases 2–5 — the op-type-change confirm, the required-op-type guard, the
> send-to-Expected block, and the corrupt-data report — are **still to do**.


Incident: op type left as Import, parcels entered, op type flipped to Export,
parcels re-entered → LDUD showed ~2x quantity. Vessel deleted as a workaround.

## Root causes (found in code, not guessed)

1. **The doubling is a real double-count, not stale UI.**
   `modules/LDUD01/model.py` `get_data` builds `vcn_cargo` from THREE sources —
   `vcn_cargo_declaration`, `vcn_export_cargo_declaration`, `vcn_consigners` —
   with **no `operation_type` filter**, then sums them into
   `bl_quantities_display`. A VCN carrying stale import consigners *and* new
   export rows shows the sum of both.
   Same bug in `modules/VCN01/model.py::_sync_header_cargo`, which UNIONs both
   tables into `vcn_header.cargo_type`.

2. **Flipping op type orphans parcels silently.**
   Import parcels = `vcn_consigners`, export parcels =
   `vcn_export_cargo_declaration`. `operation_type` is a plain Tabulator list
   editor (`vcn01.html`) saved through the generic header UPDATE. Nothing
   deletes, migrates or even warns about the old table's rows.
   Worse: `ldud_parcel_ops.parcel_ids` is a CSV of ids resolved **at read time**
   by `_parcel_table_for_ldud`. After a flip the same integers point into a
   different table — LDUD ops silently re-bind to unrelated parcels or render
   as `#123`.

3. **Blank op type silently means Import.**
   `openDetailsModal`: `const operationType = data.operation_type || 'Import'`.
   Parcels entered before the op type is set land in `vcn_consigners`. This is
   exactly incident step 4.

4. **Send back to Expected discards work with only a `confirm()`.**
   `send_back_to_expected` soft-deletes the LDUD, NULLs its `vcn_id`, then hard
   -deletes `vcn_header` (parcels cascade away). Re-fetching from EV01 creates a
   **new** vcn_header id, so the archived LDUD/LUEU can never re-attach. Incident
   step 2.

## Plan

### Phase 1 — Stop the double-count (highest value, no UI) — DONE
- **1a** `LDUD01/model.py::get_data` — select the cargo source by the VCN's
  `operation_type` (Export → export declaration; else consigners + legacy
  `vcn_cargo_declaration`). One extra column on the existing `vcn_header`
  fetch, no new query.
- **1b** `VCN01/model.py::_sync_header_cargo` — same: single source by op type,
  drop the UNION.
- Test: `tests/test_ldud01_cargo_by_optype.py` — VCN with rows in both tables
  reports only its own side's total.

### Phase 2 — Op-type change guard (the requested prompt)
- **2a** Server (`VCN01/model.py::save_header`): on an existing row, if
  `operation_type` differs from stored AND the old side's parcel table has rows
  → raise with the count. `views.save` returns **409** with
  `{needs_confirm: true, parcel_count, from, to}`.
- **2b** Confirmed retry: client resends with `confirm_parcel_wipe: true`;
  server deletes the old-side parcels **and** the `ldud_parcel_ops` rows whose
  `parcel_ids` reference them, inside `save_header`'s transaction. Without the
  ops cleanup the flip just re-points them at the wrong table.
- **2c** Client (`vcn01.html::saveAll`): on that 409, show
  `confirm("Changing Import → Export deletes N import parcel(s) and their
  Loading/Unloading operations. Import and export parcels are different
  records and cannot be converted. Continue?")` and retry with the flag.
  Driving it from the 409 covers every save path, not just the grid cell.
- **2d** Refuse the flip outright when `doc_status='Approved'` or the VCN is
  billed (`_billed_locked` already covers billed; add the Approved check) —
  those parcels are already closed/invoiced.
- **2e** Log the flip to `approval_log` (module VCN01, action `Op Type Changed`,
  comment `Import → Export, N parcels deleted`). The table and viewer already
  exist; ~3 lines.
- Test: `tests/test_vcn01_optype_change.py` — flip without flag → 409 with
  count; with flag → old parcels and their `ldud_parcel_ops` gone.

### Phase 3 — Op type required before detail entry
- **3a** Drop the `|| 'Import'` default; `openDetailsModal` refuses to open with
  a blank op type: "Set Operation Type on this row first."
- **3b** Real guard is server-side: `consigners/save` and `export_cargo/save`
  reject when the header's op type is blank or doesn't match the target table
  (400, "Set Operation Type to Import/Export before adding parcels").
- Test: parcel save against a blank-op-type VCN → 400.

### Phase 4 — Send back to Expected
- **4a** `send_to_expected` refuses when the VCN's LDUD has any
  `ldud_parcel_ops` or `lueu_parcel_log` rows: "Vessel has N Loading/Unloading
  operation(s) and M logbook entries. Clear them in LDUD01/LUEU01 first."
  A freshly-moved VCN with nothing entered still goes back in one click.
- Test: VCN with a parcel op → 409; without → succeeds.

### Phase 5 — Clean up existing damage
- **5a** Throwaway script: list VCNs holding rows in **both** parcel tables
  (the flip signature). Report only — an operator decides which side is real.
  Not a migration, not scheduled.

## Deliberately skipped
- Converting import parcels → export parcels on flip. Different shapes (BL
  no/date are import-only) and the quantities usually change anyway; deleting
  with an explicit prompt is honest, converting silently is not.
- A `source_table` column on `ldud_parcel_ops`. Phase 2 removes the ambiguity
  it would guard against. Add it if parcels ever need to span both tables.
- Rebuilding RP01/FINV01 against parcels — separate, already-known work.

## Order
1 → 2 → 3 → 4 → 5. Phase 1 alone fixes the number on screen; Phase 2 stops the
incident recurring. Each phase ships independently.

---

# Part B — Reopen-to-Draft lockdown (admin only, proof required)

Separate ask, same incident family: reopening an approved/closed record is what
let the vessel be edited into an inconsistent state.

**Decision:** the Admin panel's existing **Vessel Closure** tab is the one and
only reopen path. No new tab. Module-level send-back goes away entirely.

**Assumption to confirm:** that tab lists LDUD rows only today. Since an
approved VCN must still be amendable (a shipping-bill amendment is what started
this incident), Phase 8 extends the same tab to list Approved VCN01 rows
alongside the LDUD ones. If approved VCNs should instead be permanently
un-reopenable, drop 8b and the tab stays LDUD-only.

## Current state (read from code)

- **VCN01** `views.send_back` — gated on `approver_id` **or** `is_admin`, text
  comment, writes `approval_log` action `Back to Draft`. UI: `↩ Draft` button on
  Approved rows (`vcn01.html`, `_approval_actions`) + send-back modal.
- **LDUD01** `views.reopen` — same approver-or-admin gate, text reason,
  `Force Reopen (Billed)` admin override, emails the closer. UI: Reopen modal.
- **Admin `vessels` tab** — `/admin/api/ldud/open_vessel`, Closed/Partial →
  Draft, admin-only, logs to `approval_log`, and **silently deletes every
  `ldud_proof_documents` row**.
- **Admin `mbc` tab is dead UI.** MBC01 no longer exists; `/admin/api/mbc/
  approvals` and `/admin/api/mbc/reset_approval` are **not defined anywhere** —
  the tab 404s on click. Deleting it is removal only, no backend change.
- **Audit table** is `approval_log`. AUD01 is an nginx *access-log* viewer,
  unrelated.
- **Proof-upload precedent**: `ldud_proof_documents` (BYTEA via
  `s8t9u0v1w2x3_ldud_proof_bytea`), `vcn_header.igm_document` BYTEA. Blob-in-
  Postgres is the house style; no filesystem.

## Phase 6 — Delete MBC from Admin — DONE

- **6a** `templates/admin.html`: drop the `mbc` tab button, the `mbc-tab` div,
  the `showTab` branch, and `loadMbcApprovals` / its reset handler (~45 lines).
- **6b** `modules/AUD01/views.py`: drop the dead `_MODULE_NAMES` keys `MBC01`,
  `MBCM01`, `MBCDS01`.
- Nothing else to do — there is no MBC backend left to remove.

## Phase 7 — Remove send-to-draft from VCN01 and LDUD01 — DONE

- **7a** VCN01: delete `POST /api/module/VCN01/send_back`, the `↩ Draft`
  button, and the send-back modal. Approvers keep Approve, lose un-approve.
- **7b** LDUD01: delete `POST /api/module/LDUD01/reopen`, its modal, and
  `openReopenModal`. Keep `reopen_record` / `log_closure_action` in the model —
  the Admin endpoint calls them.
- **7c** Move the closer-notification email (`_get_closer_email` + `_queue_mail`)
  to the Admin endpoint so the operator still learns their record was reopened.
- **7d** Keep `Force Reopen (Billed)` as a distinct `approval_log` action so a
  billed reopen stays greppable; it becomes admin-only-by-construction.
- Delete the routes outright rather than leaving guarded-but-dead ones.

## Phase 8 — Reopen tab: VCN rows + mandatory proof — DONE

- **8a** Migration `jnpa49_approval_log_proof`:
  `ALTER TABLE approval_log ADD COLUMN IF NOT EXISTS proof_bytes BYTEA,
   proof_filename TEXT, proof_mime TEXT`. No new table, no join, existing
   readers unaffected.
  **Codebase constraint**: `approval_log` reads must stay explicit-column
  (`get_approval_log` already is). A `SELECT *` would drag BYTEA into a JSON
  response — the trap `VCN01.get_data` already dodges with
  `r.pop('igm_document')`.
- **8b** `/admin/api/ldud/vessels` becomes `/admin/api/reopen/pending`,
  returning both: Approved `vcn_header` rows and Closed/Partial Close/Approved
  `ldud_header` rows (excluding `is_deleted`), each tagged `module`, doc num,
  vessel, status, and a billed flag from `fin_model.is_vcn_billed`.
- **8c** Row action opens a modal: **reason (required, as today)** + **image
  upload (required, new)** + preview thumbnail. Multipart POST to
  `/admin/api/reopen` with `module`, `id`, `comment`, `file`. Server: admin
  check → validation → status flip → `approval_log` insert carrying the image →
  closer email. One transaction. `open_vessel` is replaced by it, not kept
  alongside — otherwise there are two ways to reopen an LDUD, one without proof.
- **8d** Validation at the boundary, not negotiable: file present, extension in
  `{.png, .jpg, .jpeg}`, `Content-Type` starts with `image/`, size ≤ 5 MB.
  Missing or wrong → 400, plain message. Mirrors the IGM PDF check.
- **8e** Confirm text must name the side effects the current code performs
  silently: for LDUD, *"N proof-of-quantity document(s) will be permanently
  deleted"*; for a billed record, *"the billing lock stays in force"*.
- **8f** `GET /api/approval-proof/<int:log_id>` → `send_file` inline, **login
  required, not admin-only**. An audit trail only the actor can read is not an
  audit trail. `get_approval_log` returns `has_proof`; the VCN01
  `showApprovalLog` popover and the LDUD01 closure-log render a `📎 Proof` link.
- **8g** Retitle the tab "Vessel Closure" → **"Reopen to Draft"**; it now covers
  both modules.

## Tests

- `tests/test_reopen_proof.py`: non-admin → 403; admin without a file → 400;
  admin with a `.pdf` → 400; admin with a small PNG → status is Draft and the
  `approval_log` row carries `proof_bytes` + filename.
- Assert `get_approval_log` output stays JSON-serialisable (no raw bytes).

## Deliberately skipped
- Filesystem/S3 proof storage. Reopens are rare and the codebase already stores
  PDFs as BYTEA. Revisit if proofs get bulky or numerous.
- Server-side thumbnailing. 5 MB cap + browser preview covers it.
- A generic "reopen any module" framework. Two modules, two branches in one
  endpoint. Add the third when there is a third.
- Multi-image proof. One image was the ask; a follow-up table can be added later
  without touching callers.

## Open questions
1. Confirm the assumption above: should Approved VCN01 rows appear in the
   Reopen tab, or become permanently un-reopenable?
2. On a VCN01 reopen, does anything need re-validating, or is Draft enough?
3. LDUD emails the closer on reopen — who gets notified on a **VCN01** reopen?
   There is no equivalent lookup today; `created_by` is the obvious candidate.

## Order
Part A 1 → 2 → 3 → 4 → 5, then Part B 6 → 7 → 8. Part B Phase 6 is a pure
deletion and can ship any time.
