# Feasibility — LUEU01 Additional Quantity + Remove Parcel, and the billed-reopen rule

> **Status 2026-09-07: ALL FOUR STEPS IMPLEMENTED.** 162 tests pass (2 failures
> pre-date this work). Migrations `jnpa62`–`jnpa64` applied. New tests:
> `test_vcn01_approved_lock.py`, `test_lueu01_additional_qty.py`,
> `test_parcel_removal.py`; `test_finance_parity_integration.py` updated.

Assessment against the current code for the three discussion points.
**Verdict: all three are feasible. Two carry a correctness trap that must be
designed around, not discovered later.**

---

## 1. Send to Draft — admin only, non-billed only

**Decided 2026-09-07: a billed vessel is NEVER reopened. The point of the rule
is to stop a bill being reopened at all, so there is no override, no admin
escape hatch, no `Force Reopen (Billed)`.**

Already done: send-back removed from VCN01 and LDUD01, moved to Admin ▸ Reopen
to Draft, admin-only, mandatory reason + proof image on the audit row.

### Still to change

What is built today still carries the old behaviour — an admin *can* reopen a
billed vessel, logged as `Force Reopen (Billed)`. That was inherited from the
pre-existing `LDUD01.reopen` and must now go.

| # | Where | Change |
|---|---|---|
| 1.1 | `ADMIN/views.py::reopen_record` | billed → **409** and stop. Delete the `Force Reopen (Billed)` branch; `action` is always `Back to Draft`. |
| 1.2 | `ADMIN/views.py::reopen_pending` | keep the `is_billed` flag — the row stays visible so an admin can see *why* it cannot be reopened. |
| 1.3 | `templates/admin.html` | billed row: disable the button, keep the `BILLED` badge, show "cancel the bill first" instead of the current warning-and-proceed. |
| 1.4 | `tests/test_finance_parity_integration.py` | currently asserts the billed override **succeeds**; flip to asserting the 409 refusal. |
| 1.5 | `tests/test_reopen_proof.py` | add: billed vessel → 409, status unchanged. |

~20 lines plus the two tests.

**Consequence, accepted:** a wrongly-billed vessel can only be corrected by
cancelling the bill first. There is no longer any path that edits a billed
record.

---

## 1b. VCN01 edit policy — Draft open, Approved shut

**Decided 2026-09-07: "The VCN will not be able to be edited other than admin
reopening a non-billed vessel; if it is in Draft, let them do whatever they
want."**

Three states, one rule each:

| State | Rule |
|---|---|
| **Draft** | Fully editable. No extra gating beyond module permissions. |
| **Approved** | Locked to *everyone*, approver included. The only way to edit is an admin reopen to Draft. |
| **Billed** | Permanently locked. No reopen (§1). |

### Gap — the current code does not enforce this

**(a) An approver can still edit an Approved VCN.** `VCN01/views.py::save`:

```
if current_status == 'Approved':
    if not is_approver:
        return jsonify({'error': 'Cannot edit an approved record'}), 403
    data['doc_status'] = 'Approved'          # approver edits in place
```

Under the new rule this branch goes away — Approved is a 403 for everyone, with
a message pointing at Admin ▸ Reopen to Draft.

**(b) The parcel sub-tables have no Approved check at all.** `save_consigner`,
`delete_consigner`, `save_export_cargo` and `delete_export_cargo` check only
`_billed_locked`. So today, anyone with edit permission can add, change or
delete parcels on an **Approved** VCN through the API. The screen hides it
(`isDraftRow(status)` gates the buttons client-side) but nothing on the server
does. This is the same shape as the incident: the UI implies a lock the backend
does not hold.

**(c) `isDraftRow` treats a blank status as Draft** —
`!docStatus || docStatus === 'Draft' || docStatus === 'Pending'`. Harmless while
every row has a status, but it means a NULL `doc_status` silently reads as
editable. Worth pinning down alongside the `created_by` NULLs.

### Plan

| # | Where | Change |
|---|---|---|
| 1b.1 | `VCN01/views.py::save` | Approved → 403 for everyone. Drop the approver-edits-in-place branch. |
| 1b.2 | `VCN01/views.py` | add a shared `_approved_locked(vcn_id)` guard beside `_billed_locked`; apply to `save_consigner`, `delete_consigner`, `save_export_cargo`, `delete_export_cargo`. |
| 1b.3 | `VCN01/views.py::send_to_expected` | already refuses Approved — leave as is. |
| 1b.4 | tests | Approved VCN → parcel save/delete 403; Draft VCN → 200. |

**Deliberate exception — Delays stay editable after approval.** The code says so
explicitly:

> *"Delays are logged whenever they're discovered — often after the VCN is
> closed/approved — so they're gated on permission only, never on the lock."*

`save_delay` / `delete_delay` keep their current behaviour. **Confirm this is
still wanted** — it is the one thing that stays writable on an Approved VCN.

**Effort: ~30 lines plus tests.** Worth doing with §1, since together they are
the whole "nothing edits an approved or billed vessel" rule.

---

## 2. Additional Quantity in BL

**Verdict: feasible, and smaller than expected — billing already supports it.**

### What already works
`FIN01.get_billables` does **not** bill the declared quantity once the vessel is
closed. It bills the actual logged quantity:

```
stage = 'actual' if ldud_status in ('Closed','Partial Close') else 'proforma'
qty   = actual if actual is not None else declared        # FIN01/model.py
```

where `actual` comes from `_actual_qty_map` — `SUM(lueu_parcel_log.quantity)`
excluding `is_deleted` and `is_shortclose`, apportioned pro-rata across merged
ops. **So a top-up row logged in LUEU01 already reaches billing with zero
changes to FIN01.**

### What blocks it — three upstream caps

**(a) The completion cap silently hides the excess.** `LUEU01.parcels_for_vessel`
walks log rows in order and stops counting once the target is reached:

```
if tgt > 0 and (a[0] + a[2]) >= tgt - 1e-6:
    continue          # parcel already complete — ignore this row
```

So today an over-logged row **is billed but is invisible on the LUEU01 screen**
— logged_qty, run hours and avg rate all exclude it. That divergence between
what billing charges and what the screen shows is the single most important
thing to fix, and it exists right now, before any new feature.

**(b) `LDUD01.save_parcel_op` refuses an op quantity above the VCN parcel
quantity** (`"Quantity … exceeds the available VCN parcel quantity"`). It caps
the op row, not log rows, but blocks raising the op to match.

**(c) Full Close becomes unreachable.** `get_closure_eligibility` computes
`can_full_close = abs((ops_total + shortclose_total) - bl_total) < 0.005`, and
`bl_total` comes from the VCN parcel quantities. Extra quantity makes
`ops_total > bl_total`, so the vessel can only ever Partial Close.

### Design — raise the target, do NOT insert a synthetic row

**Decided 2026-09-07: the VCN is never edited. The declared quantity stays as
declared; the amendment lives on the LDUD parcel op.**

An earlier draft proposed mirroring Short Close with a flagged synthetic log
row. That is the wrong shape here, because the two cases are opposites:

| | Short Close | Additional Quantity |
|---|---|---|
| Was the cargo handled? | **No** — leftover, written off | **Yes** — physically discharged |
| Is it already in the logbook? | No — nothing to log | **Yes** — logged hour by hour, then silently dropped by the completion cap |
| So the fix is… | invent a row to close the gap | **raise the target** so the real rows count |

Short Close invents a row because that quantity never existed. Additional
Quantity does not need one — the operator has already logged every tonne. The
only thing wrong is that the target is too low, so the cap discards the tail.

**Therefore: one amendment field on the parcel op, not a log row.**

Migration `ldud_parcel_ops`:
```
additional_qty     NUMERIC DEFAULT 0
additional_reason  TEXT
additional_by      TEXT
additional_date    TEXT
```

Effective target becomes:
```
sum(src_qty for parcel_ids) or op_qty      +  additional_qty
        ^ VCN declared, untouched              ^ LUEU01 amendment
```

Why this is the cheaper and safer option:
- **It fixes bug (a) for free.** Raising the target means the real logged rows
  stop hitting the completion cap — no separate fix needed.
- **Billing needs no change.** `_actual_qty_map` already sums the real log rows.
- **BPL01 and RP01 Berth_plan get it free** — both read `target_qty` from
  `LUEU01.get_started_parcels`.
- **Nothing is invented.** The logbook keeps only quantities that were actually
  handled, so avg flow rate and run hours stay honest with no exclusion rules.

### ⚠ The real risk: target resolution is already duplicated

`sum(src_qty…) or op_qty` is written out in **four** places today:

1. `LUEU01.get_started_parcels` (the per-parcel `targets` block)
2. `LUEU01._single_parcel_target` — docstring already says *"Mirrors the
   per-parcel target logic in get_started_parcels"*
3. `LDUD01.get_closure_eligibility` — the `bl_total` loop
4. `RP01/JJLTPL/jjltpl.py` — comment refers to *"the target-resolution block in
   get_started_parcels"*

Adding `additional_qty` to three of four is exactly how the Hodaka Galaxy
double-count happened: the same quantity computed differently in different
screens. **Centralise it first** — one helper in `LUEU01/model.py` that all four
call — then add the field in that one place.

### Touch points

| # | Where | Change |
|---|---|---|
| 1 | migration | 4 columns on `ldud_parcel_ops` |
| 2 | `LUEU01/model.py` | new shared target helper; repoint (1)–(4) above at it |
| 3 | `LUEU01/model.py` | `set_additional_qty()` / `clear_additional_qty()` |
| 4 | `LDUD01.get_closure_eligibility` | `bl_total` includes additional → Full Close reachable |
| 5 | `LDUD01.get_data` | `bl_quantities_display` includes additional (one extra query) |
| 6 | `LDUD01.save_parcel_op` | cap becomes declared + additional |
| 7 | `lueu01.html` | "Additional Qty" button + modal (qty + reason), and Revert |
| 8 | reports | BPL01 / RP01 Berth_plan free via `get_started_parcels`; check JJLTPL |

**Effort: ~1 day**, roughly half of it the centralisation in (2). No FIN01
change.

### Open question
Does raising the BL need proof or approval? It increases what the customer is
invoiced, and it sits beside a closure flow that already demands a Proof of
Quantity document and a password. Reason text is the minimum; proof upload is
the consistent option.

---

## 3. Remove Parcel

**Verdict: feasible, but the flag must go on the VCN parcel row, not on
`ldud_parcel_ops`. Flagging the op does the opposite of what is intended.**

### The trap

`FIN01.get_billables` selects billable parcels **straight from the VCN tables**:

```
FROM vcn_consigners c JOIN vcn_header h ...   WHERE c.importer_name = %s
UNION ALL
FROM vcn_export_cargo_declaration e JOIN vcn_header h ...
```

`ldud_parcel_ops` never filters that list — it only supplies the actual
quantity, via `_actual_qty_map`, whose docstring is explicit:

> *Parcels not covered by any op are absent from the map (caller falls back to
> the declared quantity).*

So if a parcel is flagged removed on `ldud_parcel_ops` and excluded from the
map, the parcel becomes *absent*, and billing falls back to **the full declared
quantity**. Flagging the op would bill the removed parcel at 100%, not 0%.

### Recommendation — flag the VCN parcel, drive it from LUEU01

**Decided 2026-09-07: agreed. The flag goes on the VCN parcel row and reports simply exclude flagged rows. The operator still never opens VCN01 — the LUEU01 action sets it.**

Add `is_removed BOOLEAN DEFAULT FALSE` + `removed_reason TEXT`, `removed_by`,
`removed_date` to **both** `vcn_consigners` and `vcn_export_cargo_declaration`,
and set it from a LUEU01 action. The operator never opens VCN01 — which is the
actual requirement — but the flag lands where every consumer already looks.

Consumers to filter (`WHERE is_removed IS NOT TRUE`):
- `FIN01.get_billables` — both legs of the UNION *(billing)*
- `FIN01.is_vcn_billed` / `_billed_vcn_ids` *(lock)*
- `VCN01.get_picker_parcels` *(LDUD parcel picker)*
- `VCN01.get_approval_eligibility` *(so a removed parcel is not "incomplete")*
- `VCN01._sync_header_cargo` *(header cargo list)*
- `LDUD01.get_closure_eligibility` → `bl_total` *(closure maths)*
- `LDUD01.get_data` → `bl_quantities_display` *(the grid figure)*
- `LUEU01.parcels_for_vessel` → `targets` *(Remaining)*

That list is long, but every one of those is a place the parcel is currently
counted — missing one means a removed parcel silently keeps counting somewhere.

### Rules to settle
1. **Removing a billed parcel must be refused.** `parcel_charge_billed` is keyed
   by `(VCN_IMPORT|VCN_EXPORT, parcel_id)`; once billed, removal needs a credit
   note, not a flag. Reuse the existing billed-lock check.
2. **Existing LUEU log rows against a removed parcel**: keep them (audit), but
   exclude from every total. A removed parcel with logged quantity is a
   contradiction worth warning about at the point of removal.
3. **Reversible?** Short Close has `revert_shortclose`. A Restore Parcel action
   is nearly free if the flag is nullable, and avoids an admin ticket for every
   mis-click.
4. **Who may remove?** Closure needs proof + password. Removing a parcel changes
   what the customer is billed, so it likely deserves at least the same.

**Effort: ~1–1.5 days**, most of it the consumer sweep rather than the flag.

---

## Recommended order

1. **Billed-reopen refusal + VCN edit lock (§1, §1b)** — one rule, ~50 lines
   plus tests. Closes the gap where the API still permits what the screen hides.
2. **Centralise target resolution** — prerequisite for (3); pure refactor, no
   behaviour change, and it retires the four-way duplication on its own merit.
3. **Additional Quantity (§2)** — the `additional_qty` field. Fixes the
   completion-cap divergence in the same change. No FIN01 change.
4. **Remove Parcel (§3)** — largest blast radius; do it last, with the consumer
   list above as the checklist.

All four are done. See the status note at the top.

Still outstanding from the Hodaka Galaxy plan: the Operation Type change
warning, mandatory Operation Type before detail entry, and blocking
send-to-Expected once Loading/Unloading entries exist.
