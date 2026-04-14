# WIP Meeting Questions
**Meeting:** Wednesday, March 25, 2026 at 9:00 AM PDT
**Attendees:** Nicole Leasure, Cindy Jordan (+ Josh)
**Purpose:** Resolve open design questions before building the write/save path

---

## 1. Prior Projected Profit (Col R on GAAP / Col Q on Ops)

**Question for Nicole:**
When you start a new WIP month, where does the Prior Projected Profit column get its starting value?

- Does it carry over automatically from last month's run?
- Is it something you manually copy or enter at the start of each cycle?
- Or does it come from somewhere in Vista/Viewpoint?

**Why it matters:** This determines whether we store it in our new database and carry it forward month-to-month, or whether we re-derive it from Vista each time.

---

## 2. JV Workflow *(lower priority — ask only if time permits)*

**Question for Nicole:**
We're deferring JV tab functionality until the Jobs workflow is validated. When we get there, can you let us know how the JV review cycle works and whether it was ever fully functional in the original tool?

---

## 3. Editable Columns on Jobs-Ops (Yellow Columns)

**Question for Ops / Nicole:**
We want to make sure we only allow editing on the right columns in the Jobs-Ops tab. Can you confirm which columns Ops should be able to change?

We believe it is the override revenue, override cost, bonus profit, completion date, and notes columns — but we want to confirm the exact set before we lock it down.

---

## 4. Editable Columns on Jobs-GAAP (Yellow Columns)

**Question for Cindy / Nicole:**
Same question for the Jobs-GAAP tab — which columns should Accounting be able to edit during their final review (Stage 3)?

---

## 5. Multi-Level Ops Approval — How Does It Actually Work?

**Question for Nicole:**
In the column mapping meeting you mentioned the approval chain goes: PMs first → Tonia/Grant review → final sign-off. How does that work in practice with the workbook?

- Do multiple people open the same workbook file and take turns?
- Does one person collect everyone's input before marking Ops Final Approval?
- Are there sub-stages within Stage 2 we need to support, or does it all resolve to a single "Ops Final Approval: Yes" click?

**Why it matters:** If there are sub-stages (e.g., PM edits first, then DM reviews), we need to design the locking and role behavior accordingly. If it's one person consolidating input then clicking approve, we can keep it simple.

---

## 6. Carry-Forward Overrides to the Next Month

**Question for Nicole:**
When you open a new month's WIP cycle, do you expect to see last month's override projections (revenue, cost) pre-loaded as a starting point — or do you prefer to start from the Vista-calculated values each time and re-enter overrides fresh?

**Why it matters:** We can either carry forward prior month overrides automatically or always start clean from Vista. Carrying forward is less re-entry work; starting clean is simpler and avoids stale data.

---

## 7. Completion Date — Should It Update Vista?

**Question for Cindy / Nicole:**
When someone enters or changes a job's Completion Date in the WIP tool, should that value write back to the corresponding field in Vista/Viewpoint — or should it stay local to the WIP tool only?

---

## 8. Production Server — Database Name Confirmation

**Question for Nicole or IT:**
We have the production Vista server address (`10.112.11.8`). What is the database name on that server we should be connecting to?

---

## 9. Division / Department Scope

**Question for Nicole:**
When you run the WIP for a given month, do you always run all divisions at once, or do different people run different divisions separately?

**Why it matters:** The batch system creates one batch per division. If multiple people run different divisions simultaneously, we need to make sure batches don't conflict.

---

## Notes / Answers
*(fill in during meeting)*

| # | Answer | Notes |
|---|--------|-------|
| 1 | | |
| 2 | | |
| 3 | | |
| 4 | | |
| 5 | | |
| 6 | | |
| 7 | | |
| 8 | | |
| 9 | | |
