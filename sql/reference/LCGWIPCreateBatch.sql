-- Reference copy of Michael's LCGWIPCreateBatch from WipDb
-- Pulled 2026-04-08 by Josh Garrison (SA access)
-- This is the proc that populates WIPDetail with the initial job list + Vista data
-- DO NOT EXECUTE — reference only for understanding filtering and defaulting logic

-- Key findings:
-- 1. Job filter: Open/SoftClosed = all included; HardClosed = only if MonthClosed in WIP year
-- 2. Override columns default to Vista projected values when no bJCOR/bJCOP override exists
-- 3. Bonus profit is CALCULATED in SQL when no manual override
-- 4. Cost override has a floor: max(override_or_projected, actual_cost)
-- 5. Uses bJCCP for cost, bJCIP for billing (not bJCCD/vrvJBProgressBills)
-- 6. Final WHERE excludes jobs where ActualCost=0 AND contract amounts=0

-- See ARCHITECTURE.md for how this maps to our WIP_Vista_Query.sql
