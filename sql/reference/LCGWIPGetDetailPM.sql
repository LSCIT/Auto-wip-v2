-- Reference copy of Michael's LCGWIPGetDetailPM from WipDb
-- Pulled 2026-04-08 by Josh Garrison (SA access)
-- This is a simple SELECT from WIPDetail — no filtering logic here.
-- The filtering happens in LCGWIPCreateBatch when WIPDetail is populated.
-- Key: joins to PnpMain.pnp.WipPermissionsView3 for user-level permissions.

-- SELECT * FROM WIPDetail WD
-- JOIN #DeptList dl ON dl.Co = WD.JCCo AND dl.Dept = WD.Department
-- JOIN PnpMain.pnp.WipPermissionsView3 wp ON wp.JCCo = WD.JCCo AND wp.Job = WD.Contract
-- WHERE WD.JCCo = @Co AND WD.Month = @Month AND wp.UserId = @UserName
