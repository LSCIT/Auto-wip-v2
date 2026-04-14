# WIP Column Mapping: Stored Proc Fields → Sheet Columns → Vista Source

## LCGWIPGetDetailPM Recordset Fields → Sheet Named Ranges

### Core Job Identification
| Proc Field | Sheet Column Name | Sheet | Vista Source |
|-----------|------------------|-------|-------------|
| Contract | COLJobNumber | Both | bJCJM.Contract (RTRIM) |
| ContractDescription | COLJobDesc | Both | bJCCM.Description |
| PM | COLPrjMngr | Both | bJCMP.Name (via bJCJM.ProjectMgr) |
| Department | (grouping only) | Both | bJCCM.Department → bJCDM.Description |
| ContractStatus | COLZContractStatus | Both | bJCJM.JobStatus |
| CompletionDate | COLCompDate | Both | bJCCM.ActualCloseDate |

### Financial - From Viewpoint
| Proc Field | Sheet Column Name | Sheet | Vista Source |
|-----------|------------------|-------|-------------|
| OrigContractAmt | (part of COLCurConAmt calc) | Both | bJCCM.OrigContractAmt |
| COContractAmt | (part of COLCurConAmt calc) | Both | SUM(bJCID.ContractAmt WHERE JCTransType='CO') |
| **COLCurConAmt** | **=OrigContractAmt + COContractAmt** | Both | **Calculated in VBA** |
| ProjContract | COLPMProjRev | Both | bJCCM.ContractAmt (or projected from bJCOR) |
| ProjCost | COLPMProjCost | Both | SUM(bJCCD.ProjCost WHERE JCTransType='PF' AND CostType=99) |
| ActualCost | COLJTDCost | Both | SUM(bJCCD.ActualCost) with PR/EM date logic |
| CYActualCost | COLCYCost | Both | SUM(bJCCD.ActualCost WHERE date >= YrStart) |
| BilledAmt | COLBILLBillings | Both | bJCCM.BilledAmt |

### Override Fields (from WipDb — Phase 1: leave blank/zero)
| Proc Field | Sheet Column Name | Sheet | Notes |
|-----------|------------------|-------|-------|
| OpsRev | COLOvrRevProj (if OpsRevPlugged="Y") | Ops | WipDb override, Phase 2 |
| OpsCost | COLOvrCostProj (if OpsCostPlugged="Y") | Ops | WipDb override, Phase 2 |
| GAAPRev | COLOvrRevProj (if GAAPRevPlugged="Y") | GAAP | WipDb override, Phase 2 |
| GAAPCost | COLOvrCostProj (if GAAPCostPlugged="Y") | GAAP | WipDb override, Phase 2 |
| BonusProfit | COLJTDBonusProfit (if BonusProfitPlugged="Y") | Ops | WipDb override, Phase 2 |
| BonusProfitNotes | COLJTDBonusProfitNotes | Ops | WipDb override, Phase 2 |
| PriorYrBonusProfit | COLAPYBonusProfit | Ops | WipDb override, Phase 2 |

### Status/Approval Fields (from WipDb — Phase 1: leave blank)
| Proc Field | Sheet Column Name | Sheet | Notes |
|-----------|------------------|-------|-------|
| Close | COLClose | Ops | "C" if closed by user, Phase 2 |
| Completed | COLDone | Ops | "Y" if approved, Phase 2 |
| CompletedGAAP | COLGAAPDone | GAAP | "Y" if approved, Phase 2 |
| UserName | COLZUserName | Both | Approval user, Phase 2 |
| BatchSeq | COLZBatchSeq | Both | Batch tracking, Phase 2 |
| RowVersion | COLZRowVersion | Both | Concurrency, Phase 2 |

### Historical Data (from WipDb — Phase 1: leave blank)
| Proc Field | Sheet Column Name | Notes |
|-----------|------------------|-------|
| OrgActualCost | COLZORGJTDCost | Original JTD cost before override |
| OrgCYActualCost | COLZORGCYCost | Original CY cost before override |
| OrgBilledAmt | COLZORGBilledAmt | Original billed amt before override |
| OrgCYBilledAmt | COLZORGCYBilledAmt | Original CY billed before override |
| OpsRev | COLZOPsRev | Stored ops revenue |
| OpsCost | COLZOPsCost | Stored ops cost |
| GAAPRev | COLZGAAPRev | Stored GAAP revenue |
| GAAPCost | COLZGAAPCost | Stored GAAP cost |
| OpsRevNotes | COLZOPsRevNotes | Notes |
| OpsCostNotes | COLZOPsCostNotes | Notes |
| GAAPRevNotes | COLZGAAPRevNotes | Notes |
| GAAPCostNotes | COLZGAAPCostNotes | Notes |

### Prior Year Values (from WipDb — stored at year-end)
| Proc Field | Sheet Column Name | Notes |
|-----------|------------------|-------|
| PriorYearGAAPRev | COLZGAAPPYRev | Year-end GAAP projected revenue |
| PriorYearGAAPCost | COLZGAAPPYCost | Year-end GAAP projected cost |
| PriorYearOpsRev | COLZOpsPYRev | Year-end Ops projected revenue |
| PriorYearOpsCost | COLZOpsPYCost | Year-end Ops projected cost |

### 6-Month Trend Data (from WipDb — shown as comments)
| Proc Fields | Used For | Notes |
|------------|----------|-------|
| LastProjContract1-6 | COLPMProjRev comment tooltip | Revenue trend |
| LastProjCost1-6 | COLPMProjCost comment tooltip | Cost trend |
| LastOpsRev1-6 + Plugged flags | COLOvrRevProj comment tooltip | Ops override trend |
| LastOpsCost1-6 + Plugged flags | COLOvrCostProj comment tooltip | Ops cost override trend |
| LastGAAPRev1-6 + Plugged flags | COLOvrRevProj comment tooltip (GAAP) | GAAP override trend |
| LastGAAPCost1-6 + Plugged flags | COLOvrCostProj comment tooltip (GAAP) | GAAP cost override trend |
| LastBonusProfit | COLZPriorBonusProfit | Prior month bonus |
| LastActualCost | (used in calculations) | Prior month JTD cost |

## VBA-Side Calculations (computed from proc data, NOT returned by proc)

### Prior Year Calcs
```
PYCost = ActualCost - CYActualCost
PYPctComp = PYCost / PriorYearGAAPCost  (GAAP sheet)
          = PYCost / PriorYearOpsCost   (Ops sheet)
PYEarnedRev = PriorYearGAAPRev * PYPctComp

COLAPYPctComp = PYPctComp
COLAPYRev     = PYEarnedRev              (GAAP sheet)
              = PYCost + PriorYrBonusProfit (Ops sheet)
COLAPYCalcProfit = PYEarnedRev - PYCost   (Ops only)
COLAPYCost    = PYCost
```

### Prior GAAP Projected Profit (the "March WIP plug" lookup)
```
LastGAAPProjContract = first "P"-plugged value from LastGAAPRev1-6
LastGAAPProjCost = first "P"-plugged value from LastGAAPCost1-6
LastGAAPPctComp = LastActualCost / LastGAAPProjCost
LastGAAPEarnedRev = LastGAAPProjContract * LastGAAPPctComp
COLZPriorJTDGAAPProfit = LastGAAPProjContract - LastGAAPProjCost
```

### Prior Ops Projected Profit
```
COLZPriorOPsProfit = LastOpsRev - LastOpsCost
COLZPriorJTDOPsProfit = LastOpsRev - LastOpsCost
```

## Phase 1 Strategy: Which fields come from Vista vs. set to default

### FROM VISTA (critical for validation):
- Contract, ContractDescription, PM, Department, ContractStatus, CompletionDate
- OrigContractAmt, COContractAmt → COLCurConAmt
- ProjContract (ContractAmt) → COLPMProjRev
- ProjCost (from bJCCD PF transactions) → COLPMProjCost
- ActualCost (from bJCCD actual transactions) → COLJTDCost
- CYActualCost (from bJCCD, current year only) → COLCYCost
- BilledAmt (from bJCCM) → COLBILLBillings

### SET TO DEFAULT (override/historical data, Phase 2):
- All override fields (Ops/GAAP Rev/Cost + plugged flags) → 0 / ""
- BonusProfit, BonusProfitNotes → 0 / ""
- Approval/status fields (Close, Completed, CompletedGAAP) → ""
- BatchSeq, RowVersion → 0 / ""
- 6-month trend data → 0 (no tooltip comments)
- Prior year stored values → 0 (these need year-end snapshot logic)

### CALCULATED IN NEW VBA (WIP business rules):
- % Complete = JTDCost / ProjCost (or CurrentEstimate if ProjCost=0)
- Earned Revenue: if %Complete < 0.10 → JTDCost; else → ContractAmt × %Complete
- Projected Profit = ContractAmt - ProjCost
- Remaining Cost = ProjCost - JTDCost
- Prior Projected Profit (Col R) — needs March baseline from JCCP/budProjInfo
- GAAP quarterly rule (only Mar/Jun/Sep/Dec)
- All downstream cascading values (Prev Yr Revenue, Curr Yr Revenue, etc.)
