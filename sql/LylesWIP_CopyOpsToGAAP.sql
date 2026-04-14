-- =============================================================================
-- STORED PROC: LylesWIPCopyOpsToGAAP
-- Copies Ops override values to GAAP override columns for all jobs in a
-- company/month where GAAP overrides are currently NULL.
-- Called from "Copy Ops to GAAP" button on Start sheet after OpsApproved.
--
-- Logic:
--   - Only copies where GAAPRevOverride IS NULL (won't overwrite existing GAAP edits)
--   - Sets GAAPRevPlugged/GAAPCostPlugged = 1 for copied values
--   - Returns count of rows updated
-- =============================================================================
IF OBJECT_ID('dbo.LylesWIPCopyOpsToGAAP', 'P') IS NOT NULL
    DROP PROCEDURE dbo.LylesWIPCopyOpsToGAAP;
GO
CREATE PROCEDURE dbo.LylesWIPCopyOpsToGAAP
    @JCCo       TINYINT,
    @WipMonth   DATE
AS
BEGIN
    SET NOCOUNT ON;

    UPDATE dbo.WipJobData
    SET GAAPRevOverride  = OpsRevOverride,
        GAAPRevPlugged   = OpsRevPlugged,
        GAAPCostOverride = OpsCostOverride,
        GAAPCostPlugged  = OpsCostPlugged,
        UpdatedAt        = GETDATE()
    WHERE JCCo = @JCCo
      AND WipMonth = @WipMonth
      AND GAAPRevOverride IS NULL
      AND GAAPCostOverride IS NULL
      AND (OpsRevOverride IS NOT NULL OR OpsCostOverride IS NOT NULL);

    SELECT @@ROWCOUNT AS RowsCopied;
END
GO
