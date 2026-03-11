# MARU – Support Runbook

## Common Issues

### Script will not run
**Cause:** PowerShell execution policy
**Fix:** Launch with -ExecutionPolicy Bypass

### No attachments saved
**Cause:** Skip Already Downloaded enabled or filters too narrow
**Fix:** Disable skip option or widen date range

### Mailbox not found
**Cause:** Incorrect mailbox name or permissions
**Fix:** Verify mailbox access in Outlook

## Diagnostics
- Primary log: MARU_Log.csv
- Status values: Downloaded, Skipped
