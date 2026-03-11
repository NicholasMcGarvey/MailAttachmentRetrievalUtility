# MARU – Audit & Design Appendix

## Design Decisions
- Windows + Outlook COM for mailbox access
- PowerShell 5.1 for enterprise compatibility
- Detached worker process for UI stability

## Known Limitations
- Requires Outlook desktop
- Not supported on macOS/Linux
- Performance depends on mailbox size

## Data Handling
- Logs contain email addresses and subjects
- Attachments may contain sensitive data

## Update Controls
- Version manifest with SHA-256 checksums
- Automatic backup of replaced files
