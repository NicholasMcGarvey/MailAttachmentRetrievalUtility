# MailAttachmentRetrievalUtility
Internal Windows-based Outlook utility for bulk retrieval of email attachments using configurable filters, logging, saved profiles, and a versioned self-update mechanism.

# MARU – Mail Attachment Retrieval Utility

**Internal use only.**

MARU is a Windows-based Outlook utility for bulk retrieval of email attachments using configurable filters, logging, saved profiles, and a versioned self-update mechanism.

## Features
- Outlook mailbox and folder scanning
- Flexible filters (subject, sender, recipient, date)
- Bulk attachment export
- CSV logging for audit and re-run safety
- GUI with saved profiles
- Checksum-validated self-updating

## Requirements
- Windows
- Outlook Desktop (MAPI)
- PowerShell 5.1

## Getting Started
```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Maru_UI.ps1
```

## Repository Structure
```
MARU/
├── MARU.ps1
├── Maru_UI.ps1
├── MARU_Update.ps1
├── MARU_Publish.ps1
├── MARU_Configs.json
├── MARU_Log.csv
└── MARU_Version.json
```

## Support
See SUPPORT.md for troubleshooting and escalation guidance.

## Audit & Risk
See AUDIT.md for design decisions, limitations, and compliance notes.
