# MARU – Quick Start

## Purpose
Get MARU running in under 5 minutes to retrieve email attachments from Outlook.

## Requirements
- Windows
- Outlook Desktop (signed in)
- PowerShell 5.1

## Launch
```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Maru_UI.ps1
```

## Basic Steps
1. Click **+ New** to create a profile
2. Enter Mailbox Name and Folder (e.g. Inbox)
3. Set **Days Back** or a date range
4. Choose **Save To Folders**
5. Click **Run**

## Output
- Attachments saved to disk
- CSV log written unless *No Log* is enabled

## Tips
- Start with Days Back before adding filters
- Use Skip Already Downloaded to avoid duplicates
- Check MARU_Log.csv for results
