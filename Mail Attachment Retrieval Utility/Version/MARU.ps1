[CmdletBinding()]
param(
    [string[]]$FilterSubject,
    [string[]]$FilterSender,
    [string[]]$FilterTo,
    [string[]]$FilterCC,
    [string[]]$FilterBCC,
    [string[]]$FilterToOrCC,
    [datetime]$FromDate,
    [datetime]$ToDate,
    [int]$DaysBack,

    [string]$MailBoxName,
    [string]$MailBoxFolderName = "Inbox",
    [switch]$SearchSubFolders,
    [string[]]$SaveToFolders = @(Join-Path -Path $env:TEMP -ChildPath "Attachments"),

    # Logging parameters
    [string]$LogPath,
    [bool]$SkipAlreadyDownloaded = $true,
    [switch]$NoLog,

    # Folder resolution
    [ValidateSet("First","Last")]
    [string]$CreateFolderPreference = "Last",
    [ValidateSet("Suffix","Overwrite","Skip","Error")]
    [string]$FileCollisionAction = "Suffix",

    [string]$RelativeBase = $PSScriptRoot
)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


function Quote-Csv([string]$s) {
    if ($null -eq $s) { $s = '' }
    '"' + ($s -replace '"','""') + '"'
}

# Normalize text to avoid CR/LF breaking rows
function _NormText([string]$s) {
    if ($null -eq $s) { return '' }
    return ($s -replace '[\r\n]+',' ')
}


function Convert-ToSafeFileName {
    param(
        [Parameter(Mandatory)][string]$Name,
        [string]$Replacement = "_"
    )
    # Replace characters invalid on Windows file systems
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $Name.ToCharArray()) {
        if ($invalid -contains $ch) { [void]$sb.Append($Replacement) } else { [void]$sb.Append($ch) }
    }
    # Collapse whitespace/newlines that sometimes come through display UIs or forwarded items
    $safe = $sb.ToString() -replace '[\r\n]+',' ' -replace '\s{2,}',' '
    return $safe.Trim()
}

function Join-SafePath {
    param(
        [Parameter(Mandatory)][string]$Directory,
        [Parameter(Mandatory)][string]$FileName
    )
    # Trim, remove newlines, and get absolute path
    $dir = ($Directory -replace '[\r\n]+',' ').Trim()
    $file = Convert-ToSafeFileName -Name $FileName
    $full = [System.IO.Path]::Combine($dir, $file)
    return ([System.IO.Path]::GetFullPath($full))
}

function Normalize-PathForSave {
    param([Parameter(Mandatory)][string]$Path)
    # Remove CR/LF and any leading/trailing whitespace that could be parsed as new pipeline parts
    $p = ($Path -replace '[\r\n]+',' ').Trim()
    # DO NOT unquote here; pass clean literal to COM SaveAsFile
    return $p
}

function Resolve-InputPath {
    param(
        [string]$Path,
        [switch]$MustExist,
        [switch]$EnsureExists
    )
    $Path = $Path.Trim()
    if (-not [System.IO.Path]::IsPathRooted($Path)) {
        $Path = Join-Path $RelativeBase $Path
    }
    $Path = [System.IO.Path]::GetFullPath($Path)
    if ($MustExist -and -not (Test-Path $Path)) {
        throw "Input path does not exist: $Path"
    }
    if ($EnsureExists) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
    return $Path
}

function Get-LogFilePath {
    param([string]$BasePath)
    $resPath = Resolve-InputPath -Path $BasePath
    return Join-Path -Path $resPath -ChildPath "MARU_Log.csv"
}

<#
function Initialize-Log {
    param([string]$LogFile)
    #$headers = "Timestamp,MailBox,Folder,EmailSubject,ReceivedDate,AttachmentFileName,AttachmentSize,SavedTo,Status"
    $headers = "Timestamp,Status,MailBox,Folder,EmailSubject,ReceivedDate,AttachmentFileName,AttachmentSize,SavedTo,FromName,FromEmail,ToEmails,CcEmails"

    if (-not (Test-Path $LogFile)) {
        $headers | Out-File -FilePath $LogFile -Encoding utf8
        Write-Verbose "Log created: $LogFile"
    }
}#>

function Initialize-Log {
    param([string]$LogFile)

    $headers = 'Timestamp,Status,MailBox,Folder,EmailSubject,ReceivedDate,AttachmentFileName,AttachmentSize,SavedTo,FromName,FromEmail,ToEmails,CcEmails'
    if (-not (Test-Path $LogFile)) {
        $headers | Out-File -FilePath $LogFile -Encoding utf8
    }
}

function New-LogKey {
    param(
        [string]$MailBox,
        [string]$Folder,
        [string]$ReceivedDate,
        [string]$FileName,
        [string]$FileSize
    )
    return "$MailBox|$Folder|$ReceivedDate|$FileName|$FileSize"
}

function Resolve-SaveFolder {
    [OutputType([string])]
    param(
        [string[]]$Folders,
        [ValidateSet("First","Last")][string]$CreatePreference = "Last"
    )
    if ($Folders.Count -eq 1) { $Folders = $Folders.Split('|') }
    foreach ($folder in $Folders) {
        $rslvFld = Resolve-InputPath -Path $folder.Trim()
        if (Test-Path $rslvFld) {
            Write-Verbose "Resolved save folder: '$rslvFld'"
            return $rslvFld
        }
    }
    $folderToCreate = if ($CreatePreference -eq "First") { $Folders[0] } else { $Folders[-1] }
    $folderToCreate = Resolve-InputPath $folderToCreate.Trim()
    Write-Warning "No folders in SaveToFolders exist. Creating: '$folderToCreate'"
    New-Item -ItemType Directory -Path $folderToCreate -Force | Out-Null
    return $folderToCreate
}

# ---------------------------------------------------------------------------
# Build-LogLine: formats a single CSV log line without writing to disk.
# Accumulate lines in a StringBuilder and flush with Flush-LogBuffer
# after each folder completes to reduce per-item I/O overhead.
# ---------------------------------------------------------------------------
<#
function Build-LogLine {
    param(
        [string]$MailBox,
        [string]$Folder,
        [string]$EmailSubject,
        [string]$ReceivedDate,
        [string]$FileName,
        [long]$FileSize,
        [string]$SavedTo,
        
        [string]$FromName,
        [string]$FromEmail,
        [string]$ToEmails,
        [string]$CcEmails,

        [ValidateSet("Downloaded","Skipped")][string]$Status

    )
    $safeSubject  = '"' + ($EmailSubject -replace '"','""') + '"'
    $safeMailBox  = '"' + ($MailBox      -replace '"','""') + '"'
    $safeFolder   = '"' + ($Folder       -replace '"','""') + '"'
    $safeSavedTo  = '"' + ($SavedTo      -replace '"','""') + '"'
    $safeFileName = '"' + ($FileName     -replace '"','""') + '"'
    #return "$([datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')),$safeMailBox,$Status,$safeFolder,$safeSubject,$ReceivedDate,$safeFileName,$FileSize,$safeSavedTo,$(& $q $FromName),$(& $q $FromEmail),$(& $q $ToEmails),$(& $q $CcEmails)"
    return "$([datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')),$safeMailBox,$Status,$safeFolder,$safeSubject,$ReceivedDate,$safeFileName,$FileSize,$safeSavedTo, $(quote-csv $FromName),$(quote-csv $FromEmail),$(quote-csv $ToEmails),$(quote-csv $CcEmails)"
}
#>

# CSV quoting helper: wrap in quotes and double any embedded quotes
function Quote-Csv([string]$s) {
    if ($null -eq $s) { $s = '' }
    '"' + ($s -replace '"','""') + '"'
}

# Normalize text to avoid CR/LF breaking rows
function _NormText([string]$s) {
    if ($null -eq $s) { return '' }
    return ($s -replace '[\r\n]+',' ')
}

function Build-LogLine {
    param(
        [string]$MailBox,
        [string]$Folder,
        [string]$EmailSubject,
        [string]$ReceivedDate,
        [string]$FromName,
        [string]$FromEmail,
        [string]$ToEmails,
        [string]$CcEmails,
        [string]$FileName,
        [long]  $FileSize,
        [string]$SavedTo,
        [ValidateSet('Downloaded','Skipped')][string]$Status
    )

    $fields = @(
        [datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')   # timestamp (safe as plain)
        (Quote-Csv (_NormText $Status))                   # quote status too for consistency
        (Quote-Csv (_NormText $MailBox))
        (Quote-Csv (_NormText $Folder))
        (Quote-Csv (_NormText $EmailSubject))
        (_NormText $ReceivedDate)                         # already formatted; leave plain or quote if you prefer
        (Quote-Csv (_NormText $FileName))
        ($FileSize)                                       # numeric
        (Quote-Csv (_NormText $SavedTo))
        (Quote-Csv (_NormText $FromName))
        (Quote-Csv (_NormText $FromEmail))
        (Quote-Csv (_NormText $ToEmails))
        (Quote-Csv (_NormText $CcEmails))
        
                           
    )

    return ($fields -join ',')
}



function Flush-LogBuffer {
    param(
        [System.Text.StringBuilder]$Buffer,
        [string]$LogFile
    )
    if ($Buffer.Length -eq 0) { return }
    $Buffer.ToString() | Out-File -FilePath (Resolve-InputPath $LogFile) -Append -Encoding utf8
    [void]$Buffer.Clear()
}

function Build-DaslOrClause {
    param([string]$Field, [string[]]$Values)
    if (-not $Values -or $Values.Count -eq 0) { return $null }
    $clauses = $Values | ForEach-Object { """$Field"" LIKE '%$_%" }
    return "(" + ($clauses -join " OR ") + ")"
}

function Test-MatchesAny {
    param([string]$Value, [string[]]$Filters)
    if (-not $Filters -or $Filters.Count -eq 0) { return $true }
    foreach ($f in $Filters) {
        if ($Value -match [regex]::Escape($f)) { return $true }
    }
    return $false
}

# ---------------------------------------------------------------------------
# Safe ContentId accessor - Outlook throws a COM exception when the
# 0x3712001E property is absent on some attachment types. Treat missing
# as $null so the attachment is still evaluated rather than the whole
# email being silently skipped.
# ---------------------------------------------------------------------------
function Get-AttachmentContentId {
    param([object]$Attachment)
    try {
        return $Attachment.PropertyAccessor.GetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x3712001E")
    } catch {
        return $null
    }
}

# ---------------------------------------------------------------------------
# Returns a flat list of @{Folder=<COM>; Path=<string>} starting from root.
# Recurses all levels deep when $Recurse is $true.
# ---------------------------------------------------------------------------
function Get-MailFolders {
    param(
        [object]$RootFolder,
        [string]$RootPath,
        [bool]$Recurse
    )
    $results = [System.Collections.Generic.List[PSCustomObject]]::new()
    $results.Add([PSCustomObject]@{ Folder = $RootFolder; Path = $RootPath })
    if ($Recurse) {
        foreach ($sub in $RootFolder.Folders) {
            foreach ($item in (Get-MailFolders -RootFolder $sub -RootPath "$RootPath/$($sub.Name)" -Recurse $true)) {
                $results.Add($item)
            }
        }
    }
    return $results
}

# ---------------------------------------------------------------------------
# Generate To and CC recipient lists with emails -- instead of Display name
# ---------------------------------------------------------------------------

# Returns an array of recipient strings for the given MailItem and type (1=To, 2=CC)
function Get-RecipientStrings {
    param(
        [Parameter(Mandatory)][object]$MailItem,
        [ValidateSet(1,2)][int]$Type,  # 1=To, 2=CC
        [switch]$EmailOnly
    )

    $results = New-Object System.Collections.Generic.List[string]
    foreach ($recip in $MailItem.Recipients) {
        # Recipient.Type: 1=To, 2=CC, 3=Bcc
        if ($recip.Type -ne $Type) { continue }

        $smtp = $null
        try { $smtp = $recip.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x39FE001E') } catch {}
        if ($EmailOnly) {
            if (-not [string]::IsNullOrWhiteSpace($smtp)) { [void]$results.Add($smtp) }
            continue
        }

        $text = $smtp

        # Fallbacks: Address, then Name (display)
        if ([string]::IsNullOrWhiteSpace($text)) { $text = $recip.Address }
        if ([string]::IsNullOrWhiteSpace($text)) { $text = $recip.Name }

        if (-not [string]::IsNullOrWhiteSpace($text)) {
            [void]$results.Add($text)
        }
    }
    return ,$results.ToArray()
}

function Test-AnyMatch {
    param(
        [string[]]$Haystack,    # e.g., array of recipient strings
        [string[]]$Needles      # your filter values
    )
    if (-not $Needles -or $Needles.Count -eq 0) { return $true }
    foreach ($n in $Needles) {
        # Escape as literal fragment; treat as substring/regex match
        $pattern = [regex]::Escape($n)
        foreach ($h in $Haystack) {
            if ($h -match $pattern) { return $true }
        }
    }
    return $false
}

# ---------------------------------------------------------------------------
# Resolve save folder & log
# ---------------------------------------------------------------------------

[string]$resolvedSaveFolder = Resolve-SaveFolder -Folders $SaveToFolders -CreatePreference $CreateFolderPreference
Write-Verbose "Save folder: '$resolvedSaveFolder'"

$resolvedLogPath = if ($LogPath) {
    $LogPath = Resolve-InputPath -Path $LogPath.Trim() -EnsureExists
    Get-LogFilePath -BasePath $LogPath
} else {
    Get-LogFilePath -BasePath $resolvedSaveFolder
}
Write-Verbose "Log file: $resolvedLogPath"

if (-not $NoLog) { Initialize-Log -LogFile $resolvedLogPath }

$downloadedSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
if (-not $NoLog -and $SkipAlreadyDownloaded -and (Test-Path $resolvedLogPath)) {
    foreach ($row in (Import-Csv -Path $resolvedLogPath)) {
        if ($row.Status -eq "Downloaded") {
            $key = "$($row.MailBox)|$($row.Folder)|$($row.ReceivedDate)|$($row.AttachmentFileName)|$($row.AttachmentSize)"
            [void]$downloadedSet.Add($key)
        }
    }
    Write-Verbose "Loaded $($downloadedSet.Count) previously downloaded entries from log."
}

# Log write buffer - flushed once per folder to reduce per-item disk I/O
$logBuffer = [System.Text.StringBuilder]::new()

# ---------------------------------------------------------------------------
# Resolve date range
# ---------------------------------------------------------------------------

$now               = Get-Date
$effectiveFromDate = $null
$effectiveToDate   = $null
$useDateFilter     = $false

$hasFrom = $PSBoundParameters.ContainsKey('FromDate')
$hasTo   = $PSBoundParameters.ContainsKey('ToDate')
$hasDays = $PSBoundParameters.ContainsKey('DaysBack')

if ($hasFrom -and $hasTo) {
    if ($hasDays) { Write-Warning "Both -FromDate/-ToDate and -DaysBack supplied. Explicit range takes priority." }
    $effectiveFromDate = $FromDate; $effectiveToDate = $ToDate; $useDateFilter = $true
} elseif ($hasFrom) {
    $effectiveFromDate = $FromDate; $effectiveToDate = $now; $useDateFilter = $true
} elseif ($hasTo -and $hasDays) {
    $effectiveToDate = $ToDate; $effectiveFromDate = $ToDate.AddDays(-$DaysBack); $useDateFilter = $true
} elseif ($hasDays) {
    $effectiveFromDate = $now.AddDays(-$DaysBack); $effectiveToDate = $now; $useDateFilter = $true
}

if ($useDateFilter) {
    Write-Verbose "Date filter: $($effectiveFromDate.ToString('yyyy-MM-dd HH:mm:ss')) to $($effectiveToDate.ToString('yyyy-MM-dd HH:mm:ss'))"
}

# ---------------------------------------------------------------------------
# Outlook connection & mailbox validation
# ---------------------------------------------------------------------------

$Outlook   = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")

# Validate mailbox - list available on failure to aid diagnosis
$SharedMailbox = $Namespace.Folders | Where-Object { $_.Name -eq $MailBoxName } | Select-Object -First 1
if ($null -eq $SharedMailbox) {
    $available = ($Namespace.Folders | ForEach-Object { "  - $($_.Name)" }) -join "`n"
    throw "Mailbox not found: '$MailBoxName'`nAvailable mailboxes:`n$available"
}

# Resolve nested folder path - delimiter is /
# e.g. "Inbox/Project A/2025" traverses three levels deep
$folderParts   = $MailBoxFolderName -split '/' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
$MailBoxFolder = $SharedMailbox

foreach ($part in $folderParts) {
    $next = $MailBoxFolder.Folders | Where-Object { $_.Name -eq $part } | Select-Object -First 1
    if ($null -eq $next) {
        $available = ($MailBoxFolder.Folders | ForEach-Object { "  - $($_.Name)" }) -join "`n"
        throw "Folder '$part' not found under '$($MailBoxFolder.Name)'`nAvailable folders:`n$available"
    }
    $MailBoxFolder = $next
}

# ---------------------------------------------------------------------------
# Resolve folders to search
# ---------------------------------------------------------------------------

$foldersToSearch = Get-MailFolders -RootFolder $MailBoxFolder -RootPath $MailBoxFolderName -Recurse $SearchSubFolders.IsPresent
Write-Verbose "Folders to search: $($foldersToSearch.Count)"

# ---------------------------------------------------------------------------
# Process - once per folder
# ---------------------------------------------------------------------------

$downloadCount = 0
$skippedCount  = 0

foreach ($folderEntry in $foldersToSearch) {

    $currentFolder     = $folderEntry.Folder
    $currentFolderPath = $folderEntry.Path
    Write-Verbose "Processing folder: '$currentFolderPath'"

    # Build DASL filter for this folder
    $daslClauses = [System.Collections.Generic.List[string]]::new()
    $daslClauses.Add('"urn:schemas:httpmail:hasattachment" = 1')


    if ($FilterSubject -and $FilterSubject.Count -gt 0) {
        $FilteredItems = $FilteredItems | Where-Object {
            $subj = $_.Subject
            #$daslClauses.add("")
            $daslClauses.Add('"urn:schemas:httpmail:subject" like ' + "'%$FilterSubject%'")
            #$filter = "@SQL=""http://schemas.microsoft.com/mapi/proptag/0x0037001F"" CI_PHRASE 'Meeting Notes'"
            #$FilterSubject | Where-Object { $subj -match [regex]::Escape($_) } | Select-Object -First 1
        }
    }

    if ($useDateFilter) {
        #$daslFrom = $effectiveFromDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        #$daslTo   = $effectiveToDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        $daslFrom = $effectiveFromDate.ToUniversalTime().ToString("MM/dd/yyyy hh:mm tt")
        $daslTo   = $effectiveToDate.ToUniversalTime().ToString("MM/dd/yyyy hh:mm tt")

        $daslClauses.Add("""urn:schemas:httpmail:datereceived"" >= '$daslFrom'")
        $daslClauses.Add("""urn:schemas:httpmail:datereceived"" <= '$daslTo'")
    }

    # Guard each sender clause independently - only join non-null pieces so
    # a null component never produces a syntactically invalid DASL string.
    if ($FilterSender -and $FilterSender.Count -gt 0) {
        $nameClause   = Build-DaslOrClause -Field "urn:schemas:httpmail:fromname"  -Values $FilterSender
        $emailClause  = Build-DaslOrClause -Field "urn:schemas:httpmail:fromemail" -Values $FilterSender
        $senderParts  = @($nameClause, $emailClause) | Where-Object { $_ }
        if ($senderParts.Count -gt 0) {
            $daslClauses.Add("(" + ($senderParts -join " OR ") + ")")
        }
    }

    $daslFilter  = "@SQL=" + ($daslClauses -join " AND ")
    $daslApplied = $false
    $currentFolder.Items.Sort("[ReceivedTime]")
    $currentFolder.Items.IncludeRecurrences = $false

    try {
        $FilteredItems = $currentFolder.Items.Restrict($daslFilter)
        $daslApplied   = $true
        Write-Verbose "DASL filter applied: $daslFilter"
    } catch {
        Write-Warning "Outlook .Restrict() failed ($_). Falling back to PowerShell filtering."
    }

    if (-not $daslApplied) {
        $FilteredItems = $currentFolder.Items
        if ($useDateFilter) {
            $FilteredItems = $FilteredItems | Where-Object {
                $_.ReceivedTime -ge $effectiveFromDate -and $_.ReceivedTime -le $effectiveToDate
            }
        }
        $FilteredItems = $FilteredItems | Where-Object { $_.Attachments.Count -gt 0 }
        if ($FilterSender -and $FilterSender.Count -gt 0) {
            $FilteredItems = $FilteredItems | Where-Object {
                (Test-MatchesAny $_.SenderName         $FilterSender) -or
                (Test-MatchesAny $_.SenderEmailAddress $FilterSender)
            }
        }
    }




    # PS-side filters - applied after DASL or fallback
    <#
    if ($FilterSubject -and $FilterSubject.Count -gt 0) {
        $FilteredItems = $FilteredItems | Where-Object {
            $subj = $_.Subject
            $FilterSubject | Where-Object { $subj -match [regex]::Escape($_) } | Select-Object -First 1
        }
    }
    #>

    <#Filters on Display Name instead of email address
    if ($FilterTo -and $FilterTo.Count -gt 0) {
        $FilteredItems = $FilteredItems | Where-Object {
            $to = $_.To
            $FilterTo | Where-Object { $to -match [regex]::Escape($_) } | Select-Object -First 1
        }
    }
    if ($FilterCC -and $FilterCC.Count -gt 0) {
        $FilteredItems = $FilteredItems | Where-Object {
            $cc = $_.CC
            $FilterCC | Where-Object { $cc -match [regex]::Escape($_) } | Select-Object -First 1
        }
    }
    #>

    <#
    # To filter (Recipient.Type = 1)
    if ($FilterTo -and $FilterTo.Count -gt 0) {
        $FilteredItems = $FilteredItems | Where-Object {
            $toVals = Get-RecipientStrings -MailItem $_ -Type 1
            Test-AnyMatch -Haystack $toVals -Needles $FilterTo
        }
    }

    # CC filter (Recipient.Type = 2)
    if ($FilterCC -and $FilterCC.Count -gt 0) {
        $FilteredItems = $FilteredItems | Where-Object {
            $ccVals = Get-RecipientStrings -MailItem $_ -Type 2
            Test-AnyMatch -Haystack $ccVals -Needles $FilterCC
        }
    }


    if ($FilterBCC -and $FilterBCC.Count -gt 0) {
        $FilteredItems = $FilteredItems | Where-Object {
            $bcc = $_.BCC
            $FilterBCC | Where-Object { $bcc -match [regex]::Escape($_) } | Select-Object -First 1
        }
    }

    #>

        # ---- Combined To OR CC (takes precedence if provided) ----
    if ($FilterToOrCC -and $FilterToOrCC.Count -gt 0) {
        $FilteredItems = $FilteredItems | Where-Object {
            $toVals = Get-RecipientStrings -MailItem $_ -Type 1 -EmailOnly:$EmailOnly.IsPresent
            $ccVals = Get-RecipientStrings -MailItem $_ -Type 2 -EmailOnly:$EmailOnly.IsPresent
            $vals = @($toVals + $ccVals)
            Test-AnyMatch -Haystack $vals -Needles $FilterToOrCC
        }
    }
    else {
        # ---- Separate To ----
        if ($FilterTo -and $FilterTo.Count -gt 0) {
            $FilteredItems = $FilteredItems | Where-Object {
                $vals = Get-RecipientStrings -MailItem $_ -Type 1 -EmailOnly:$EmailOnly.IsPresent
                Test-AnyMatch -Haystack $vals -Needles $FilterTo
            }
        }
        # ---- Separate CC ----
        if ($FilterCC -and $FilterCC.Count -gt 0) {
            $FilteredItems = $FilteredItems | Where-Object {
                $vals = Get-RecipientStrings -MailItem $_ -Type 2 -EmailOnly:$EmailOnly.IsPresent
                Test-AnyMatch -Haystack $vals -Needles $FilterCC
            }
        }
    }



    write-verbose "FilteredItems: $($FilteredItems.count)"

    # Process emails in this folder
    foreach ($mailItem in $FilteredItems) {
        
        <#
        $attResults = foreach ($attachment in $mailItem.Attachments) {
            $contentId = Get-AttachmentContentId $attachment
                write-host $attachment.filename " -- type: " $attachment.type " -- contentId: " $contentId
            if ($attachment.Type -in 1,5 -and [string]::IsNullOrEmpty($contentId)) { $attachment }
        }
        #>

        write-verbose "=========================================="


        $attResults = foreach ($attachment in $mailItem.Attachments) {
            # Properties
            $pa = $attachment.PropertyAccessor

            # PR_ATTACH_CONTENT_ID (string)
            $cid = $null
            try { $cid = $pa.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x3712001E') } catch {}

            # PR_ATTACH_CONTENT_LOCATION (string)
            $cl = $null
            try { $cl = $pa.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x3713001E') } catch {}

            # PR_ATTACHMENT_HIDDEN (bool)
            $hidden = $false
            try { $hidden = $pa.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x7FFE000B') } catch {}

            $type = $attachment.Type           # 1=ByValue, 5=EmbeddedItem(.msg)
            $name = $attachment.FileName
            $size = $attachment.Size
            $ext  = [System.IO.Path]::GetExtension($name)

            Write-verbose ("Attachment Details: {0} -- type: {1} -- contentId: {2} -- hidden: {3} -- size: {4} -- ext: {5}" -f $name, $type,  $cid ,$hidden ,$size,  $ext)

            # Keep: typical "real" files and attached .msg
            $isCandidateType = ($type -in 1,5)
            $hasBody = ($size -gt 0)
            $isWantedExt = $ext -match '^\.(pdf|xlsx?|csv|zip|docx?|pptx?|msg)$'
            $isWantedExt = $true

            # Heuristic for inline assets (logos etc.)
            $isInlineImageExt = $ext -match '^\.(png|jpe?g|gif|svg)$'
            $looksInline = $hidden -or (
                $isInlineImageExt -and $size -lt 150kb -and (
                    -not [string]::IsNullOrEmpty($cid) -or -not [string]::IsNullOrEmpty($cl)
                )
            )

            if ($isCandidateType -and $hasBody -and $isWantedExt -and -not $looksInline) {
                $attachment
            }
        }
    


        Write-Verbose ("Subject: {0} From: {1} att: {2} results: {3}" -f $mailItem.Subject, $mailItem.SenderName, $mailItem.attachments.count, $attResults.count)

        if ($null -eq $attResults) { continue }

        Write-Host "Folder:        $currentFolderPath"
        Write-Host "Email Subject: $($mailItem.Subject)"
        Write-Host "Received:      $($mailItem.ReceivedTime)"
        Write-Host "From:          $($mailItem.SenderName) <$($mailItem.SenderEmailAddress)>"
        Write-Host "--------------------------------"

        $receivedDateStr = $mailItem.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss")

        # --- capture From, To, CC for the log ---
        $fromName  = $mailItem.SenderName
        $fromEmail = $mailItem.SenderEmailAddress
        $toList    = (Get-RecipientStrings -MailItem $mailItem -Type 1 -EmailOnly:$true) -join '; '   # Type 1 = To
        $ccList    = (Get-RecipientStrings -MailItem $mailItem -Type 2 -EmailOnly:$true) -join '; '   # Type 2 = CC
        # ---------------------------------------------

        foreach ($att in @($attResults)) {



            $logKey = New-LogKey `
                -MailBox      $MailBoxName `
                -Folder       $currentFolderPath `
                -ReceivedDate $receivedDateStr `
                -FileName     $att.FileName `
                -FileSize     $att.Size

            if ($SkipAlreadyDownloaded -and $downloadedSet.Contains($logKey)) {
                Write-Host "  [SKIPPED] $($att.FileName) (already downloaded)"
                $skippedCount++
                if (-not $NoLog) {
                    [void]$logBuffer.AppendLine((Build-LogLine -MailBox $MailBoxName -Folder $currentFolderPath `
                        -EmailSubject $mailItem.Subject -ReceivedDate $receivedDateStr `
                        -FileName $att.FileName -FileSize $att.Size -SavedTo "" -Status "Skipped"))
                }
                continue
            }

            #$exportPath = Join-Path -Path $resolvedSaveFolder -ChildPath $att.FileName
            $exportPath = Join-SafePath -Directory $resolvedSaveFolder -FileName $att.FileName

            if (Test-Path $exportPath) {
                switch ($FileCollisionAction) {
                    
                    "Suffix" {
                        $baseName = [System.IO.Path]::GetFileNameWithoutExtension( (Convert-ToSafeFileName $att.FileName) )
                        $extension = [System.IO.Path]::GetExtension($att.FileName)
                        $counter = 1
                        do {
                            $candidate = "{0}_{1}{2}" -f $baseName, $counter, $extension
                            $exportPath = Join-SafePath -Directory $resolvedSaveFolder -FileName $candidate
                            $counter++
                        } while (Test-Path $exportPath)
                        Write-Host " [SUFFIX] Collision - saving as: $(Split-Path $exportPath -Leaf)"
                    }

                    "Overwrite" {
                        Write-Host "  [OVERWRITE] Existing file will be replaced: $exportPath"
                    }
                    "Skip" {
                        Write-Host "  [SKIPPED] File already exists (FileCollisionAction=Skip): $($att.FileName)"
                        $skippedCount++
                        <#
                        if (-not $NoLog) {
                            [void]$logBuffer.AppendLine((Build-LogLine -MailBox $MailBoxName -Folder $currentFolderPath `
                                -EmailSubject $mailItem.Subject -ReceivedDate $receivedDateStr `
                                -FileName $att.FileName -FileSize $att.Size -SavedTo $exportPath -Status "Skipped"))
                        }
                        #>
                        if (-not $NoLog) {
                            [void]$logBuffer.AppendLine((Build-LogLine `
                                -MailBox       $MailBoxName `
                                -Folder        $currentFolderPath `
                                -EmailSubject  $mailItem.Subject `
                                -ReceivedDate  $receivedDateStr `
                                -FromName      $fromName `
                                -FromEmail     $fromEmail `
                                -ToEmails      $toList `
                                -CcEmails      $ccList `
                                -FileName      $att.FileName `
                                -FileSize      $att.Size `
                                -SavedTo       "" `
                                -Status        "Skipped"))
                        }


                        
                        continue
                    }
                    "Error" {
                        throw "File already exists and FileCollisionAction=Error: $exportPath"
                    }
                }
            }

            Write-Host "  [SAVING]  $($att.FileName) -> $exportPath"
            try {
            
                # Temporary: show any control chars so we can see if CR/LF are present
                $exportPath.ToCharArray() | Where-Object { [int]$_ -lt 32 } | ForEach-Object { '{0}({1})' -f $_, [int]$_ } | Out-Host

                $att.SaveAsFile($exportPath)
                $downloadCount++
                [void]$downloadedSet.Add($logKey)
                <#
                if (-not $NoLog) {
                    [void]$logBuffer.AppendLine((Build-LogLine -MailBox $MailBoxName -Folder $currentFolderPath `
                        -EmailSubject $mailItem.Subject -ReceivedDate $receivedDateStr `
                        -FileName $att.FileName -FileSize $att.Size -SavedTo $exportPath -Status "Downloaded"))
                }
                #>

                if (-not $NoLog) {
                    [void]$logBuffer.AppendLine((Build-LogLine `
                        -MailBox       $MailBoxName `
                        -Folder        $currentFolderPath `
                        -EmailSubject  $mailItem.Subject `
                        -ReceivedDate  $receivedDateStr `
                        -FromName      $fromName `
                        -FromEmail     $fromEmail `
                        -ToEmails      $toList `
                        -CcEmails      $ccList `
                        -FileName      $att.FileName `
                        -FileSize      $att.Size `
                        -SavedTo       $exportPath `
                        -Status        "Downloaded"))
                }

            } 
            catch {
                # Extra diagnostics for the '& in pipeline' class of failures
                $msg = $_.Exception.Message
                $pos = if ($_.InvocationInfo) { $_.InvocationInfo.PositionMessage } else { "" }
                Write-Warning (" Failed to save {0}: {1}{2}{3}" -f $att.FileName, $msg,
                               [Environment]::NewLine,
                               ($pos -replace '\r?\n+$',''))
            }

        }
    }

    # Flush log buffer for this folder before releasing COM objects
    if (-not $NoLog) { Flush-LogBuffer -Buffer $logBuffer -LogFile $resolvedLogPath }

    # Deterministically release COM references for this folder to prevent
    # handle accumulation over long recursive runs.
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($FilteredItems) } catch {}
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($currentFolder) } catch {}
    $FilteredItems = $null
    $currentFolder = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

} # end foreach folderEntry

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "==============================="
if ($useDateFilter) {
    Write-Host "  From       : $($effectiveFromDate.ToString('yyyy-MM-dd HH:mm:ss'))"
    Write-Host "  To         : $($effectiveToDate.ToString('yyyy-MM-dd HH:mm:ss'))"
}
Write-Host "  Save Folder: $resolvedSaveFolder"
Write-Host "  Downloaded : $downloadCount"
Write-Host "  Skipped    : $skippedCount"
if (-not $NoLog) { Write-Host "  Log        : $resolvedLogPath" }
Write-Host "==============================="
