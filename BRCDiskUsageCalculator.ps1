<#
BRC Disk Usage Calculator
Windows Server 2019 / Windows PowerShell 5.1

What it does:
- Opens a GUI
- Lets you select or type a folder path / UNC share path
- Walks the tree and measures the size of every directory and every file
- Writes one CSV containing a row per directory and a row per file
- Saves the CSV incrementally while it scans, and can resume an interrupted scan

Notes:
- Directory rows carry both a direct size (files sitting in that folder) and a
  total size (that folder plus everything underneath it).
- File rows carry the file length in the same size columns so the whole CSV can
  be sorted by SizeBytes to find the biggest items regardless of type.
- Reparse points (junctions and symlinks) are not traversed by default so their
  target data is not counted twice.
- Progress is checkpointed to a sidecar file next to the CSV ("<csv>.resume").
  A cancelled, crashed or power-cut scan can be restarted from that checkpoint
  without rescanning completed folders and without duplicating rows.
- This tool is read-only against the folders it scans. It does not modify files
  or folders under the scan root.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ------------------------------------------------------------
# Global state
# ------------------------------------------------------------

$script:ResumeFormatVersion = "1"

$script:CancelRequested = $false

# Output handles
$script:CsvStream = $null
$script:CsvWriter = $null
$script:LedgerStream = $null
$script:LedgerWriter = $null

# Checkpoint bookkeeping
$script:PendingFiles = @{}
$script:PendingSubtrees = New-Object 'System.Collections.Generic.List[object]'
$script:LastCheckpoint = $null
$script:CheckpointCount = 0

# Resume state loaded from a previous run
$script:ResumeSubtreeDone = @{}
$script:ResumeFilesDone = @{}

# Options, set once per scan so the recursive walker stays readable
$script:OptRecurse = $true
$script:OptIncludeRoot = $true
$script:OptIncludeDirectories = $true
$script:OptIncludeFiles = $true
$script:OptFollowReparsePoints = $false
$script:OptMinFileSizeBytes = [long]0
$script:OptCheckpointSeconds = 30

# Counters
$script:FolderCount = 0
$script:FileCount = 0
$script:ErrorCount = 0
$script:DirectoryRows = 0
$script:FileRows = 0
$script:SkippedFolders = 0

$script:StatusLabel = $null

$script:Columns = @(
    "Path",
    "ParentPath",
    "Name",
    "ItemType",
    "Extension",
    "Depth",
    "SizeBytes",
    "SizeMB",
    "SizeGB",
    "SizeFriendly",
    "DirectSizeBytes",
    "DirectSizeMB",
    "TotalSizeBytes",
    "TotalSizeMB",
    "TotalSizeGB",
    "PercentOfScanTotal",
    "DirectFileCount",
    "DirectSubfolderCount",
    "TotalFileCount",
    "TotalSubfolderCount",
    "CreationTime",
    "LastWriteTime",
    "LastAccessTime",
    "Attributes",
    "IsReparsePoint",
    "ScanRoot",
    "ErrorCount",
    "Errors"
)

# Zero based position of the columns the finalise pass needs
$script:SizeBytesColumnIndex = 6
$script:PercentColumnIndex = 15

# ------------------------------------------------------------
# CSV helpers
# ------------------------------------------------------------

function ConvertTo-CsvField {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return ""
    }

    $Text = [string]$Value

    if ($Text -match '[,"\r\n]') {
        return '"' + $Text.Replace('"', '""') + '"'
    }

    return $Text
}

function Write-CsvRow {
    param(
        [System.IO.StreamWriter]$Writer,
        [string[]]$Values
    )

    $Escaped = foreach ($Value in $Values) {
        ConvertTo-CsvField -Value $Value
    }

    $Writer.WriteLine(($Escaped -join ","))
}

function Split-CsvLine {
    param(
        [string]$Line
    )

    $Fields = New-Object 'System.Collections.Generic.List[string]'
    $Current = New-Object System.Text.StringBuilder
    $InQuotes = $false

    for ($Index = 0; $Index -lt $Line.Length; $Index++) {
        $Char = $Line[$Index]

        if ($InQuotes) {
            if ($Char -eq '"') {
                if ((($Index + 1) -lt $Line.Length) -and ($Line[$Index + 1] -eq '"')) {
                    [void]$Current.Append('"')
                    $Index++
                } else {
                    $InQuotes = $false
                }
            } else {
                [void]$Current.Append($Char)
            }
        } else {
            if ($Char -eq '"') {
                $InQuotes = $true
            } elseif ($Char -eq ',') {
                [void]$Fields.Add($Current.ToString())
                [void]$Current.Clear()
            } else {
                [void]$Current.Append($Char)
            }
        }
    }

    [void]$Fields.Add($Current.ToString())

    return $Fields.ToArray()
}

# ------------------------------------------------------------
# Formatting helpers
# ------------------------------------------------------------

function Format-Number {
    param(
        [AllowNull()]
        [object]$Value,
        [int]$Decimals = 0
    )

    if ($null -eq $Value) {
        return ""
    }

    if ($Value -is [string] -and [string]::IsNullOrWhiteSpace($Value)) {
        return ""
    }

    $Rounded = [math]::Round([double]$Value, $Decimals)

    return $Rounded.ToString(
        ("F{0}" -f $Decimals),
        [System.Globalization.CultureInfo]::InvariantCulture
    )
}

function Format-Bytes {
    param(
        [long]$Bytes
    )

    if ($Bytes -ge 1TB) {
        return "{0} TB" -f (Format-Number -Value ($Bytes / 1TB) -Decimals 2)
    }

    if ($Bytes -ge 1GB) {
        return "{0} GB" -f (Format-Number -Value ($Bytes / 1GB) -Decimals 2)
    }

    if ($Bytes -ge 1MB) {
        return "{0} MB" -f (Format-Number -Value ($Bytes / 1MB) -Decimals 2)
    }

    if ($Bytes -ge 1KB) {
        return "{0} KB" -f (Format-Number -Value ($Bytes / 1KB) -Decimals 2)
    }

    return "$Bytes B"
}

function Format-DateForCsv {
    param(
        [AllowNull()]
        [Nullable[datetime]]$DateValue
    )

    if ($null -eq $DateValue) {
        return ""
    }

    return ([datetime]$DateValue).ToString("yyyy-MM-dd HH:mm:ss")
}

function Get-DepthFromRoot {
    param(
        [string]$RootPath,
        [string]$CurrentPath
    )

    try {
        $RootFull = [System.IO.Path]::GetFullPath($RootPath).TrimEnd('\', '/')
        $CurrentFull = [System.IO.Path]::GetFullPath($CurrentPath).TrimEnd('\', '/')

        if ($CurrentFull.Length -le $RootFull.Length) {
            return 0
        }

        $Relative = $CurrentFull.Substring($RootFull.Length).TrimStart('\', '/')

        if ([string]::IsNullOrWhiteSpace($Relative)) {
            return 0
        }

        return (($Relative -split '[\\/]+').Count)
    } catch {
        return ""
    }
}

function Test-SamePath {
    param(
        [string]$FirstPath,
        [string]$SecondPath
    )

    try {
        $First = [System.IO.Path]::GetFullPath($FirstPath).TrimEnd('\', '/')
        $Second = [System.IO.Path]::GetFullPath($SecondPath).TrimEnd('\', '/')

        return ($First -ieq $Second)
    } catch {
        return ($FirstPath -ieq $SecondPath)
    }
}

function Test-IsReparsePoint {
    param(
        [System.IO.FileSystemInfo]$Info
    )

    if ($null -eq $Info) {
        return $false
    }

    return (($Info.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0)
}

function Get-PercentOfTotal {
    param(
        [long]$Bytes,
        [long]$TotalBytes
    )

    if ($TotalBytes -le 0) {
        return ""
    }

    return (Format-Number -Value ((100.0 * $Bytes) / $TotalBytes) -Decimals 4)
}

# ------------------------------------------------------------
# Row writers
#
# PercentOfScanTotal is left empty while the scan runs because the scan total is
# not known until the root folder finishes. It is filled in by the finalise pass
# once the scan completes.
# ------------------------------------------------------------

function Write-DirectoryRow {
    param(
        [System.IO.StreamWriter]$Writer,
        [psobject]$Entry,
        [string]$RootPath
    )

    $RowValues = @(
        $Entry.Path,
        $Entry.ParentPath,
        $Entry.Name,
        "Directory",
        "",
        [string]$Entry.Depth,
        [string]$Entry.TotalSizeBytes,
        (Format-Number -Value ($Entry.TotalSizeBytes / 1MB) -Decimals 3),
        (Format-Number -Value ($Entry.TotalSizeBytes / 1GB) -Decimals 4),
        (Format-Bytes -Bytes $Entry.TotalSizeBytes),
        [string]$Entry.DirectSizeBytes,
        (Format-Number -Value ($Entry.DirectSizeBytes / 1MB) -Decimals 3),
        [string]$Entry.TotalSizeBytes,
        (Format-Number -Value ($Entry.TotalSizeBytes / 1MB) -Decimals 3),
        (Format-Number -Value ($Entry.TotalSizeBytes / 1GB) -Decimals 4),
        "",
        [string]$Entry.DirectFileCount,
        [string]$Entry.DirectSubfolderCount,
        [string]$Entry.TotalFileCount,
        [string]$Entry.TotalSubfolderCount,
        (Format-DateForCsv -DateValue $Entry.CreationTime),
        (Format-DateForCsv -DateValue $Entry.LastWriteTime),
        (Format-DateForCsv -DateValue $Entry.LastAccessTime),
        $Entry.Attributes,
        [string]$Entry.IsReparsePoint,
        $RootPath,
        [string]$Entry.Errors.Count,
        (($Entry.Errors | Sort-Object -Unique) -join "; ")
    )

    Write-CsvRow -Writer $Writer -Values $RowValues
}

function Write-FileRow {
    param(
        [System.IO.StreamWriter]$Writer,
        [System.IO.FileInfo]$FileInfo,
        [string]$ParentPath,
        [string]$RootPath
    )

    $SizeBytes = [long]$FileInfo.Length

    $RowValues = @(
        $FileInfo.FullName,
        $ParentPath,
        $FileInfo.Name,
        "File",
        $FileInfo.Extension,
        [string](Get-DepthFromRoot -RootPath $RootPath -CurrentPath $FileInfo.FullName),
        [string]$SizeBytes,
        (Format-Number -Value ($SizeBytes / 1MB) -Decimals 3),
        (Format-Number -Value ($SizeBytes / 1GB) -Decimals 4),
        (Format-Bytes -Bytes $SizeBytes),
        "",
        "",
        [string]$SizeBytes,
        (Format-Number -Value ($SizeBytes / 1MB) -Decimals 3),
        (Format-Number -Value ($SizeBytes / 1GB) -Decimals 4),
        "",
        "",
        "",
        "",
        "",
        (Format-DateForCsv -DateValue $FileInfo.CreationTime),
        (Format-DateForCsv -DateValue $FileInfo.LastWriteTime),
        (Format-DateForCsv -DateValue $FileInfo.LastAccessTime),
        [string]$FileInfo.Attributes,
        [string](Test-IsReparsePoint -Info $FileInfo),
        $RootPath,
        "0",
        ""
    )

    Write-CsvRow -Writer $Writer -Values $RowValues
}

function Write-ErrorRow {
    param(
        [System.IO.StreamWriter]$Writer,
        [string]$Path,
        [string]$RootPath,
        [string]$Message
    )

    $RowValues = @(
        $Path,
        "",
        "",
        "Error",
        "",
        [string](Get-DepthFromRoot -RootPath $RootPath -CurrentPath $Path),
        "", "", "", "",
        "", "",
        "", "", "",
        "",
        "", "", "", "",
        "", "", "",
        "",
        "",
        $RootPath,
        "1",
        $Message
    )

    Write-CsvRow -Writer $Writer -Values $RowValues
}

# ------------------------------------------------------------
# Checkpoint ledger
#
# The ledger is an append only, tab separated sidecar file. Every data line ends
# with the CSV byte offset that was flushed to disk when that line was written:
#
#   F <TAB> path <TAB> directBytes <TAB> directFiles <TAB> directSubfolders <TAB> csvOffset
#   D <TAB> path <TAB> totalBytes  <TAB> totalFiles  <TAB> totalSubfolders  <TAB> csvOffset
#
# F means "the file rows for this folder are written". D means "this folder and
# everything under it is written, including its own directory row".
#
# The invariant that makes resuming safe: everything in the CSV before the last
# recorded offset belongs to a folder that has an F or D line. Anything written
# after that offset is unrecorded work, so a resume truncates the CSV back to the
# offset and redoes it. That is what stops rows being duplicated or lost when a
# scan dies part way through a folder.
#
# Windows forbids tab, CR and LF in file and folder names, so a tab separated
# ledger needs no escaping.
# ------------------------------------------------------------

function Get-ResumeFilePath {
    param(
        [string]$CsvPath
    )

    return "$CsvPath.resume"
}

function Get-OptionsFingerprint {
    param(
        [string]$RootPath,
        [bool]$Recurse,
        [bool]$IncludeRoot,
        [bool]$IncludeDirectories,
        [bool]$IncludeFiles,
        [bool]$FollowReparsePoints,
        [long]$MinFileSizeBytes
    )

    $Parts = @(
        "v$($script:ResumeFormatVersion)",
        "cols=$($script:Columns.Count)",
        "root=$($RootPath.TrimEnd('\', '/'))",
        "recurse=$Recurse",
        "includeRoot=$IncludeRoot",
        "dirs=$IncludeDirectories",
        "files=$IncludeFiles",
        "reparse=$FollowReparsePoints",
        "minBytes=$MinFileSizeBytes"
    )

    return ($Parts -join "|")
}

function Get-ParentPathValue {
    param(
        [string]$Path
    )

    try {
        return [System.IO.Path]::GetDirectoryName($Path.TrimEnd('\', '/'))
    } catch {
        return $null
    }
}

function Test-HasCompletedAncestor {
    param(
        [string]$Path,
        [hashtable]$CompletedSet
    )

    $Current = Get-ParentPathValue -Path $Path

    while (-not [string]::IsNullOrWhiteSpace($Current)) {
        if ($CompletedSet.ContainsKey($Current)) {
            return $true
        }

        $Next = Get-ParentPathValue -Path $Current

        if ($Next -eq $Current) {
            break
        }

        $Current = $Next
    }

    return $false
}

function Read-ResumeLedger {
    param(
        [string]$ResumePath,
        [string]$Fingerprint,
        [string]$CsvPath
    )

    $Result = [pscustomobject]@{
        Valid        = $false
        Reason       = ""
        CsvLength    = [long]0
        SubtreeDone  = @{}
        FilesDone    = @{}
        Records      = New-Object 'System.Collections.Generic.List[object]'
    }

    if (-not (Test-Path -LiteralPath $ResumePath -PathType Leaf)) {
        $Result.Reason = "No checkpoint file was found."
        return $Result
    }

    $LedgerFingerprint = ""
    $HeaderCheckpoint = [long]0
    $MaxOffset = [long]0

    $Reader = $null

    try {
        $Reader = New-Object System.IO.StreamReader($ResumePath, [System.Text.Encoding]::UTF8, $true)

        while ($null -ne ($Line = $Reader.ReadLine())) {
            if ($Line.StartsWith("#")) {
                $HeaderParts = $Line.Split("`t")

                if ($HeaderParts.Count -ge 2) {
                    switch ($HeaderParts[0]) {
                        "#Fingerprint" { $LedgerFingerprint = $HeaderParts[1] }
                        "#Checkpoint"  {
                            $ParsedHeader = [long]0
                            if ([long]::TryParse($HeaderParts[1], [ref]$ParsedHeader)) {
                                $HeaderCheckpoint = $ParsedHeader
                            }
                        }
                    }
                }

                continue
            }

            if ([string]::IsNullOrWhiteSpace($Line)) {
                continue
            }

            $Parts = $Line.Split("`t")

            # A torn final line from a crash fails this check, and everything
            # after a bad line is discarded rather than trusted.
            if ($Parts.Count -ne 6) {
                break
            }

            $Bytes = [long]0
            $Files = [long]0
            $Subfolders = [long]0
            $Offset = [long]0

            if (-not ([long]::TryParse($Parts[2], [ref]$Bytes) -and
                      [long]::TryParse($Parts[3], [ref]$Files) -and
                      [long]::TryParse($Parts[4], [ref]$Subfolders) -and
                      [long]::TryParse($Parts[5], [ref]$Offset))) {
                break
            }

            if ($Parts[0] -ne "F" -and $Parts[0] -ne "D") {
                break
            }

            $Record = [pscustomobject]@{
                Type       = $Parts[0]
                Path       = $Parts[1]
                Bytes      = $Bytes
                Files      = $Files
                Subfolders = $Subfolders
                Offset     = $Offset
            }

            $Result.Records.Add($Record)

            if ($Offset -gt $MaxOffset) {
                $MaxOffset = $Offset
            }
        }
    } catch {
        $Result.Reason = "The checkpoint file could not be read: $($_.Exception.Message)"
        return $Result
    } finally {
        if ($Reader) {
            $Reader.Close()
            $Reader.Dispose()
        }
    }

    if ($LedgerFingerprint -ne $Fingerprint) {
        $Result.Reason = "The checkpoint was taken with different scan settings or a different folder."
        return $Result
    }

    $CsvLength = [math]::Max($MaxOffset, $HeaderCheckpoint)

    if ($CsvLength -le 0) {
        $Result.Reason = "The checkpoint does not contain any completed work yet."
        return $Result
    }

    if (-not (Test-Path -LiteralPath $CsvPath -PathType Leaf)) {
        $Result.Reason = "The checkpoint refers to a CSV that no longer exists."
        return $Result
    }

    $ActualLength = (New-Object System.IO.FileInfo($CsvPath)).Length

    if ($ActualLength -lt $CsvLength) {
        $Result.Reason = "The CSV is smaller than the checkpoint expects, so it has been replaced or truncated since the last run."
        return $Result
    }

    foreach ($Record in $Result.Records) {
        if ($Record.Type -eq "D") {
            $Result.SubtreeDone[$Record.Path] = $Record
        } else {
            $Result.FilesDone[$Record.Path] = $Record
        }
    }

    # A folder inside a completed subtree never needs to be looked at again, so
    # its record is dropped. This keeps the ledger from growing without bound
    # across repeated resumes of a large share.
    foreach ($Key in @($Result.FilesDone.Keys)) {
        if ($Result.SubtreeDone.ContainsKey($Key) -or (Test-HasCompletedAncestor -Path $Key -CompletedSet $Result.SubtreeDone)) {
            $Result.FilesDone.Remove($Key)
        }
    }

    foreach ($Key in @($Result.SubtreeDone.Keys)) {
        if (Test-HasCompletedAncestor -Path $Key -CompletedSet $Result.SubtreeDone) {
            $Result.SubtreeDone.Remove($Key)
        }
    }

    $Result.CsvLength = $CsvLength
    $Result.Valid = $true

    return $Result
}

function Open-Ledger {
    param(
        [string]$ResumePath,
        [string]$Fingerprint,
        [string]$RootPath,
        [string]$CsvPath,
        [long]$CsvLength,
        [psobject]$ResumeState
    )

    $Encoding = New-Object System.Text.UTF8Encoding($false)
    $TempPath = "$ResumePath.new"

    $Stream = New-Object System.IO.FileStream($TempPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::Read)
    $Writer = New-Object System.IO.StreamWriter($Stream, $Encoding)

    $Writer.WriteLine("#BRCDiskUsageCalculator resume ledger")
    $Writer.WriteLine("#Fingerprint`t$Fingerprint")
    $Writer.WriteLine("#Root`t$RootPath")
    $Writer.WriteLine("#Csv`t$CsvPath")
    $Writer.WriteLine("#Checkpoint`t$CsvLength")
    $Writer.WriteLine("#Updated`t$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))")

    if ($null -ne $ResumeState) {
        foreach ($Record in $ResumeState.SubtreeDone.Values) {
            $Writer.WriteLine(("D`t{0}`t{1}`t{2}`t{3}`t{4}" -f $Record.Path, $Record.Bytes, $Record.Files, $Record.Subfolders, $Record.Offset))
        }

        foreach ($Record in $ResumeState.FilesDone.Values) {
            $Writer.WriteLine(("F`t{0}`t{1}`t{2}`t{3}`t{4}" -f $Record.Path, $Record.Bytes, $Record.Files, $Record.Subfolders, $Record.Offset))
        }
    }

    $Writer.Flush()
    $Stream.Flush($true)
    $Writer.Close()
    $Writer.Dispose()

    Move-Item -LiteralPath $TempPath -Destination $ResumePath -Force

    $script:LedgerStream = New-Object System.IO.FileStream($ResumePath, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::Read)
    $script:LedgerWriter = New-Object System.IO.StreamWriter($script:LedgerStream, $Encoding)
}

function Write-CompletedLedger {
    param(
        [string]$ResumePath,
        [string]$Fingerprint,
        [string]$RootPath,
        [string]$CsvPath,
        [long]$TotalBytes,
        [long]$TotalFiles,
        [long]$TotalSubfolders
    )

    # Called once the whole tree is done and the CSV is in its final shape. The
    # offsets taken during the scan are stale by now, because filling in
    # PercentOfScanTotal rewrites the file and changes its length. A single
    # record covering the root, carrying the real end of the finished CSV, keeps
    # a kept checkpoint truthful: re-running it truncates to the current end of
    # the file, which is a no-op, and skips the whole tree.
    $FinalLength = [long]0

    if (Test-Path -LiteralPath $CsvPath -PathType Leaf) {
        $FinalLength = (New-Object System.IO.FileInfo($CsvPath)).Length
    }

    $Encoding = New-Object System.Text.UTF8Encoding($false)
    $TempPath = "$ResumePath.new"

    $Stream = New-Object System.IO.FileStream($TempPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::Read)
    $Writer = New-Object System.IO.StreamWriter($Stream, $Encoding)

    $Writer.WriteLine("#BRCDiskUsageCalculator resume ledger")
    $Writer.WriteLine("#Fingerprint`t$Fingerprint")
    $Writer.WriteLine("#Root`t$RootPath")
    $Writer.WriteLine("#Csv`t$CsvPath")
    $Writer.WriteLine("#Checkpoint`t$FinalLength")
    $Writer.WriteLine("#Updated`t$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))")
    $Writer.WriteLine("#Complete`ttrue")
    $Writer.WriteLine(("D`t{0}`t{1}`t{2}`t{3}`t{4}" -f $RootPath, $TotalBytes, $TotalFiles, $TotalSubfolders, $FinalLength))

    $Writer.Flush()
    $Stream.Flush($true)
    $Writer.Close()
    $Writer.Dispose()

    Move-Item -LiteralPath $TempPath -Destination $ResumePath -Force
}

function Invoke-Checkpoint {
    param(
        [switch]$Force
    )

    if ($null -eq $script:CsvWriter -or $null -eq $script:LedgerWriter) {
        return
    }

    if (-not $Force) {
        if (((Get-Date) - $script:LastCheckpoint).TotalSeconds -lt $script:OptCheckpointSeconds) {
            return
        }
    }

    if ($script:PendingFiles.Count -eq 0 -and $script:PendingSubtrees.Count -eq 0) {
        $script:LastCheckpoint = Get-Date
        return
    }

    # Order matters. The CSV bytes must be on disk before the ledger claims them,
    # otherwise a crash in between would leave the ledger pointing past the end
    # of the CSV and the resumed scan would skip rows that were never written.
    $script:CsvWriter.Flush()
    $script:CsvStream.Flush($true)

    $Offset = $script:CsvStream.Position

    $CompletedNow = @{}

    foreach ($Record in $script:PendingSubtrees) {
        $CompletedNow[$Record.Path] = $true
    }

    foreach ($Key in @($script:PendingFiles.Keys)) {
        if ($CompletedNow.ContainsKey($Key)) {
            continue
        }

        $Record = $script:PendingFiles[$Key]
        $script:LedgerWriter.WriteLine(("F`t{0}`t{1}`t{2}`t{3}`t{4}" -f $Key, $Record.Bytes, $Record.Files, $Record.Subfolders, $Offset))
    }

    foreach ($Record in $script:PendingSubtrees) {
        $script:LedgerWriter.WriteLine(("D`t{0}`t{1}`t{2}`t{3}`t{4}" -f $Record.Path, $Record.Bytes, $Record.Files, $Record.Subfolders, $Offset))
    }

    $script:LedgerWriter.Flush()
    $script:LedgerStream.Flush($true)

    $script:PendingFiles.Clear()
    $script:PendingSubtrees.Clear()
    $script:LastCheckpoint = Get-Date
    $script:CheckpointCount++
}

# ------------------------------------------------------------
# The scan itself
#
# Depth first, and a folder's own directory row is written only once its whole
# subtree is done. That ordering is what lets the CSV be appended to safely: a
# completed folder is never revisited, so its rows never need to change.
# ------------------------------------------------------------

function Update-ScanStatus {
    $script:StatusLabel.Text = "Scanning... $($script:FolderCount) folders, $($script:FileCount) files, $($script:DirectoryRows + $script:FileRows) rows written, $($script:SkippedFolders) folders skipped from checkpoint."
    [System.Windows.Forms.Application]::DoEvents()
}

function Invoke-DirectoryScan {
    param(
        [string]$CurrentPath,
        [string]$RootPath
    )

    if ($script:ResumeSubtreeDone.ContainsKey($CurrentPath)) {
        $Done = $script:ResumeSubtreeDone[$CurrentPath]
        $script:SkippedFolders++

        return [pscustomobject]@{
            TotalSizeBytes      = [long]$Done.Bytes
            TotalFileCount      = [long]$Done.Files
            TotalSubfolderCount = [long]$Done.Subfolders
        }
    }

    $script:FolderCount++

    if (($script:FolderCount % 25) -eq 0) {
        Update-ScanStatus
    }

    $Errors = New-Object 'System.Collections.Generic.List[string]'

    $DirectSizeBytes = [long]0
    $DirectFileCount = [long]0
    $DirectSubfolderCount = [long]0

    $CreationTime = $null
    $LastWriteTime = $null
    $LastAccessTime = $null
    $Attributes = ""
    $ParentPath = ""
    $Name = ""
    $IsReparsePoint = $false

    try {
        $DirInfo = New-Object System.IO.DirectoryInfo($CurrentPath)

        $Name = $DirInfo.Name
        $CreationTime = $DirInfo.CreationTime
        $LastWriteTime = $DirInfo.LastWriteTime
        $LastAccessTime = $DirInfo.LastAccessTime
        $Attributes = [string]$DirInfo.Attributes
        $IsReparsePoint = Test-IsReparsePoint -Info $DirInfo

        if ($null -ne $DirInfo.Parent) {
            $ParentPath = $DirInfo.Parent.FullName
        }
    } catch {
        $Errors.Add("Could not read directory info: $($_.Exception.Message)")
        $script:ErrorCount++
    }

    $IsRoot = Test-SamePath -FirstPath $CurrentPath -SecondPath $RootPath
    $Traversed = $true

    if ($IsReparsePoint -and -not $script:OptFollowReparsePoints -and -not $IsRoot) {
        $Traversed = $false
        $Errors.Add("Reparse point was not traversed. Enable 'Follow reparse points' to include its contents.")
    }

    $ChildPaths = New-Object 'System.Collections.Generic.List[string]'
    $FilesRecord = $script:ResumeFilesDone[$CurrentPath]
    $EnumerationFailed = $false

    if ($Traversed -and $null -ne $FilesRecord) {
        # The file rows for this folder are already in the CSV from an earlier
        # run, so only the subfolders still need walking.
        $DirectSizeBytes = [long]$FilesRecord.Bytes
        $DirectFileCount = [long]$FilesRecord.Files
        $script:SkippedFolders++

        try {
            foreach ($ChildPath in [System.IO.Directory]::EnumerateDirectories($CurrentPath)) {
                $DirectSubfolderCount++
                $ChildPaths.Add($ChildPath)
            }
        } catch {
            $Errors.Add("Could not enumerate subfolders: $($_.Exception.Message)")
            $script:ErrorCount++
            $EnumerationFailed = $true
        }
    } elseif ($Traversed) {
        try {
            foreach ($FilePath in [System.IO.Directory]::EnumerateFiles($CurrentPath)) {
                if ($script:CancelRequested) {
                    break
                }

                $DirectFileCount++
                $script:FileCount++

                try {
                    $FileInfo = New-Object System.IO.FileInfo($FilePath)
                    $FileLength = [long]$FileInfo.Length
                    $DirectSizeBytes += $FileLength

                    if ($script:OptIncludeFiles -and $FileLength -ge $script:OptMinFileSizeBytes) {
                        Write-FileRow `
                            -Writer $script:CsvWriter `
                            -FileInfo $FileInfo `
                            -ParentPath $CurrentPath `
                            -RootPath $RootPath

                        $script:FileRows++

                        if (($script:FileRows % 500) -eq 0) {
                            Update-ScanStatus
                        }
                    }
                } catch {
                    $script:ErrorCount++

                    Write-ErrorRow `
                        -Writer $script:CsvWriter `
                        -Path $FilePath `
                        -RootPath $RootPath `
                        -Message "Could not read file info: $($_.Exception.Message)"
                }
            }
        } catch {
            $Errors.Add("Could not enumerate files: $($_.Exception.Message)")
            $script:ErrorCount++
            $EnumerationFailed = $true
        }

        try {
            foreach ($ChildPath in [System.IO.Directory]::EnumerateDirectories($CurrentPath)) {
                $DirectSubfolderCount++
                $ChildPaths.Add($ChildPath)
            }
        } catch {
            $Errors.Add("Could not enumerate subfolders: $($_.Exception.Message)")
            $script:ErrorCount++
            $EnumerationFailed = $true
        }

        # A folder that hit an enumeration error is deliberately left out of the
        # ledger. Redoing it on resume is cheap and it keeps the error text on
        # the directory row rather than losing it to the checkpoint.
        if (-not $EnumerationFailed -and -not $script:CancelRequested) {
            $script:PendingFiles[$CurrentPath] = [pscustomobject]@{
                Bytes      = $DirectSizeBytes
                Files      = $DirectFileCount
                Subfolders = $DirectSubfolderCount
            }

            Invoke-Checkpoint
        }
    }

    $TotalSizeBytes = $DirectSizeBytes
    $TotalFileCount = $DirectFileCount
    $TotalSubfolderCount = $DirectSubfolderCount

    if ($script:OptRecurse) {
        foreach ($ChildPath in $ChildPaths) {
            if ($script:CancelRequested) {
                break
            }

            $ChildResult = Invoke-DirectoryScan -CurrentPath $ChildPath -RootPath $RootPath

            # A cancelled child is incomplete, so its partial totals must not be
            # folded into this folder.
            if ($script:CancelRequested) {
                break
            }

            $TotalSizeBytes += [long]$ChildResult.TotalSizeBytes
            $TotalFileCount += [long]$ChildResult.TotalFileCount
            $TotalSubfolderCount += [long]$ChildResult.TotalSubfolderCount
        }
    }

    $Result = [pscustomobject]@{
        TotalSizeBytes      = $TotalSizeBytes
        TotalFileCount      = $TotalFileCount
        TotalSubfolderCount = $TotalSubfolderCount
    }

    if ($script:CancelRequested) {
        return $Result
    }

    if ($script:OptIncludeDirectories -and ($script:OptIncludeRoot -or -not $IsRoot)) {
        $Entry = [pscustomobject]@{
            Path                 = $CurrentPath
            ParentPath           = $ParentPath
            Name                 = $Name
            Depth                = (Get-DepthFromRoot -RootPath $RootPath -CurrentPath $CurrentPath)
            DirectSizeBytes      = $DirectSizeBytes
            TotalSizeBytes       = $TotalSizeBytes
            DirectFileCount      = $DirectFileCount
            DirectSubfolderCount = $DirectSubfolderCount
            TotalFileCount       = $TotalFileCount
            TotalSubfolderCount  = $TotalSubfolderCount
            CreationTime         = $CreationTime
            LastWriteTime        = $LastWriteTime
            LastAccessTime       = $LastAccessTime
            Attributes           = $Attributes
            IsReparsePoint       = $IsReparsePoint
            Errors               = $Errors
        }

        Write-DirectoryRow -Writer $script:CsvWriter -Entry $Entry -RootPath $RootPath
        $script:DirectoryRows++
    }

    $script:PendingSubtrees.Add([pscustomobject]@{
        Path       = $CurrentPath
        Bytes      = $TotalSizeBytes
        Files      = $TotalFileCount
        Subfolders = $TotalSubfolderCount
    })

    Invoke-Checkpoint

    return $Result
}

# ------------------------------------------------------------
# Finalise pass: fill PercentOfScanTotal now the scan total is known
# ------------------------------------------------------------

function Complete-PercentColumn {
    param(
        [string]$CsvPath,
        [long]$ScanTotalBytes,
        [System.Windows.Forms.Label]$StatusLabel
    )

    if ($ScanTotalBytes -le 0) {
        return [pscustomobject]@{ Succeeded = $true; Message = "" }
    }

    $TempPath = "$CsvPath.tmp"
    $BackupPath = "$CsvPath.bak"

    $Reader = $null
    $Writer = $null
    $RecordCount = 0
    $Succeeded = $false

    try {
        $Reader = New-Object System.IO.StreamReader($CsvPath, [System.Text.Encoding]::UTF8, $true)
        $Writer = New-Object System.IO.StreamWriter($TempPath, $false, (New-Object System.Text.UTF8Encoding($true)))

        $IsHeader = $true
        $Record = $null
        $QuoteCount = 0

        while ($null -ne ($Line = $Reader.ReadLine())) {
            if ($null -eq $Record) {
                $Record = $Line
                $QuoteCount = 0
            } else {
                $Record = $Record + "`r`n" + $Line
            }

            $QuoteCount += ([regex]::Matches($Line, '"')).Count

            # An odd number of quotes means a field contains a line break and
            # the record continues on the next physical line.
            if (($QuoteCount % 2) -ne 0) {
                continue
            }

            if ($IsHeader) {
                $Writer.WriteLine($Record)
                $IsHeader = $false
            } else {
                $Fields = Split-CsvLine -Line $Record

                if ($Fields.Count -eq $script:Columns.Count) {
                    $SizeValue = [long]0

                    if ([long]::TryParse($Fields[$script:SizeBytesColumnIndex], [ref]$SizeValue)) {
                        $Fields[$script:PercentColumnIndex] = Get-PercentOfTotal -Bytes $SizeValue -TotalBytes $ScanTotalBytes
                    }

                    $Escaped = foreach ($Field in $Fields) {
                        ConvertTo-CsvField -Value $Field
                    }

                    $Writer.WriteLine(($Escaped -join ","))
                } else {
                    $Writer.WriteLine($Record)
                }

                $RecordCount++

                if (($RecordCount % 20000) -eq 0) {
                    $StatusLabel.Text = "Finalising percentages... $RecordCount rows."
                    [System.Windows.Forms.Application]::DoEvents()
                }
            }

            $Record = $null
        }

        if ($null -ne $Record) {
            $Writer.WriteLine($Record)
        }

        $Writer.Flush()
        $Writer.Close()
        $Writer.Dispose()
        $Writer = $null

        $Reader.Close()
        $Reader.Dispose()
        $Reader = $null

        if (Test-Path -LiteralPath $BackupPath -PathType Leaf) {
            Remove-Item -LiteralPath $BackupPath -Force
        }

        Move-Item -LiteralPath $CsvPath -Destination $BackupPath -Force
        Move-Item -LiteralPath $TempPath -Destination $CsvPath -Force
        Remove-Item -LiteralPath $BackupPath -Force

        $Succeeded = $true

        return [pscustomobject]@{ Succeeded = $true; Message = "" }
    } catch {
        return [pscustomobject]@{
            Succeeded = $false
            Message   = "$($_.Exception.Message)"
        }
    } finally {
        if ($Writer) {
            $Writer.Close()
            $Writer.Dispose()
        }

        if ($Reader) {
            $Reader.Close()
            $Reader.Dispose()
        }

        # The scan data itself is never at risk here: the original CSV is only
        # swapped out once the rewrite has been written and closed.
        if (-not $Succeeded) {
            try {
                if (Test-Path -LiteralPath $TempPath -PathType Leaf) {
                    Remove-Item -LiteralPath $TempPath -Force
                }
            } catch {}
        }
    }
}

# ------------------------------------------------------------
# Orchestration
# ------------------------------------------------------------

function Start-DiskUsageScan {
    param(
        [string]$RootPath,
        [string]$CsvPath,
        [bool]$Recurse,
        [bool]$IncludeRoot,
        [bool]$IncludeDirectories,
        [bool]$IncludeFiles,
        [bool]$FollowReparsePoints,
        [bool]$OpenWhenComplete,
        [bool]$ResumeIfPossible,
        [bool]$FillPercentages,
        [bool]$RemoveCheckpointWhenComplete,
        [int]$CheckpointSeconds,
        [double]$MinFileSizeMb,
        [System.Windows.Forms.Label]$StatusLabel,
        [System.Windows.Forms.ProgressBar]$ProgressBar
    )

    if ([string]::IsNullOrWhiteSpace($RootPath)) {
        [System.Windows.Forms.MessageBox]::Show("Please select or enter a folder/share path.", "Missing path", "OK", "Warning") | Out-Null
        return
    }

    if (-not (Test-Path -LiteralPath $RootPath -PathType Container)) {
        [System.Windows.Forms.MessageBox]::Show("The selected path does not exist or is not accessible:`r`n`r`n$RootPath", "Invalid path", "OK", "Error") | Out-Null
        return
    }

    if ([string]::IsNullOrWhiteSpace($CsvPath)) {
        [System.Windows.Forms.MessageBox]::Show("Please select an output CSV path.", "Missing CSV path", "OK", "Warning") | Out-Null
        return
    }

    if (-not $IncludeDirectories -and -not $IncludeFiles) {
        [System.Windows.Forms.MessageBox]::Show("Select at least one of 'Include directory rows' or 'Include file rows'.", "Nothing to export", "OK", "Warning") | Out-Null
        return
    }

    $MinFileSizeBytes = [long]([math]::Round($MinFileSizeMb * 1MB))
    $ResumePath = Get-ResumeFilePath -CsvPath $CsvPath

    $Fingerprint = Get-OptionsFingerprint `
        -RootPath $RootPath `
        -Recurse $Recurse `
        -IncludeRoot $IncludeRoot `
        -IncludeDirectories $IncludeDirectories `
        -IncludeFiles $IncludeFiles `
        -FollowReparsePoints $FollowReparsePoints `
        -MinFileSizeBytes $MinFileSizeBytes

    $ResumeState = $null

    if (Test-Path -LiteralPath $ResumePath -PathType Leaf) {
        if ($ResumeIfPossible) {
            $Loaded = Read-ResumeLedger -ResumePath $ResumePath -Fingerprint $Fingerprint -CsvPath $CsvPath

            if ($Loaded.Valid) {
                $ResumeState = $Loaded
            } else {
                $Answer = [System.Windows.Forms.MessageBox]::Show(
                    "A checkpoint exists but it cannot be used to resume this scan.`r`n`r`nReason: $($Loaded.Reason)`r`n`r`nStart a new scan instead? This overwrites:`r`n$CsvPath",
                    "Checkpoint cannot be resumed",
                    "YesNo",
                    "Warning"
                )

                if ($Answer -ne [System.Windows.Forms.DialogResult]::Yes) {
                    $StatusLabel.Text = "Cancelled. Nothing was written."
                    return
                }
            }
        } else {
            $Answer = [System.Windows.Forms.MessageBox]::Show(
                "'Resume from checkpoint' is switched off and a checkpoint already exists for this CSV.`r`n`r`nStarting a new scan discards that checkpoint and overwrites:`r`n$CsvPath`r`n`r`nContinue?",
                "Existing checkpoint will be discarded",
                "YesNo",
                "Warning"
            )

            if ($Answer -ne [System.Windows.Forms.DialogResult]::Yes) {
                $StatusLabel.Text = "Cancelled. Nothing was written."
                return
            }
        }
    } elseif ((Test-Path -LiteralPath $CsvPath -PathType Leaf) -and ((New-Object System.IO.FileInfo($CsvPath)).Length -gt 0)) {
        $Answer = [System.Windows.Forms.MessageBox]::Show(
            "This CSV already exists and there is no checkpoint to resume from, so it will be overwritten:`r`n`r`n$CsvPath`r`n`r`nContinue?",
            "Overwrite existing CSV",
            "YesNo",
            "Warning"
        )

        if ($Answer -ne [System.Windows.Forms.DialogResult]::Yes) {
            $StatusLabel.Text = "Cancelled. Nothing was written."
            return
        }
    }

    # Reset per scan state
    $script:CancelRequested = $false
    $script:PendingFiles = @{}
    $script:PendingSubtrees = New-Object 'System.Collections.Generic.List[object]'
    $script:CheckpointCount = 0
    $script:FolderCount = 0
    $script:FileCount = 0
    $script:ErrorCount = 0
    $script:DirectoryRows = 0
    $script:FileRows = 0
    $script:SkippedFolders = 0
    $script:StatusLabel = $StatusLabel

    $script:OptRecurse = $Recurse
    $script:OptIncludeRoot = $IncludeRoot
    $script:OptIncludeDirectories = $IncludeDirectories
    $script:OptIncludeFiles = $IncludeFiles
    $script:OptFollowReparsePoints = $FollowReparsePoints
    $script:OptMinFileSizeBytes = $MinFileSizeBytes
    $script:OptCheckpointSeconds = $CheckpointSeconds

    if ($null -ne $ResumeState) {
        $script:ResumeSubtreeDone = $ResumeState.SubtreeDone
        $script:ResumeFilesDone = $ResumeState.FilesDone
    } else {
        $script:ResumeSubtreeDone = @{}
        $script:ResumeFilesDone = @{}
    }

    $ProgressBar.Style = "Marquee"
    $ProgressBar.MarqueeAnimationSpeed = 25
    $StatusLabel.Text = if ($null -ne $ResumeState) { "Resuming from checkpoint..." } else { "Starting scan..." }
    [System.Windows.Forms.Application]::DoEvents()

    $StartTime = Get-Date
    $script:LastCheckpoint = $StartTime

    $Completed = $false
    $ScanTotalBytes = [long]0
    $ScanTotalFiles = [long]0
    $ScanTotalSubfolders = [long]0

    try {
        $OutputDirectory = Split-Path -Path $CsvPath -Parent

        if (-not [string]::IsNullOrWhiteSpace($OutputDirectory)) {
            if (-not (Test-Path -LiteralPath $OutputDirectory -PathType Container)) {
                New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
            }
        }

        $CsvEncoding = New-Object System.Text.UTF8Encoding($false)

        if ($null -ne $ResumeState) {
            $script:CsvStream = New-Object System.IO.FileStream($CsvPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Write, [System.IO.FileShare]::Read)

            # Anything past the checkpoint is work that was never recorded as
            # finished, so it is cut off and redone rather than left to duplicate.
            $script:CsvStream.SetLength($ResumeState.CsvLength)
            $script:CsvStream.Position = $ResumeState.CsvLength

            $script:CsvWriter = New-Object System.IO.StreamWriter($script:CsvStream, $CsvEncoding)
        } else {
            $script:CsvStream = New-Object System.IO.FileStream($CsvPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::Read)

            $Preamble = (New-Object System.Text.UTF8Encoding($true)).GetPreamble()
            $script:CsvStream.Write($Preamble, 0, $Preamble.Length)

            $script:CsvWriter = New-Object System.IO.StreamWriter($script:CsvStream, $CsvEncoding)
            Write-CsvRow -Writer $script:CsvWriter -Values $script:Columns
        }

        $script:CsvWriter.AutoFlush = $false

        Open-Ledger `
            -ResumePath $ResumePath `
            -Fingerprint $Fingerprint `
            -RootPath $RootPath `
            -CsvPath $CsvPath `
            -CsvLength $(if ($null -ne $ResumeState) { $ResumeState.CsvLength } else { [long]0 }) `
            -ResumeState $ResumeState

        $RootResult = Invoke-DirectoryScan -CurrentPath $RootPath -RootPath $RootPath

        Invoke-Checkpoint -Force

        $ScanTotalBytes = [long]$RootResult.TotalSizeBytes
        $ScanTotalFiles = [long]$RootResult.TotalFileCount
        $ScanTotalSubfolders = [long]$RootResult.TotalSubfolderCount
        $Completed = -not $script:CancelRequested

        $script:CsvWriter.Flush()
        $script:CsvStream.Flush($true)
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "The scan failed:`r`n`r`n$($_.Exception.Message)`r`n`r`nWhatever was written before the failure is still in the CSV. The checkpoint file has been kept so the scan can be resumed.",
            "Scan failed",
            "OK",
            "Error"
        ) | Out-Null

        $StatusLabel.Text = "Scan failed. Checkpoint kept for resume."

        $ProgressBar.MarqueeAnimationSpeed = 0
        $ProgressBar.Style = "Blocks"
        return
    } finally {
        if ($script:CsvWriter) {
            $script:CsvWriter.Close()
            $script:CsvWriter.Dispose()
            $script:CsvWriter = $null
            $script:CsvStream = $null
        }

        if ($script:LedgerWriter) {
            $script:LedgerWriter.Close()
            $script:LedgerWriter.Dispose()
            $script:LedgerWriter = $null
            $script:LedgerStream = $null
        }
    }

    $PercentNote = ""

    if ($Completed -and $FillPercentages) {
        $StatusLabel.Text = "Finalising percentages..."
        [System.Windows.Forms.Application]::DoEvents()

        $PercentResult = Complete-PercentColumn -CsvPath $CsvPath -ScanTotalBytes $ScanTotalBytes -StatusLabel $StatusLabel

        if (-not $PercentResult.Succeeded) {
            $PercentNote = "`r`n`r`nPercentOfScanTotal could not be filled in: $($PercentResult.Message)`r`nThe scan data itself is intact."
        }
    } elseif ($Completed) {
        $PercentNote = "`r`n`r`nPercentOfScanTotal was left empty because the finalise step is switched off."
    }

    $Elapsed = (Get-Date) - $StartTime
    $ElapsedText = "{0:hh\:mm\:ss}" -f $Elapsed

    if ($Completed) {
        if ($RemoveCheckpointWhenComplete) {
            try {
                if (Test-Path -LiteralPath $ResumePath -PathType Leaf) {
                    Remove-Item -LiteralPath $ResumePath -Force
                }
            } catch {
                $PercentNote += "`r`n`r`nThe checkpoint file could not be deleted: $($_.Exception.Message)"
            }
        } else {
            try {
                Write-CompletedLedger `
                    -ResumePath $ResumePath `
                    -Fingerprint $Fingerprint `
                    -RootPath $RootPath `
                    -CsvPath $CsvPath `
                    -TotalBytes $ScanTotalBytes `
                    -TotalFiles $ScanTotalFiles `
                    -TotalSubfolders $ScanTotalSubfolders
            } catch {
                $PercentNote += "`r`n`r`nThe checkpoint file could not be updated after completion: $($_.Exception.Message)`r`nDelete '$ResumePath' before running this scan again."
            }
        }

        $StatusLabel.Text = "Complete. $($script:FolderCount) folders scanned, $($script:SkippedFolders) resumed, $(Format-Bytes -Bytes $ScanTotalBytes). Errors: $($script:ErrorCount)"

        [System.Windows.Forms.MessageBox]::Show(
            "Disk usage scan complete.`r`n`r`nFolders scanned this run: $($script:FolderCount)`r`nFolders skipped from checkpoint: $($script:SkippedFolders)`r`nFiles measured this run: $($script:FileCount)`r`nTotal size: $(Format-Bytes -Bytes $ScanTotalBytes) ($ScanTotalBytes bytes)`r`n`r`nDirectory rows written this run: $($script:DirectoryRows)`r`nFile rows written this run: $($script:FileRows)`r`nCheckpoints saved: $($script:CheckpointCount)`r`nErrors: $($script:ErrorCount)`r`nElapsed: $ElapsedText`r`n`r`nCSV saved to:`r`n$CsvPath$PercentNote",
            "Complete",
            "OK",
            "Information"
        ) | Out-Null

        if ($OpenWhenComplete) {
            try {
                Start-Process -FilePath $CsvPath | Out-Null
            } catch {
                [System.Windows.Forms.MessageBox]::Show(
                    "The CSV was saved but could not be opened automatically:`r`n`r`n$($_.Exception.Message)",
                    "Could not open CSV",
                    "OK",
                    "Warning"
                ) | Out-Null
            }
        }
    } else {
        $StatusLabel.Text = "Cancelled at $($script:FolderCount) folders. Checkpoint saved, tick 'Resume' and start again to continue."

        [System.Windows.Forms.MessageBox]::Show(
            "Scan cancelled.`r`n`r`nFolders scanned this run: $($script:FolderCount)`r`nFiles measured this run: $($script:FileCount)`r`nRows written this run: $($script:DirectoryRows + $script:FileRows)`r`nCheckpoints saved: $($script:CheckpointCount)`r`nElapsed: $ElapsedText`r`n`r`nEverything up to the last checkpoint is saved in:`r`n$CsvPath`r`n`r`nTo carry on, leave 'Resume from checkpoint if one exists' ticked and press Start Scan again with the same path and settings.`r`n`r`nPercentOfScanTotal stays empty until a run finishes the whole tree.",
            "Cancelled",
            "OK",
            "Information"
        ) | Out-Null
    }

    $ProgressBar.MarqueeAnimationSpeed = 0
    $ProgressBar.Style = "Blocks"
    $script:CancelRequested = $false
}

# ------------------------------------------------------------
# GUI
# ------------------------------------------------------------

[System.Windows.Forms.Application]::EnableVisualStyles()

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "BRC Disk Usage Calculator"
$Form.Size = New-Object System.Drawing.Size(820, 545)
$Form.StartPosition = "CenterScreen"
$Form.MaximizeBox = $false
$Form.FormBorderStyle = "FixedDialog"

$LabelPath = New-Object System.Windows.Forms.Label
$LabelPath.Text = "Folder path / UNC share path:"
$LabelPath.Location = New-Object System.Drawing.Point(20, 20)
$LabelPath.Size = New-Object System.Drawing.Size(220, 20)
$Form.Controls.Add($LabelPath)

$TextPath = New-Object System.Windows.Forms.TextBox
$TextPath.Location = New-Object System.Drawing.Point(20, 45)
$TextPath.Size = New-Object System.Drawing.Size(650, 24)
$Form.Controls.Add($TextPath)

$ButtonBrowse = New-Object System.Windows.Forms.Button
$ButtonBrowse.Text = "Browse..."
$ButtonBrowse.Location = New-Object System.Drawing.Point(680, 43)
$ButtonBrowse.Size = New-Object System.Drawing.Size(95, 28)
$Form.Controls.Add($ButtonBrowse)

$LabelCsv = New-Object System.Windows.Forms.Label
$LabelCsv.Text = "Output CSV file:"
$LabelCsv.Location = New-Object System.Drawing.Point(20, 85)
$LabelCsv.Size = New-Object System.Drawing.Size(220, 20)
$Form.Controls.Add($LabelCsv)

$TextCsv = New-Object System.Windows.Forms.TextBox
$TextCsv.Location = New-Object System.Drawing.Point(20, 110)
$TextCsv.Size = New-Object System.Drawing.Size(650, 24)

$DefaultCsvName = "DiskUsage_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss")
$TextCsv.Text = Join-Path -Path ([Environment]::GetFolderPath("Desktop")) -ChildPath $DefaultCsvName

$Form.Controls.Add($TextCsv)

$ButtonCsv = New-Object System.Windows.Forms.Button
$ButtonCsv.Text = "Save As..."
$ButtonCsv.Location = New-Object System.Drawing.Point(680, 108)
$ButtonCsv.Size = New-Object System.Drawing.Size(95, 28)
$Form.Controls.Add($ButtonCsv)

$LabelMinSize = New-Object System.Windows.Forms.Label
$LabelMinSize.Text = "Only report files at least:"
$LabelMinSize.Location = New-Object System.Drawing.Point(20, 155)
$LabelMinSize.Size = New-Object System.Drawing.Size(170, 20)
$Form.Controls.Add($LabelMinSize)

$NumericMinSize = New-Object System.Windows.Forms.NumericUpDown
$NumericMinSize.Location = New-Object System.Drawing.Point(190, 153)
$NumericMinSize.Size = New-Object System.Drawing.Size(80, 24)
$NumericMinSize.Minimum = 0
$NumericMinSize.Maximum = 1000000
$NumericMinSize.DecimalPlaces = 2
$NumericMinSize.Increment = 1
$NumericMinSize.Value = 0
$Form.Controls.Add($NumericMinSize)

$LabelMinSizeUnit = New-Object System.Windows.Forms.Label
$LabelMinSizeUnit.Text = "MB  (0 = report every file; folder totals always include every file)"
$LabelMinSizeUnit.Location = New-Object System.Drawing.Point(280, 155)
$LabelMinSizeUnit.Size = New-Object System.Drawing.Size(495, 20)
$Form.Controls.Add($LabelMinSizeUnit)

$CheckRecurse = New-Object System.Windows.Forms.CheckBox
$CheckRecurse.Text = "Recurse subfolders"
$CheckRecurse.Location = New-Object System.Drawing.Point(20, 190)
$CheckRecurse.Size = New-Object System.Drawing.Size(180, 24)
$CheckRecurse.Checked = $true
$Form.Controls.Add($CheckRecurse)

$CheckIncludeRoot = New-Object System.Windows.Forms.CheckBox
$CheckIncludeRoot.Text = "Include selected root folder in output"
$CheckIncludeRoot.Location = New-Object System.Drawing.Point(220, 190)
$CheckIncludeRoot.Size = New-Object System.Drawing.Size(260, 24)
$CheckIncludeRoot.Checked = $true
$Form.Controls.Add($CheckIncludeRoot)

$CheckFollowReparse = New-Object System.Windows.Forms.CheckBox
$CheckFollowReparse.Text = "Follow reparse points (junctions/symlinks)"
$CheckFollowReparse.Location = New-Object System.Drawing.Point(500, 190)
$CheckFollowReparse.Size = New-Object System.Drawing.Size(280, 24)
$CheckFollowReparse.Checked = $false
$Form.Controls.Add($CheckFollowReparse)

$CheckIncludeDirectories = New-Object System.Windows.Forms.CheckBox
$CheckIncludeDirectories.Text = "Include directory rows"
$CheckIncludeDirectories.Location = New-Object System.Drawing.Point(20, 220)
$CheckIncludeDirectories.Size = New-Object System.Drawing.Size(180, 24)
$CheckIncludeDirectories.Checked = $true
$Form.Controls.Add($CheckIncludeDirectories)

$CheckIncludeFiles = New-Object System.Windows.Forms.CheckBox
$CheckIncludeFiles.Text = "Include file rows"
$CheckIncludeFiles.Location = New-Object System.Drawing.Point(220, 220)
$CheckIncludeFiles.Size = New-Object System.Drawing.Size(260, 24)
$CheckIncludeFiles.Checked = $true
$Form.Controls.Add($CheckIncludeFiles)

$CheckOpenWhenDone = New-Object System.Windows.Forms.CheckBox
$CheckOpenWhenDone.Text = "Open CSV when complete"
$CheckOpenWhenDone.Location = New-Object System.Drawing.Point(500, 220)
$CheckOpenWhenDone.Size = New-Object System.Drawing.Size(280, 24)
$CheckOpenWhenDone.Checked = $false
$Form.Controls.Add($CheckOpenWhenDone)

$GroupResume = New-Object System.Windows.Forms.GroupBox
$GroupResume.Text = "Incremental save and resume"
$GroupResume.Location = New-Object System.Drawing.Point(20, 252)
$GroupResume.Size = New-Object System.Drawing.Size(755, 110)
$Form.Controls.Add($GroupResume)

$CheckResume = New-Object System.Windows.Forms.CheckBox
$CheckResume.Text = "Resume from checkpoint if one exists"
$CheckResume.Location = New-Object System.Drawing.Point(15, 25)
$CheckResume.Size = New-Object System.Drawing.Size(270, 24)
$CheckResume.Checked = $true
$GroupResume.Controls.Add($CheckResume)

$CheckFillPercent = New-Object System.Windows.Forms.CheckBox
$CheckFillPercent.Text = "Fill PercentOfScanTotal when the scan finishes"
$CheckFillPercent.Location = New-Object System.Drawing.Point(300, 25)
$CheckFillPercent.Size = New-Object System.Drawing.Size(320, 24)
$CheckFillPercent.Checked = $true
$GroupResume.Controls.Add($CheckFillPercent)

$CheckRemoveCheckpoint = New-Object System.Windows.Forms.CheckBox
$CheckRemoveCheckpoint.Text = "Delete checkpoint file when complete"
$CheckRemoveCheckpoint.Location = New-Object System.Drawing.Point(15, 52)
$CheckRemoveCheckpoint.Size = New-Object System.Drawing.Size(270, 24)
$CheckRemoveCheckpoint.Checked = $true
$GroupResume.Controls.Add($CheckRemoveCheckpoint)

$LabelCheckpoint = New-Object System.Windows.Forms.Label
$LabelCheckpoint.Text = "Save a checkpoint at least every:"
$LabelCheckpoint.Location = New-Object System.Drawing.Point(300, 55)
$LabelCheckpoint.Size = New-Object System.Drawing.Size(200, 20)
$GroupResume.Controls.Add($LabelCheckpoint)

$NumericCheckpoint = New-Object System.Windows.Forms.NumericUpDown
$NumericCheckpoint.Location = New-Object System.Drawing.Point(505, 53)
$NumericCheckpoint.Size = New-Object System.Drawing.Size(70, 24)
$NumericCheckpoint.Minimum = 1
$NumericCheckpoint.Maximum = 3600
$NumericCheckpoint.Value = 30
$GroupResume.Controls.Add($NumericCheckpoint)

$LabelCheckpointUnit = New-Object System.Windows.Forms.Label
$LabelCheckpointUnit.Text = "seconds"
$LabelCheckpointUnit.Location = New-Object System.Drawing.Point(585, 55)
$LabelCheckpointUnit.Size = New-Object System.Drawing.Size(80, 20)
$GroupResume.Controls.Add($LabelCheckpointUnit)

$LabelResumeHelp = New-Object System.Windows.Forms.Label
$LabelResumeHelp.Text = "Rows are appended to the CSV as the scan runs. Progress is checkpointed to '<your csv>.resume', so a cancelled or crashed scan can be restarted from where it stopped using the same path and settings."
$LabelResumeHelp.Location = New-Object System.Drawing.Point(15, 80)
$LabelResumeHelp.Size = New-Object System.Drawing.Size(725, 30)
$GroupResume.Controls.Add($LabelResumeHelp)

$ModeHelp = New-Object System.Windows.Forms.Label
$ModeHelp.Text = "Tip: directory rows report both DirectSizeBytes (files in that folder only) and TotalSizeBytes (folder plus everything under it). Sort the CSV by SizeBytes to see the largest folders and files together."
$ModeHelp.Location = New-Object System.Drawing.Point(20, 372)
$ModeHelp.Size = New-Object System.Drawing.Size(755, 34)
$Form.Controls.Add($ModeHelp)

$ProgressBar = New-Object System.Windows.Forms.ProgressBar
$ProgressBar.Location = New-Object System.Drawing.Point(20, 412)
$ProgressBar.Size = New-Object System.Drawing.Size(755, 24)
$ProgressBar.Style = "Blocks"
$Form.Controls.Add($ProgressBar)

$StatusLabel = New-Object System.Windows.Forms.Label
$StatusLabel.Text = "Ready."
$StatusLabel.Location = New-Object System.Drawing.Point(20, 444)
$StatusLabel.Size = New-Object System.Drawing.Size(755, 24)
$Form.Controls.Add($StatusLabel)

$ButtonStart = New-Object System.Windows.Forms.Button
$ButtonStart.Text = "Start Scan"
$ButtonStart.Location = New-Object System.Drawing.Point(470, 472)
$ButtonStart.Size = New-Object System.Drawing.Size(95, 30)
$Form.Controls.Add($ButtonStart)

$ButtonCancel = New-Object System.Windows.Forms.Button
$ButtonCancel.Text = "Cancel"
$ButtonCancel.Location = New-Object System.Drawing.Point(575, 472)
$ButtonCancel.Size = New-Object System.Drawing.Size(95, 30)
$ButtonCancel.Enabled = $false
$Form.Controls.Add($ButtonCancel)

$ButtonClose = New-Object System.Windows.Forms.Button
$ButtonClose.Text = "Close"
$ButtonClose.Location = New-Object System.Drawing.Point(680, 472)
$ButtonClose.Size = New-Object System.Drawing.Size(95, 30)
$Form.Controls.Add($ButtonClose)

# ------------------------------------------------------------
# GUI events
# ------------------------------------------------------------

$ButtonBrowse.Add_Click({
    $FolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderDialog.Description = "Select the folder or share root to measure"
    $FolderDialog.ShowNewFolderButton = $false

    if (-not [string]::IsNullOrWhiteSpace($TextPath.Text)) {
        try {
            if (Test-Path -LiteralPath $TextPath.Text -PathType Container) {
                $FolderDialog.SelectedPath = $TextPath.Text
            }
        } catch {}
    }

    $Result = $FolderDialog.ShowDialog()

    if ($Result -eq [System.Windows.Forms.DialogResult]::OK) {
        $TextPath.Text = $FolderDialog.SelectedPath
    }
})

$ButtonCsv.Add_Click({
    $SaveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveDialog.Title = "Save disk usage CSV"
    $SaveDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $SaveDialog.OverwritePrompt = $false
    $SaveDialog.FileName = [System.IO.Path]::GetFileName($TextCsv.Text)

    $InitialDirectory = Split-Path -Path $TextCsv.Text -Parent
    if (-not [string]::IsNullOrWhiteSpace($InitialDirectory)) {
        $SaveDialog.InitialDirectory = $InitialDirectory
    }

    $Result = $SaveDialog.ShowDialog()

    if ($Result -eq [System.Windows.Forms.DialogResult]::OK) {
        $TextCsv.Text = $SaveDialog.FileName
    }
})

$ButtonCancel.Add_Click({
    $script:CancelRequested = $true
    $StatusLabel.Text = "Cancelling, saving checkpoint..."
})

$ButtonStart.Add_Click({
    $ButtonStart.Enabled = $false
    $ButtonClose.Enabled = $false
    $ButtonCancel.Enabled = $true

    try {
        Start-DiskUsageScan `
            -RootPath $TextPath.Text `
            -CsvPath $TextCsv.Text `
            -Recurse $CheckRecurse.Checked `
            -IncludeRoot $CheckIncludeRoot.Checked `
            -IncludeDirectories $CheckIncludeDirectories.Checked `
            -IncludeFiles $CheckIncludeFiles.Checked `
            -FollowReparsePoints $CheckFollowReparse.Checked `
            -OpenWhenComplete $CheckOpenWhenDone.Checked `
            -ResumeIfPossible $CheckResume.Checked `
            -FillPercentages $CheckFillPercent.Checked `
            -RemoveCheckpointWhenComplete $CheckRemoveCheckpoint.Checked `
            -CheckpointSeconds ([int]$NumericCheckpoint.Value) `
            -MinFileSizeMb ([double]$NumericMinSize.Value) `
            -StatusLabel $StatusLabel `
            -ProgressBar $ProgressBar
    } finally {
        $ButtonStart.Enabled = $true
        $ButtonClose.Enabled = $true
        $ButtonCancel.Enabled = $false
    }
})

$ButtonClose.Add_Click({
    $Form.Close()
})

[void]$Form.ShowDialog()
