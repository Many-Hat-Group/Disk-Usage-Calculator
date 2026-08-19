<#
BRC Stale Data Archiver
Windows Server 2019 / Windows PowerShell 5.1

What it does:
- Opens a GUI
- Lets you select the CSV that BRC-StaleDirectoryFinder.ps1 produced
- Reads the folder paths out of that CSV (optionally only the rows where IsStale = True)
- Lets you select a destination folder or UNC share
- Checks the destination has enough free space BEFORE anything is copied, and
  refuses to start if the copy would leave less than the configured free space
  headroom (20 GB by default)
- Copies every file under each listed folder, preserving the folder structure
- Hashes each file while it is being read, hashes the file that landed at the
  destination, and confirms the two checksums match
- Writes a verification report CSV listing every file, both checksums and the result

Notes:
- The source data is only ever read. Nothing is deleted or modified at the source.
- File copies do NOT carry NTFS permissions across. Whoever can read the
  destination can read the archived data. Check the destination ACLs first.
- Free space is re-checked while the copy runs, so a share filling up from
  somewhere else still stops the job safely.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

if (-not ([System.Management.Automation.PSTypeName]'BRC.NativeDisk').Type) {
    Add-Type -Namespace BRC -Name NativeDisk -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError = true, CharSet = System.Runtime.InteropServices.CharSet.Auto)]
[return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
public static extern bool GetDiskFreeSpaceEx(
    string lpDirectoryName,
    out ulong lpFreeBytesAvailable,
    out ulong lpTotalNumberOfBytes,
    out ulong lpTotalNumberOfFreeBytes);
'@
}

$Script:CancelRequested = $false
$Script:IsRunning = $false

# ---------------------------------------------------------------------------
# CSV helpers - same escaping style as the other BRC tools
# ---------------------------------------------------------------------------

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

function Format-DateForCsv {
    param(
        [AllowNull()]
        [object]$DateValue
    )

    if ($null -eq $DateValue) {
        return ""
    }

    try {
        return ([datetime]$DateValue).ToString("yyyy-MM-dd HH:mm:ss")
    } catch {
        return ""
    }
}

function Format-Bytes {
    param([double]$Bytes)

    if ($Bytes -ge 1TB) { return ("{0:N2} TB" -f ($Bytes / 1TB)) }
    if ($Bytes -ge 1GB) { return ("{0:N2} GB" -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ("{0:N2} MB" -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ("{0:N2} KB" -f ($Bytes / 1KB)) }
    return ("{0:N0} bytes" -f $Bytes)
}

# ---------------------------------------------------------------------------
# Long path helpers. Windows PowerShell 5.1 trips over paths beyond 260
# characters, which is exactly what old archive folders are full of.
# ---------------------------------------------------------------------------

function Get-LongPath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $Path
    }

    if ($Path.StartsWith('\\?\')) {
        return $Path
    }

    if ($Path.StartsWith('\\')) {
        return '\\?\UNC\' + $Path.Substring(2)
    }

    if ($Path -match '^[A-Za-z]:\\') {
        return '\\?\' + $Path
    }

    return $Path
}

function ConvertFrom-LongPath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $Path
    }

    if ($Path.StartsWith('\\?\UNC\')) {
        return '\\' + $Path.Substring(8)
    }

    if ($Path.StartsWith('\\?\')) {
        return $Path.Substring(4)
    }

    return $Path
}

function Test-DirectoryExistsSafe {
    param([string]$Path)

    try {
        return [System.IO.Directory]::Exists((Get-LongPath -Path $Path))
    } catch {
        return $false
    }
}

function Test-FileExistsSafe {
    param([string]$Path)

    try {
        return [System.IO.File]::Exists((Get-LongPath -Path $Path))
    } catch {
        return $false
    }
}

function New-DirectoryTree {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return
    }

    if (Test-DirectoryExistsSafe -Path $Path) {
        return
    }

    $Parent = $null
    try {
        $Parent = [System.IO.Path]::GetDirectoryName($Path)
    } catch {
        $Parent = $null
    }

    if ($Parent -and $Parent -ne $Path) {
        New-DirectoryTree -Path $Parent
    }

    [void][System.IO.Directory]::CreateDirectory((Get-LongPath -Path $Path))
}

function Get-NormalisedPath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return ""
    }

    $Trimmed = $Path.Trim().Trim('"')

    try {
        $Full = [System.IO.Path]::GetFullPath($Trimmed)
    } catch {
        $Full = $Trimmed
    }

    if ($Full.Length -gt 3) {
        $Full = $Full.TrimEnd('\')
    }

    return $Full
}

function Test-PathIsInside {
    <#
        Both paths must already have been through Get-NormalisedPath. This is
        called in loops, and normalising here instead costs two GetFullPath
        calls per comparison.
    #>
    param(
        [string]$ChildPath,
        [string]$ParentPath
    )

    if ([string]::IsNullOrWhiteSpace($ChildPath) -or [string]::IsNullOrWhiteSpace($ParentPath)) {
        return $false
    }

    $Child = $ChildPath.TrimEnd('\') + '\'
    $Parent = $ParentPath.TrimEnd('\') + '\'

    return $Child.StartsWith($Parent, [System.StringComparison]::OrdinalIgnoreCase)
}

function Get-ParentPathString {
    <#
        The parent of an already normalised path, by string only. No filesystem
        access and no GetFullPath, so it is cheap enough to call in a loop.
        Returns "" once there is nothing left above.
    #>
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return ""
    }

    $Trimmed = $Path.TrimEnd('\')
    $Index = $Trimmed.LastIndexOf('\')

    # Index 0 or 1 means we are at "\" or the "\\server" part of a UNC path,
    # so there is no useful parent left.
    if ($Index -lt 2) {
        return ""
    }

    return $Trimmed.Substring(0, $Index)
}

# ---------------------------------------------------------------------------
# Reading the BRC-StaleDirectoryFinder.ps1 CSV
# ---------------------------------------------------------------------------

$Script:PathColumnCandidates = @(
    "Path",
    "FullName",
    "FolderPath",
    "DirectoryPath",
    "SourcePath",
    "FullPath"
)

function Get-PathColumnName {
    param([string[]]$ColumnNames)

    foreach ($Candidate in $Script:PathColumnCandidates) {
        foreach ($Column in $ColumnNames) {
            if ($Column -ieq $Candidate) {
                return $Column
            }
        }
    }

    return $null
}

function Read-SourcePathsFromCsv {
    <#
        Returns the distinct root paths listed in the CSV, with any path that
        already sits underneath another listed path removed. The finder can
        report a stale parent and its stale children, and without this the
        same files would be copied and hashed more than once.
    #>
    param(
        [string]$CsvPath,
        [bool]$OnlyStaleRows,
        [scriptblock]$OnProgress
    )

    $Result = [pscustomobject]@{
        Paths = New-Object System.Collections.Generic.List[string]
        RowCount = 0
        SkippedNotStale = 0
        SkippedMissing = 0
        SkippedNested = 0
        SkippedBlank = 0
        MissingPaths = New-Object System.Collections.Generic.List[string]
        PathColumn = $null
        Errors = New-Object System.Collections.Generic.List[string]
    }

    $Rows = @(Import-Csv -LiteralPath $CsvPath)
    $Result.RowCount = $Rows.Count

    if ($Rows.Count -eq 0) {
        $Result.Errors.Add("The CSV has no data rows.")
        return $Result
    }

    $ColumnNames = @($Rows[0].PSObject.Properties.Name)
    $PathColumn = Get-PathColumnName -ColumnNames $ColumnNames

    if (-not $PathColumn) {
        $Result.Errors.Add("Could not find a path column. Looked for: $($Script:PathColumnCandidates -join ', '). Columns present: $($ColumnNames -join ', ').")
        return $Result
    }

    $Result.PathColumn = $PathColumn
    $HasStaleColumn = $ColumnNames -contains "IsStale"

    $Candidates = New-Object System.Collections.Generic.List[string]
    $Seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    $RowIndex = 0

    foreach ($Row in $Rows) {
        if ($Script:CancelRequested) {
            break
        }

        $RowIndex++

        if ($OnProgress -and ($RowIndex % 50) -eq 0) {
            & $OnProgress "Checking the folders listed in the CSV" $RowIndex $Rows.Count
        }

        $RawPath = [string]$Row.$PathColumn

        if ([string]::IsNullOrWhiteSpace($RawPath)) {
            $Result.SkippedBlank++
            continue
        }

        if ($OnlyStaleRows -and $HasStaleColumn) {
            $StaleValue = [string]$Row.IsStale
            if ($StaleValue -notmatch '^(?i)\s*(true|1|yes)\s*$') {
                $Result.SkippedNotStale++
                continue
            }
        }

        $Normalised = Get-NormalisedPath -Path $RawPath

        if ([string]::IsNullOrWhiteSpace($Normalised)) {
            $Result.SkippedBlank++
            continue
        }

        if (-not $Seen.Add($Normalised)) {
            continue
        }

        if (-not (Test-DirectoryExistsSafe -Path $Normalised) -and -not (Test-FileExistsSafe -Path $Normalised)) {
            $Result.SkippedMissing++
            if ($Result.MissingPaths.Count -lt 25) {
                $Result.MissingPaths.Add($Normalised)
            }
            continue
        }

        $Candidates.Add($Normalised)
    }

    # Shortest paths first, so a parent is always seen before its children.
    $Sorted = @($Candidates | Sort-Object -Property Length)
    $Kept = New-Object System.Collections.Generic.List[string]
    $KeptKeys = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $Checked = 0

    foreach ($Candidate in $Sorted) {
        if ($Script:CancelRequested) {
            break
        }

        # Walk this path's own parent chain and look each ancestor up directly.
        # Comparing every candidate against every kept path instead makes this
        # n-squared, which on a real CSV is minutes of frozen window.
        $IsNested = $false
        $Ancestor = Get-ParentPathString -Path $Candidate

        while (-not [string]::IsNullOrEmpty($Ancestor)) {
            if ($KeptKeys.Contains($Ancestor)) {
                $IsNested = $true
                break
            }

            $Ancestor = Get-ParentPathString -Path $Ancestor
        }

        if ($IsNested) {
            $Result.SkippedNested++
        } else {
            [void]$KeptKeys.Add($Candidate.TrimEnd('\'))
            $Kept.Add($Candidate)
        }

        $Checked++

        if ($OnProgress -and ($Checked % 500) -eq 0) {
            & $OnProgress "Removing folders that sit inside another folder" $Checked $Sorted.Count
        }
    }

    foreach ($Path in ($Kept | Sort-Object)) {
        $Result.Paths.Add($Path)
    }

    return $Result
}

# ---------------------------------------------------------------------------
# Working out what has to be copied
# ---------------------------------------------------------------------------

function Get-MirrorRelativePath {
    <#
        \\server\share\dept\old  ->  server\share\dept\old
        D:\data\old              ->  D\data\old
        So two sources with the same leaf name can never collide at the destination.
    #>
    param([string]$Path)

    $Working = $Path

    if ($Working.StartsWith('\\')) {
        $Working = $Working.Substring(2)
    } elseif ($Working -match '^([A-Za-z]):\\(.*)$') {
        $Working = $Matches[1] + '\' + $Matches[2]
    }

    # Spelled out rather than taken from GetInvalidFileNameChars() so the
    # result is the same on every host the script runs on.
    $Invalid = [char[]]@('<', '>', ':', '"', '/', '\', '|', '?', '*')
    $Parts = $Working.Split('\') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    $Clean = foreach ($Part in $Parts) {
        $Builder = New-Object System.Text.StringBuilder
        foreach ($Char in $Part.ToCharArray()) {
            if (($Invalid -contains $Char) -or ([int]$Char -lt 32)) {
                [void]$Builder.Append('_')
            } else {
                [void]$Builder.Append($Char)
            }
        }
        $Builder.ToString().TrimEnd(' ', '.')
    }

    return ($Clean -join '\')
}

function Get-DestinationRootPath {
    param(
        [string]$SourceRoot,
        [string]$DestinationRoot,
        [string]$LayoutMode,
        [System.Collections.Generic.HashSet[string]]$UsedRoots
    )

    if ($LayoutMode -eq "Folder name only") {
        $Leaf = Split-Path -Path $SourceRoot -Leaf

        if ([string]::IsNullOrWhiteSpace($Leaf)) {
            $Leaf = "Root"
        }

        $Candidate = Join-Path -Path $DestinationRoot -ChildPath $Leaf
        $Suffix = 2

        while (-not $UsedRoots.Add($Candidate.ToLowerInvariant())) {
            $Candidate = Join-Path -Path $DestinationRoot -ChildPath ("{0}_{1}" -f $Leaf, $Suffix)
            $Suffix++
        }

        return $Candidate
    }

    $Relative = Get-MirrorRelativePath -Path $SourceRoot
    $Candidate = Join-Path -Path $DestinationRoot -ChildPath $Relative
    [void]$UsedRoots.Add($Candidate.ToLowerInvariant())

    return $Candidate
}

function Get-FilesUnderPath {
    <#
        Queue based walk so one unreadable subfolder does not abort the rest.
        Enumerates through the \\?\ form and converts back, so long paths and
        legacy names with trailing dots or spaces are still picked up.
    #>
    param(
        [string]$RootPath,
        [System.Collections.Generic.List[string]]$Errors,
        [System.Collections.Generic.List[string]]$Directories,
        [scriptblock]$OnProgress
    )

    $Files = New-Object System.Collections.Generic.List[object]

    if (Test-FileExistsSafe -Path $RootPath) {
        try {
            $Info = New-Object System.IO.FileInfo((Get-LongPath -Path $RootPath))
            $Files.Add([pscustomobject]@{
                FullName = $RootPath
                RelativePath = [System.IO.Path]::GetFileName($RootPath)
                Length = [long]$Info.Length
                LastWriteTimeUtc = $Info.LastWriteTimeUtc
            })
        } catch {
            $Errors.Add("$RootPath : $($_.Exception.Message)")
        }

        return ,$Files
    }

    $Queue = New-Object System.Collections.Generic.Queue[string]
    $Queue.Enqueue($RootPath)
    $RootPrefixLength = $RootPath.TrimEnd('\', '/').Length

    while ($Queue.Count -gt 0) {
        if ($Script:CancelRequested) {
            break
        }

        $Current = $Queue.Dequeue()
        $Directory = $null

        if ($null -ne $Directories) {
            $RelativeFolder = $Current.Substring($RootPrefixLength).TrimStart('\', '/')

            if (-not [string]::IsNullOrWhiteSpace($RelativeFolder)) {
                $Directories.Add($RelativeFolder)
            }
        }

        try {
            $Directory = New-Object System.IO.DirectoryInfo((Get-LongPath -Path $Current))
        } catch {
            $Errors.Add("$Current : $($_.Exception.Message)")
            continue
        }

        try {
            foreach ($File in $Directory.EnumerateFiles()) {
                $PlainPath = ConvertFrom-LongPath -Path $File.FullName
                $Relative = $PlainPath.Substring($RootPrefixLength).TrimStart('\', '/')

                $Files.Add([pscustomobject]@{
                    FullName = $PlainPath
                    RelativePath = $Relative
                    Length = [long]$File.Length
                    LastWriteTimeUtc = $File.LastWriteTimeUtc
                })

                if ($OnProgress -and ($Files.Count % 5000) -eq 0) {
                    & $OnProgress $Files.Count

                    if ($Script:CancelRequested) {
                        break
                    }
                }
            }
        } catch {
            $Errors.Add("$Current : could not list files - $($_.Exception.Message)")
        }

        try {
            foreach ($SubDirectory in $Directory.EnumerateDirectories()) {
                # Reparse points are listed but not walked, so junction targets
                # are not copied twice.
                if (($SubDirectory.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -eq [System.IO.FileAttributes]::ReparsePoint) {
                    $Errors.Add("$(ConvertFrom-LongPath -Path $SubDirectory.FullName) : reparse point skipped")
                    continue
                }

                $Queue.Enqueue((ConvertFrom-LongPath -Path $SubDirectory.FullName))
            }
        } catch {
            $Errors.Add("$Current : could not list subfolders - $($_.Exception.Message)")
        }

        if ($OnProgress) {
            & $OnProgress $Files.Count
        }
    }

    return ,$Files
}

# ---------------------------------------------------------------------------
# Free space
# ---------------------------------------------------------------------------

function Get-FreeSpaceInfo {
    <#
        GetDiskFreeSpaceEx works for local drives, mapped drives and UNC paths.
        FreeBytesAvailable is the figure that respects per-user quotas on a
        share, so that is the number the checks are based on.
    #>
    param([string]$Path)

    $Info = [pscustomobject]@{
        Path = $Path
        FreeBytes = $null
        TotalBytes = $null
        Succeeded = $false
        Error = ""
    }

    $Probe = $Path

    # Walk up until we hit a folder that exists - the destination may not be created yet.
    while (-not [string]::IsNullOrWhiteSpace($Probe) -and -not (Test-DirectoryExistsSafe -Path $Probe)) {
        $Parent = [System.IO.Path]::GetDirectoryName($Probe)

        if ([string]::IsNullOrWhiteSpace($Parent) -or $Parent -eq $Probe) {
            break
        }

        $Probe = $Parent
    }

    if ([string]::IsNullOrWhiteSpace($Probe)) {
        $Info.Error = "No existing folder found to measure."
        return $Info
    }

    $FreeBytesAvailable = [uint64]0
    $TotalBytes = [uint64]0
    $TotalFreeBytes = [uint64]0

    try {
        $Ok = [BRC.NativeDisk]::GetDiskFreeSpaceEx(
            ($Probe.TrimEnd('\') + '\'),
            [ref]$FreeBytesAvailable,
            [ref]$TotalBytes,
            [ref]$TotalFreeBytes)

        if ($Ok) {
            $Info.FreeBytes = [double]$FreeBytesAvailable
            $Info.TotalBytes = [double]$TotalBytes
            $Info.Succeeded = $true
        } else {
            $Code = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
            $Info.Error = "GetDiskFreeSpaceEx failed on '$Probe' (Win32 error $Code)."
        }
    } catch {
        $Info.Error = $_.Exception.Message
    }

    return $Info
}

function Test-FreeSpaceForCopy {
    param(
        [string]$DestinationPath,
        [double]$RequiredBytes,
        [double]$MarginBytes
    )

    $Space = Get-FreeSpaceInfo -Path $DestinationPath

    $Result = [pscustomobject]@{
        Succeeded = $Space.Succeeded
        FreeBytes = $Space.FreeBytes
        TotalBytes = $Space.TotalBytes
        RequiredBytes = $RequiredBytes
        MarginBytes = $MarginBytes
        RemainingAfterCopy = $null
        IsSufficient = $false
        Reason = ""
    }

    if (-not $Space.Succeeded) {
        $Result.Reason = "Free space could not be read. $($Space.Error)"
        return $Result
    }

    $Remaining = $Space.FreeBytes - $RequiredBytes
    $Result.RemainingAfterCopy = $Remaining

    if ($Space.FreeBytes -lt $MarginBytes) {
        $Result.Reason = "The destination only has $(Format-Bytes $Space.FreeBytes) free, which is already below the $(Format-Bytes $MarginBytes) minimum."
        return $Result
    }

    if ($RequiredBytes -gt $Space.FreeBytes) {
        $Result.Reason = "The copy needs $(Format-Bytes $RequiredBytes) but only $(Format-Bytes $Space.FreeBytes) is free."
        return $Result
    }

    if ($Remaining -lt $MarginBytes) {
        $Result.Reason = "The copy needs $(Format-Bytes $RequiredBytes) and would leave only $(Format-Bytes $Remaining) free, below the $(Format-Bytes $MarginBytes) minimum."
        return $Result
    }

    $Result.IsSufficient = $true
    $Result.Reason = "$(Format-Bytes $Space.FreeBytes) free, $(Format-Bytes $RequiredBytes) needed, $(Format-Bytes $Remaining) would remain."

    return $Result
}

# ---------------------------------------------------------------------------
# Copy and checksum
# ---------------------------------------------------------------------------

function Get-FileChecksum {
    param(
        [string]$Path,
        [string]$Algorithm,
        [int]$BufferSize = 4194304
    )

    $Hasher = $null
    $Stream = $null

    try {
        $Hasher = [System.Security.Cryptography.HashAlgorithm]::Create($Algorithm)
        $Stream = New-Object System.IO.FileStream(
            (Get-LongPath -Path $Path),
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read,
            [System.IO.FileShare]::ReadWrite,
            $BufferSize,
            [System.IO.FileOptions]::SequentialScan)

        $Buffer = New-Object byte[] $BufferSize
        $Read = 0

        while (($Read = $Stream.Read($Buffer, 0, $BufferSize)) -gt 0) {
            [void]$Hasher.TransformBlock($Buffer, 0, $Read, $null, 0)
        }

        [void]$Hasher.TransformFinalBlock((New-Object byte[] 0), 0, 0)

        return [System.BitConverter]::ToString($Hasher.Hash).Replace("-", "")
    } finally {
        if ($Stream) { $Stream.Dispose() }
        if ($Hasher) { $Hasher.Dispose() }
    }
}

function Copy-FileWithChecksum {
    <#
        Reads the source once, hashing the bytes on their way into the
        destination file. The destination is then read back and hashed
        independently, so the comparison proves what actually landed on disk
        matches what was read from the source.
    #>
    param(
        [string]$SourcePath,
        [string]$DestinationPath,
        [string]$Algorithm,
        [scriptblock]$OnProgress,
        [int]$BufferSize = 4194304
    )

    $Outcome = [pscustomobject]@{
        SourceHash = ""
        DestinationHash = ""
        BytesCopied = [long]0
        Verified = $false
        Status = ""
        Error = ""
    }

    $Hasher = $null
    $SourceStream = $null
    $DestinationStream = $null

    try {
        $Hasher = [System.Security.Cryptography.HashAlgorithm]::Create($Algorithm)

        $SourceStream = New-Object System.IO.FileStream(
            (Get-LongPath -Path $SourcePath),
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read,
            [System.IO.FileShare]::ReadWrite,
            $BufferSize,
            [System.IO.FileOptions]::SequentialScan)

        $DestinationStream = New-Object System.IO.FileStream(
            (Get-LongPath -Path $DestinationPath),
            [System.IO.FileMode]::Create,
            [System.IO.FileAccess]::Write,
            [System.IO.FileShare]::None,
            $BufferSize,
            [System.IO.FileOptions]::SequentialScan)

        $Buffer = New-Object byte[] $BufferSize
        $Read = 0

        while (($Read = $SourceStream.Read($Buffer, 0, $BufferSize)) -gt 0) {
            $DestinationStream.Write($Buffer, 0, $Read)
            [void]$Hasher.TransformBlock($Buffer, 0, $Read, $null, 0)
            $Outcome.BytesCopied += $Read

            if ($OnProgress) {
                & $OnProgress $Read
            }

            if ($Script:CancelRequested) {
                $Outcome.Status = "Cancelled"
                $Outcome.Error = "Cancelled part way through the file."
                break
            }
        }

        [void]$Hasher.TransformFinalBlock((New-Object byte[] 0), 0, 0)
        $Outcome.SourceHash = [System.BitConverter]::ToString($Hasher.Hash).Replace("-", "")

        $DestinationStream.Flush($true)
        $DestinationStream.Dispose()
        $DestinationStream = $null

        if ($Outcome.Status -eq "Cancelled") {
            try {
                [System.IO.File]::Delete((Get-LongPath -Path $DestinationPath))
            } catch {
                $Outcome.Error += " The partial destination file could not be removed: $($_.Exception.Message)"
            }

            return $Outcome
        }

        $Outcome.DestinationHash = Get-FileChecksum -Path $DestinationPath -Algorithm $Algorithm -BufferSize $BufferSize

        if ($Outcome.SourceHash -eq $Outcome.DestinationHash) {
            $Outcome.Verified = $true
            $Outcome.Status = "Verified"
        } else {
            $Outcome.Verified = $false
            $Outcome.Status = "CHECKSUM MISMATCH"
            $Outcome.Error = "Source and destination checksums do not match. The destination copy cannot be trusted."
        }

        # Keep the original timestamps so the archive still shows how old the data is.
        try {
            $SourceInfo = New-Object System.IO.FileInfo((Get-LongPath -Path $SourcePath))
            $DestinationInfo = New-Object System.IO.FileInfo((Get-LongPath -Path $DestinationPath))
            $DestinationInfo.CreationTimeUtc = $SourceInfo.CreationTimeUtc
            $DestinationInfo.LastWriteTimeUtc = $SourceInfo.LastWriteTimeUtc
            $DestinationInfo.LastAccessTimeUtc = $SourceInfo.LastAccessTimeUtc
        } catch {
            $Outcome.Error += " Timestamps could not be copied: $($_.Exception.Message)"
        }
    } catch {
        $Outcome.Status = "Failed"
        $Outcome.Error = $_.Exception.Message
    } finally {
        if ($SourceStream) { $SourceStream.Dispose() }
        if ($DestinationStream) { $DestinationStream.Dispose() }
        if ($Hasher) { $Hasher.Dispose() }
    }

    return $Outcome
}

# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

[System.Windows.Forms.Application]::EnableVisualStyles()

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "BRC Stale Data Archiver - copy and checksum verify"
$Form.Size = New-Object System.Drawing.Size(900, 800)
$Form.StartPosition = "CenterScreen"
$Form.MaximizeBox = $false
$Form.FormBorderStyle = "FixedDialog"

$LabelCsv = New-Object System.Windows.Forms.Label
$LabelCsv.Text = "Stale directory CSV (the output of BRC-StaleDirectoryFinder.ps1):"
$LabelCsv.Location = New-Object System.Drawing.Point(20, 20)
$LabelCsv.Size = New-Object System.Drawing.Size(500, 20)
$Form.Controls.Add($LabelCsv)

$TextCsv = New-Object System.Windows.Forms.TextBox
$TextCsv.Location = New-Object System.Drawing.Point(20, 45)
$TextCsv.Size = New-Object System.Drawing.Size(710, 24)
$Form.Controls.Add($TextCsv)

$ButtonCsvBrowse = New-Object System.Windows.Forms.Button
$ButtonCsvBrowse.Text = "Browse..."
$ButtonCsvBrowse.Location = New-Object System.Drawing.Point(745, 43)
$ButtonCsvBrowse.Size = New-Object System.Drawing.Size(105, 28)
$Form.Controls.Add($ButtonCsvBrowse)

$CheckOnlyStale = New-Object System.Windows.Forms.CheckBox
$CheckOnlyStale.Text = "Only use rows where IsStale = True"
$CheckOnlyStale.Location = New-Object System.Drawing.Point(20, 76)
$CheckOnlyStale.Size = New-Object System.Drawing.Size(320, 22)
$CheckOnlyStale.Checked = $true
$Form.Controls.Add($CheckOnlyStale)

$LabelDestination = New-Object System.Windows.Forms.Label
$LabelDestination.Text = "Destination folder / UNC share:"
$LabelDestination.Location = New-Object System.Drawing.Point(20, 108)
$LabelDestination.Size = New-Object System.Drawing.Size(500, 20)
$Form.Controls.Add($LabelDestination)

$TextDestination = New-Object System.Windows.Forms.TextBox
$TextDestination.Location = New-Object System.Drawing.Point(20, 133)
$TextDestination.Size = New-Object System.Drawing.Size(710, 24)
$Form.Controls.Add($TextDestination)

$ButtonDestinationBrowse = New-Object System.Windows.Forms.Button
$ButtonDestinationBrowse.Text = "Browse..."
$ButtonDestinationBrowse.Location = New-Object System.Drawing.Point(745, 131)
$ButtonDestinationBrowse.Size = New-Object System.Drawing.Size(105, 28)
$Form.Controls.Add($ButtonDestinationBrowse)

$LabelLayout = New-Object System.Windows.Forms.Label
$LabelLayout.Text = "Destination layout:"
$LabelLayout.Location = New-Object System.Drawing.Point(20, 172)
$LabelLayout.Size = New-Object System.Drawing.Size(120, 20)
$Form.Controls.Add($LabelLayout)

$ComboLayout = New-Object System.Windows.Forms.ComboBox
$ComboLayout.Location = New-Object System.Drawing.Point(145, 169)
$ComboLayout.Size = New-Object System.Drawing.Size(280, 24)
$ComboLayout.DropDownStyle = "DropDownList"
[void]$ComboLayout.Items.Add("Mirror full source path")
[void]$ComboLayout.Items.Add("Folder name only")
$ComboLayout.SelectedItem = "Mirror full source path"
$Form.Controls.Add($ComboLayout)

$LabelHash = New-Object System.Windows.Forms.Label
$LabelHash.Text = "Checksum:"
$LabelHash.Location = New-Object System.Drawing.Point(450, 172)
$LabelHash.Size = New-Object System.Drawing.Size(75, 20)
$Form.Controls.Add($LabelHash)

$ComboHash = New-Object System.Windows.Forms.ComboBox
$ComboHash.Location = New-Object System.Drawing.Point(530, 169)
$ComboHash.Size = New-Object System.Drawing.Size(120, 24)
$ComboHash.DropDownStyle = "DropDownList"
[void]$ComboHash.Items.Add("SHA256")
[void]$ComboHash.Items.Add("SHA512")
[void]$ComboHash.Items.Add("SHA1")
[void]$ComboHash.Items.Add("MD5")
$ComboHash.SelectedItem = "SHA256"
$Form.Controls.Add($ComboHash)

$LabelMargin = New-Object System.Windows.Forms.Label
$LabelMargin.Text = "Stop if free space would drop below (GB):"
$LabelMargin.Location = New-Object System.Drawing.Point(20, 208)
$LabelMargin.Size = New-Object System.Drawing.Size(265, 20)
$Form.Controls.Add($LabelMargin)

$NumericMargin = New-Object System.Windows.Forms.NumericUpDown
$NumericMargin.Location = New-Object System.Drawing.Point(290, 205)
$NumericMargin.Size = New-Object System.Drawing.Size(80, 24)
$NumericMargin.Minimum = 0
$NumericMargin.Maximum = 100000
$NumericMargin.Value = 20
$Form.Controls.Add($NumericMargin)

$CheckSkipExisting = New-Object System.Windows.Forms.CheckBox
$CheckSkipExisting.Text = "Skip files already at the destination (still checksum verified)"
$CheckSkipExisting.Location = New-Object System.Drawing.Point(390, 206)
$CheckSkipExisting.Size = New-Object System.Drawing.Size(460, 22)
$CheckSkipExisting.Checked = $true
$Form.Controls.Add($CheckSkipExisting)

$CheckOpenReport = New-Object System.Windows.Forms.CheckBox
$CheckOpenReport.Text = "Open the verification report when finished"
$CheckOpenReport.Location = New-Object System.Drawing.Point(20, 234)
$CheckOpenReport.Size = New-Object System.Drawing.Size(340, 22)
$CheckOpenReport.Checked = $true
$Form.Controls.Add($CheckOpenReport)

$ButtonAnalyse = New-Object System.Windows.Forms.Button
$ButtonAnalyse.Text = "Analyse only"
$ButtonAnalyse.Location = New-Object System.Drawing.Point(20, 266)
$ButtonAnalyse.Size = New-Object System.Drawing.Size(150, 32)
$Form.Controls.Add($ButtonAnalyse)

$ButtonStart = New-Object System.Windows.Forms.Button
$ButtonStart.Text = "Copy and verify"
$ButtonStart.Location = New-Object System.Drawing.Point(180, 266)
$ButtonStart.Size = New-Object System.Drawing.Size(150, 32)
$Form.Controls.Add($ButtonStart)

$ButtonCancel = New-Object System.Windows.Forms.Button
$ButtonCancel.Text = "Cancel"
$ButtonCancel.Location = New-Object System.Drawing.Point(340, 266)
$ButtonCancel.Size = New-Object System.Drawing.Size(110, 32)
$ButtonCancel.Enabled = $false
$Form.Controls.Add($ButtonCancel)

$ButtonClose = New-Object System.Windows.Forms.Button
$ButtonClose.Text = "Close"
$ButtonClose.Location = New-Object System.Drawing.Point(745, 266)
$ButtonClose.Size = New-Object System.Drawing.Size(105, 32)
$Form.Controls.Add($ButtonClose)

$LabelSummary = New-Object System.Windows.Forms.Label
$LabelSummary.Text = "Select a CSV and a destination, then press Analyse only to see what would be copied."
$LabelSummary.Location = New-Object System.Drawing.Point(20, 310)
$LabelSummary.Size = New-Object System.Drawing.Size(830, 44)
$Form.Controls.Add($LabelSummary)

$ProgressBar = New-Object System.Windows.Forms.ProgressBar
$ProgressBar.Location = New-Object System.Drawing.Point(20, 358)
$ProgressBar.Size = New-Object System.Drawing.Size(830, 22)
$ProgressBar.Style = "Blocks"
$Form.Controls.Add($ProgressBar)

$StatusLabel = New-Object System.Windows.Forms.Label
$StatusLabel.Text = "Idle."
$StatusLabel.Location = New-Object System.Drawing.Point(20, 386)
$StatusLabel.Size = New-Object System.Drawing.Size(830, 20)
$Form.Controls.Add($StatusLabel)

$LabelLog = New-Object System.Windows.Forms.Label
$LabelLog.Text = "Log:"
$LabelLog.Location = New-Object System.Drawing.Point(20, 412)
$LabelLog.Size = New-Object System.Drawing.Size(200, 20)
$Form.Controls.Add($LabelLog)

$TextLog = New-Object System.Windows.Forms.TextBox
$TextLog.Location = New-Object System.Drawing.Point(20, 434)
$TextLog.Size = New-Object System.Drawing.Size(830, 290)
$TextLog.Multiline = $true
$TextLog.ReadOnly = $true
$TextLog.ScrollBars = "Vertical"
$TextLog.Font = New-Object System.Drawing.Font("Consolas", 8.5)
$Form.Controls.Add($TextLog)

$LabelReport = New-Object System.Windows.Forms.Label
$LabelReport.Text = "The verification report CSV is written into the destination folder."
$LabelReport.Location = New-Object System.Drawing.Point(20, 732)
$LabelReport.Size = New-Object System.Drawing.Size(830, 20)
$Form.Controls.Add($LabelReport)

function Write-Log {
    param([string]$Message)

    $Line = "[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $Message
    $TextLog.AppendText($Line + [Environment]::NewLine)
    [System.Windows.Forms.Application]::DoEvents()
}

function Set-Status {
    param([string]$Message)

    $StatusLabel.Text = $Message
    [System.Windows.Forms.Application]::DoEvents()
}

function Set-ControlsEnabled {
    param([bool]$Enabled)

    $ButtonAnalyse.Enabled = $Enabled
    $ButtonStart.Enabled = $Enabled
    $ButtonClose.Enabled = $Enabled
    $ButtonCsvBrowse.Enabled = $Enabled
    $ButtonDestinationBrowse.Enabled = $Enabled
    $TextCsv.Enabled = $Enabled
    $TextDestination.Enabled = $Enabled
    $ComboLayout.Enabled = $Enabled
    $ComboHash.Enabled = $Enabled
    $NumericMargin.Enabled = $Enabled
    $CheckOnlyStale.Enabled = $Enabled
    $CheckSkipExisting.Enabled = $Enabled
    $CheckOpenReport.Enabled = $Enabled
    $ButtonCancel.Enabled = (-not $Enabled)
    [System.Windows.Forms.Application]::DoEvents()
}

function Show-Warning {
    param(
        [string]$Message,
        [string]$Title = "Cannot continue"
    )

    [System.Windows.Forms.MessageBox]::Show($Message, $Title, "OK", "Warning") | Out-Null
}

function Test-DestinationIsNetwork {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $false
    }

    if ($Path.StartsWith('\\')) {
        return $true
    }

    if ($Path -match '^([A-Za-z]):') {
        try {
            $Drive = New-Object System.IO.DriveInfo($Matches[1])
            return ($Drive.DriveType -eq [System.IO.DriveType]::Network)
        } catch {
            return $false
        }
    }

    return $false
}

# ---------------------------------------------------------------------------
# Analysis - work out the full file list and check the space BEFORE copying
# ---------------------------------------------------------------------------

function Invoke-Analysis {
    param([bool]$Quiet = $false)

    $CsvPath = $TextCsv.Text.Trim().Trim('"')
    $DestinationRoot = Get-NormalisedPath -Path $TextDestination.Text

    if ([string]::IsNullOrWhiteSpace($CsvPath)) {
        Show-Warning -Message "Select the CSV that BRC-StaleDirectoryFinder.ps1 produced." -Title "Missing CSV"
        return $null
    }

    if (-not (Test-FileExistsSafe -Path $CsvPath)) {
        Show-Warning -Message "This CSV file does not exist:`r`n$CsvPath" -Title "Missing CSV"
        return $null
    }

    if ([string]::IsNullOrWhiteSpace($DestinationRoot)) {
        Show-Warning -Message "Select the destination folder or UNC share to copy into." -Title "Missing destination"
        return $null
    }

    $Script:CancelRequested = $false
    $Script:IsRunning = $true
    Set-ControlsEnabled -Enabled $false
    $ProgressBar.Style = "Marquee"
    $ProgressBar.MarqueeAnimationSpeed = 30

    try {
        Write-Log "Reading $CsvPath"
        Set-Status "Reading the CSV..."

        $CsvResult = $null

        try {
            $CsvResult = Read-SourcePathsFromCsv -CsvPath $CsvPath -OnlyStaleRows $CheckOnlyStale.Checked -OnProgress {
                param($Phase, $Done, $Total)
                Set-Status ("{0}: {1} of {2}" -f $Phase, $Done, $Total)
            }
        } catch {
            Show-Warning -Message "The CSV could not be read:`r`n$($_.Exception.Message)" -Title "CSV error"
            return $null
        }

        foreach ($CsvError in $CsvResult.Errors) {
            Write-Log "CSV: $CsvError"
        }

        if ($CsvResult.Errors.Count -gt 0) {
            Show-Warning -Message ($CsvResult.Errors -join "`r`n") -Title "CSV error"
            return $null
        }

        Write-Log ("CSV rows: {0}. Path column: {1}." -f $CsvResult.RowCount, $CsvResult.PathColumn)

        if ($CsvResult.SkippedNotStale -gt 0) {
            Write-Log ("Rows skipped because IsStale was not True: {0}" -f $CsvResult.SkippedNotStale)
        }

        if ($CsvResult.SkippedNested -gt 0) {
            Write-Log ("Paths skipped because a parent folder is already in the list: {0}" -f $CsvResult.SkippedNested)
        }

        if ($CsvResult.SkippedMissing -gt 0) {
            Write-Log ("Paths in the CSV that no longer exist: {0}" -f $CsvResult.SkippedMissing)
            foreach ($Missing in $CsvResult.MissingPaths) {
                Write-Log "  missing: $Missing"
            }
        }

        if ($CsvResult.Paths.Count -eq 0) {
            Show-Warning -Message "No usable source paths were found in the CSV." -Title "Nothing to copy"
            return $null
        }

        if ($Script:CancelRequested) {
            Write-Log "Cancelled while reading the CSV."
            Set-Status "Cancelled."
            return $null
        }

        Write-Log ("Source folders to archive: {0}" -f $CsvResult.Paths.Count)

        # A destination inside a source, or a source inside the destination,
        # would make the copy feed itself.
        Set-Status "Checking the destination does not overlap the source..."
        $OverlapChecked = 0

        foreach ($SourceRoot in $CsvResult.Paths) {
            $OverlapChecked++

            if (($OverlapChecked % 500) -eq 0) {
                [System.Windows.Forms.Application]::DoEvents()
            }

            if (Test-PathIsInside -ChildPath $DestinationRoot -ParentPath $SourceRoot) {
                Show-Warning -Message "The destination sits inside a source folder, which would copy files into themselves:`r`n`r`nSource: $SourceRoot`r`nDestination: $DestinationRoot" -Title "Invalid destination"
                return $null
            }

            if (Test-PathIsInside -ChildPath $SourceRoot -ParentPath $DestinationRoot) {
                Show-Warning -Message "A source folder sits inside the destination:`r`n`r`nSource: $SourceRoot`r`nDestination: $DestinationRoot" -Title "Invalid destination"
                return $null
            }
        }

        $Items = New-Object System.Collections.Generic.List[object]
        $Folders = New-Object System.Collections.Generic.List[string]
        $WalkErrors = New-Object System.Collections.Generic.List[string]
        $UsedRoots = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        $TotalBytes = [double]0
        $RootIndex = 0

        foreach ($SourceRoot in $CsvResult.Paths) {
            if ($Script:CancelRequested) {
                break
            }

            $RootIndex++
            Set-Status ("Listing files: {0} of {1} - {2}" -f $RootIndex, $CsvResult.Paths.Count, $SourceRoot)

            $DestinationForRoot = Get-DestinationRootPath -SourceRoot $SourceRoot -DestinationRoot $DestinationRoot -LayoutMode ([string]$ComboLayout.SelectedItem) -UsedRoots $UsedRoots

            $RootFolders = New-Object System.Collections.Generic.List[string]

            $Files = Get-FilesUnderPath -RootPath $SourceRoot -Errors $WalkErrors -Directories $RootFolders -OnProgress {
                param($Count)
                Set-Status ("Listing files: {0} of {1} - {2} ({3} files so far)" -f $RootIndex, $CsvResult.Paths.Count, $SourceRoot, $Count)
            }

            $Folders.Add($DestinationForRoot)

            foreach ($RelativeFolder in $RootFolders) {
                $Folders.Add((Join-Path -Path $DestinationForRoot -ChildPath $RelativeFolder))
            }

            $RootBytes = [double]0
            $Built = 0

            foreach ($File in $Files) {
                $Items.Add([pscustomobject]@{
                    SourceRoot = $SourceRoot
                    SourceFile = $File.FullName
                    DestinationFile = (Join-Path -Path $DestinationForRoot -ChildPath $File.RelativePath)
                    Length = $File.Length
                })

                $TotalBytes += $File.Length
                $RootBytes += $File.Length
                $Built++

                if (($Built % 5000) -eq 0) {
                    Set-Status ("Building the copy list: {0} of {1} files under {2}" -f $Built, $Files.Count, $SourceRoot)

                    if ($Script:CancelRequested) {
                        break
                    }
                }
            }

            Write-Log ("{0} -> {1} ({2} files, {3})" -f $SourceRoot, $DestinationForRoot, $Files.Count, (Format-Bytes $RootBytes))
        }

        if ($Script:CancelRequested) {
            Write-Log "Analysis cancelled."
            Set-Status "Cancelled."
            return $null
        }

        foreach ($WalkError in $WalkErrors) {
            Write-Log "Enumeration problem: $WalkError"
        }

        if ($Items.Count -eq 0) {
            Show-Warning -Message "The listed folders contain no files to copy." -Title "Nothing to copy"
            return $null
        }

        $MarginBytes = [double]$NumericMargin.Value * 1GB
        $SpaceCheck = Test-FreeSpaceForCopy -DestinationPath $DestinationRoot -RequiredBytes $TotalBytes -MarginBytes $MarginBytes

        $SummaryText = "{0} files, {1} to copy from {2} folder(s).`r`nDestination: {3}`r`n{4}" -f `
            $Items.Count,
            (Format-Bytes $TotalBytes),
            $CsvResult.Paths.Count,
            $DestinationRoot,
            $SpaceCheck.Reason

        $LabelSummary.Text = $SummaryText

        Write-Log ("Total to copy: {0} files, {1}" -f $Items.Count, (Format-Bytes $TotalBytes))

        if ($SpaceCheck.Succeeded) {
            Write-Log ("Destination free space: {0} of {1}" -f (Format-Bytes $SpaceCheck.FreeBytes), (Format-Bytes $SpaceCheck.TotalBytes))
        }

        Write-Log ("Free space check: {0}" -f $SpaceCheck.Reason)

        return [pscustomobject]@{
            Items = $Items
            Folders = $Folders
            SourceRoots = $CsvResult.Paths
            DestinationRoot = $DestinationRoot
            TotalBytes = $TotalBytes
            TotalFiles = $Items.Count
            SpaceCheck = $SpaceCheck
            MarginBytes = $MarginBytes
            WalkErrors = $WalkErrors
        }
    } finally {
        $Script:IsRunning = $false
        $ProgressBar.MarqueeAnimationSpeed = 0
        $ProgressBar.Style = "Blocks"
        Set-ControlsEnabled -Enabled $true

        if (-not $Quiet) {
            Set-Status "Analysis complete."
        }
    }
}

# ---------------------------------------------------------------------------
# The copy and verify run
# ---------------------------------------------------------------------------

function Start-CopyRun {
    $Plan = Invoke-Analysis -Quiet $true

    if ($null -eq $Plan) {
        return
    }

    $Algorithm = [string]$ComboHash.SelectedItem
    $DestinationRoot = $Plan.DestinationRoot
    $SpaceCheck = $Plan.SpaceCheck

    # ----- Free space gate. This is the hard stop. -----
    if (-not $SpaceCheck.Succeeded) {
        $Answer = [System.Windows.Forms.MessageBox]::Show(
            ("Free space on the destination could not be measured, so the {0} safety check cannot be applied.`r`n`r`n{1}`r`n`r`nContinue anyway? The copy could fill the destination." -f (Format-Bytes $Plan.MarginBytes), $SpaceCheck.Reason),
            "Free space unknown",
            "YesNo",
            "Warning",
            "Button2")

        if ($Answer -ne "Yes") {
            Write-Log "Stopped: free space could not be confirmed."
            Set-Status "Stopped - free space unknown."
            return
        }
    } elseif (-not $SpaceCheck.IsSufficient) {
        Show-Warning -Message ("Not enough free space at the destination. Nothing has been copied.`r`n`r`n{0}`r`n`r`nDestination: {1}`r`nFree now: {2}`r`nNeeded: {3}`r`nMinimum to keep free: {4}" -f `
                $SpaceCheck.Reason,
                $DestinationRoot,
                (Format-Bytes $SpaceCheck.FreeBytes),
                (Format-Bytes $Plan.TotalBytes),
                (Format-Bytes $Plan.MarginBytes)) -Title "Not enough free space"

        Write-Log "STOPPED: $($SpaceCheck.Reason)"
        Set-Status "Stopped - not enough free space."
        return
    }

    # ----- Confirmation 1: what is about to happen -----
    $ConfirmText = "About to copy:`r`n`r`n  {0} files ({1})`r`n  from {2} folder(s) listed in the CSV`r`n  to {3}`r`n`r`nEvery file is checksum verified with {4} after it is written.`r`n`r`nNothing at the source is changed or deleted.`r`n`r`nIMPORTANT: copied files do NOT keep their NTFS permissions. They inherit the permissions of the destination, so anyone who can read the destination will be able to read this data. Confirm the destination permissions are correct before continuing.`r`n`r`nProceed?" -f `
        $Plan.TotalFiles,
        (Format-Bytes $Plan.TotalBytes),
        $Plan.SourceRoots.Count,
        $DestinationRoot,
        $Algorithm

    if ([System.Windows.Forms.MessageBox]::Show($ConfirmText, "Confirm copy", "YesNo", "Question", "Button2") -ne "Yes") {
        Write-Log "Copy cancelled at the confirmation prompt."
        Set-Status "Cancelled."
        return
    }

    # ----- Confirmation 2: writing out to a network share -----
    if (Test-DestinationIsNetwork -Path $DestinationRoot) {
        $NetworkText = "The destination is a network share:`r`n`r`n  {0}`r`n`r`nData written there leaves this server and is governed by that share's permissions, not the source permissions. If the share is reachable by other teams, or published externally, this data becomes readable by them.`r`n`r`nAre you sure you want to write {1} of archived data to this share?" -f `
            $DestinationRoot,
            (Format-Bytes $Plan.TotalBytes)

        if ([System.Windows.Forms.MessageBox]::Show($NetworkText, "Confirm network destination", "YesNo", "Warning", "Button2") -ne "Yes") {
            Write-Log "Copy cancelled at the network destination prompt."
            Set-Status "Cancelled."
            return
        }
    }

    $Script:CancelRequested = $false
    $Script:IsRunning = $true
    Set-ControlsEnabled -Enabled $false

    $Writer = $null
    $ReportPath = ""
    $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $UiStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    $Copied = 0
    $VerifiedCount = 0
    $Skipped = 0
    $Mismatched = 0
    $Failed = 0
    $BytesDone = [double]0
    $BytesSinceSpaceCheck = [double]0
    $FilesSinceSpaceCheck = 0
    $AbortedForSpace = $false

    try {
        New-DirectoryTree -Path $DestinationRoot

        $ReportName = "BRC-CopyVerify_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss")
        $ReportPath = Join-Path -Path $DestinationRoot -ChildPath $ReportName

        try {
            $Writer = New-Object System.IO.StreamWriter((Get-LongPath -Path $ReportPath), $false, [System.Text.Encoding]::UTF8)
        } catch {
            $ReportPath = Join-Path -Path ([Environment]::GetFolderPath("Desktop")) -ChildPath $ReportName
            $Writer = New-Object System.IO.StreamWriter((Get-LongPath -Path $ReportPath), $false, [System.Text.Encoding]::UTF8)
            Write-Log "The report could not be written to the destination, using $ReportPath instead."
        }

        Write-CsvRow -Writer $Writer -Values @(
            "SourceRoot",
            "SourceFile",
            "DestinationFile",
            "SizeBytes",
            "Algorithm",
            "SourceChecksum",
            "DestinationChecksum",
            "Verified",
            "Status",
            "CompletedUtc",
            "Error"
        )

        $LabelReport.Text = "Verification report: $ReportPath"
        Write-Log "Verification report: $ReportPath"

        # Recreate the folder skeleton first so folders that hold no files are
        # still represented in the archive.
        Set-Status "Creating the destination folder structure..."
        $FoldersCreated = 0

        foreach ($Folder in $Plan.Folders) {
            if ($Script:CancelRequested) {
                break
            }

            try {
                New-DirectoryTree -Path $Folder
                $FoldersCreated++
            } catch {
                Write-Log "Could not create $Folder - $($_.Exception.Message)"
            }
        }

        Write-Log ("Destination folders ready: {0}" -f $FoldersCreated)
        Write-Log ("Starting copy of {0} files ({1}) using {2}." -f $Plan.TotalFiles, (Format-Bytes $Plan.TotalBytes), $Algorithm)

        if ($Algorithm -eq "MD5" -or $Algorithm -eq "SHA1") {
            Write-Log "NOTE: $Algorithm detects copy corruption but is not collision resistant. Use SHA256 if this report has to stand up as evidence the data was not altered."
        }

        $ProgressBar.Style = "Blocks"
        $ProgressBar.Minimum = 0
        $ProgressBar.Maximum = 1000
        $ProgressBar.Value = 0

        $Index = 0

        foreach ($Item in $Plan.Items) {
            if ($Script:CancelRequested) {
                break
            }

            $Index++
            $Status = ""
            $ErrorText = ""
            $SourceHash = ""
            $DestinationHash = ""
            $Verified = $false

            Set-Status ("File {0} of {1} - {2} of {3} - {4}" -f `
                $Index,
                $Plan.TotalFiles,
                (Format-Bytes $BytesDone),
                (Format-Bytes $Plan.TotalBytes),
                (Split-Path -Path $Item.SourceFile -Leaf))

            try {
                $ParentFolder = [System.IO.Path]::GetDirectoryName($Item.DestinationFile)
                New-DirectoryTree -Path $ParentFolder

                $AlreadyThere = $false

                if ($CheckSkipExisting.Checked -and (Test-FileExistsSafe -Path $Item.DestinationFile)) {
                    $ExistingInfo = New-Object System.IO.FileInfo((Get-LongPath -Path $Item.DestinationFile))

                    if ([long]$ExistingInfo.Length -eq [long]$Item.Length) {
                        $SourceHash = Get-FileChecksum -Path $Item.SourceFile -Algorithm $Algorithm
                        $DestinationHash = Get-FileChecksum -Path $Item.DestinationFile -Algorithm $Algorithm

                        if ($SourceHash -eq $DestinationHash) {
                            $AlreadyThere = $true
                            $Verified = $true
                            $Status = "Skipped - already verified"
                            $Skipped++
                            $VerifiedCount++
                        }
                    }
                }

                if (-not $AlreadyThere) {
                    $Outcome = Copy-FileWithChecksum -SourcePath $Item.SourceFile -DestinationPath $Item.DestinationFile -Algorithm $Algorithm -OnProgress {
                        param($Bytes)

                        if ($UiStopwatch.ElapsedMilliseconds -ge 250) {
                            $UiStopwatch.Restart()
                            [System.Windows.Forms.Application]::DoEvents()
                        }
                    }

                    $SourceHash = $Outcome.SourceHash
                    $DestinationHash = $Outcome.DestinationHash
                    $Verified = $Outcome.Verified
                    $Status = $Outcome.Status
                    $ErrorText = $Outcome.Error

                    if ($Outcome.Verified) {
                        $Copied++
                        $VerifiedCount++
                    } elseif ($Outcome.Status -eq "CHECKSUM MISMATCH") {
                        $Mismatched++
                        Write-Log "CHECKSUM MISMATCH: $($Item.SourceFile)"
                    } elseif ($Outcome.Status -eq "Cancelled") {
                        # counted as nothing - the partial file was removed
                    } else {
                        $Failed++
                        Write-Log "FAILED: $($Item.SourceFile) - $($Outcome.Error)"
                    }
                }
            } catch {
                $Status = "Failed"
                $ErrorText = $_.Exception.Message
                $Failed++
                Write-Log "FAILED: $($Item.SourceFile) - $ErrorText"
            }

            Write-CsvRow -Writer $Writer -Values @(
                $Item.SourceRoot,
                $Item.SourceFile,
                $Item.DestinationFile,
                [string]$Item.Length,
                $Algorithm,
                $SourceHash,
                $DestinationHash,
                [string]$Verified,
                $Status,
                (Format-DateForCsv -DateValue ([datetime]::UtcNow)),
                $ErrorText
            )

            $Writer.Flush()

            $BytesDone += $Item.Length
            $BytesSinceSpaceCheck += $Item.Length
            $FilesSinceSpaceCheck++

            if ($Plan.TotalBytes -gt 0) {
                $Fraction = $BytesDone / $Plan.TotalBytes
                $ProgressBar.Value = [Math]::Min(1000, [Math]::Max(0, [int]($Fraction * 1000)))
            }

            # Re-check free space while running. Something else writing to the
            # same share can eat the headroom after the pre-flight check passed.
            if ($FilesSinceSpaceCheck -ge 100 -or $BytesSinceSpaceCheck -ge 1GB) {
                $FilesSinceSpaceCheck = 0
                $BytesSinceSpaceCheck = 0

                $Live = Get-FreeSpaceInfo -Path $DestinationRoot

                if ($Live.Succeeded -and $Live.FreeBytes -lt $Plan.MarginBytes) {
                    $AbortedForSpace = $true
                    Write-Log ("STOPPED: destination free space fell to {0}, below the {1} minimum." -f (Format-Bytes $Live.FreeBytes), (Format-Bytes $Plan.MarginBytes))
                    break
                }
            }

            [System.Windows.Forms.Application]::DoEvents()
        }

        $Stopwatch.Stop()

        $Remaining = $Plan.TotalFiles - $Index
        $Outstanding = if ($Remaining -gt 0) { "`r`nFiles not attempted: $Remaining" } else { "" }

        $SummaryLines = @(
            ("Files copied and verified: {0}" -f $Copied),
            ("Files skipped, already verified at the destination: {0}" -f $Skipped),
            ("Checksum mismatches: {0}" -f $Mismatched),
            ("Failures: {0}" -f $Failed),
            ("Data written: {0}" -f (Format-Bytes $BytesDone)),
            ("Elapsed: {0:hh\:mm\:ss}" -f $Stopwatch.Elapsed)
        )

        foreach ($Line in $SummaryLines) {
            Write-Log $Line
        }

        $LabelSummary.Text = ("{0} verified, {1} mismatched, {2} failed, {3} written in {4:hh\:mm\:ss}." -f `
            $VerifiedCount, $Mismatched, $Failed, (Format-Bytes $BytesDone), $Stopwatch.Elapsed)

        $Title = "Copy complete"
        $Icon = "Information"

        if ($AbortedForSpace) {
            $Title = "Stopped - free space"
            $Icon = "Warning"
        } elseif ($Script:CancelRequested) {
            $Title = "Cancelled"
            $Icon = "Warning"
        } elseif ($Mismatched -gt 0 -or $Failed -gt 0) {
            $Title = "Complete with problems"
            $Icon = "Warning"
        }

        Set-Status $Title

        [System.Windows.Forms.MessageBox]::Show(
            (($SummaryLines -join "`r`n") + $Outstanding + "`r`n`r`nVerification report:`r`n$ReportPath"),
            $Title,
            "OK",
            $Icon) | Out-Null

        if ($CheckOpenReport.Checked -and (Test-FileExistsSafe -Path $ReportPath)) {
            try {
                Start-Process -FilePath $ReportPath | Out-Null
            } catch {
                Write-Log "The report could not be opened: $($_.Exception.Message)"
            }
        }
    } catch {
        Write-Log "Run failed: $($_.Exception.Message)"
        Show-Warning -Message ("The run failed:`r`n`r`n{0}" -f $_.Exception.Message) -Title "Copy failed"
        Set-Status "Failed."
    } finally {
        if ($Writer) {
            $Writer.Flush()
            $Writer.Close()
            $Writer.Dispose()
        }

        $Script:IsRunning = $false
        Set-ControlsEnabled -Enabled $true
    }
}

# ---------------------------------------------------------------------------
# Events
# ---------------------------------------------------------------------------

$ButtonCsvBrowse.Add_Click({
    $Dialog = New-Object System.Windows.Forms.OpenFileDialog
    $Dialog.Title = "Select the stale directory CSV"
    $Dialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"

    if ($Dialog.ShowDialog() -eq "OK") {
        $TextCsv.Text = $Dialog.FileName
    }
})

$ButtonDestinationBrowse.Add_Click({
    $Dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $Dialog.Description = "Select the destination folder or UNC share"
    $Dialog.ShowNewFolderButton = $true

    if ($Dialog.ShowDialog() -eq "OK") {
        $TextDestination.Text = $Dialog.SelectedPath
    }
})

$ButtonAnalyse.Add_Click({
    $TextLog.Clear()
    [void](Invoke-Analysis)
})

$ButtonStart.Add_Click({
    $TextLog.Clear()
    Start-CopyRun
})

$ButtonCancel.Add_Click({
    $Script:CancelRequested = $true
    Set-Status "Cancelling after the current file..."
    Write-Log "Cancel requested."
})

$ButtonClose.Add_Click({
    $Form.Close()
})

$Form.Add_FormClosing({
    param($Sender, $EventArgs)

    if ($Script:IsRunning) {
        $EventArgs.Cancel = $true
        $Script:CancelRequested = $true
        Set-Status "Cancelling after the current file..."
    }
})

[void]$Form.ShowDialog()
$Form.Dispose()
