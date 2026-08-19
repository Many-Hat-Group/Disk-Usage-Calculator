<#
BRC Folder Archiver
Windows Server 2019 / Windows PowerShell 5.1

What it does:
- Opens a GUI
- Lets you pick a source (live data) root and an archive destination root
- Finds folders that have not been modified within a chosen age threshold
  (days / weeks / months / years)
- Copies each stale folder to the archive, verifies the copy, then removes the
  source folder
- Leaves a .txt placeholder where the folder used to be, containing:
    * a plain-English note saying the content has been archived
    * who to contact
    * a unique 24 character reference code
    * a full breakdown of every folder, subfolder and file that was archived
- Writes a matching <ReferenceCode>.txt on the archive side, whose file name and
  file contents are both the same 24 character code, so I.T. can find the
  archived data with a single search for that code
- Exports a CSV of everything archived, with the original path and the new path
- Appends the same rows to a master index CSV in the archive root

Safety model (read this before running it live):
- "Preview only" is ON by default. Nothing is copied, moved or deleted in
  preview mode, but the CSV is still produced so you can review the plan.
- A live run requires typing ARCHIVE into a confirmation box.
- The default engine copies, verifies (file-for-file names and sizes, optional
  SHA-256), and only then deletes the source. A folder whose verification fails
  is never deleted.
- A folder that produced any read error during inventory is skipped, not moved.
- The archive root may not sit inside the source root, and vice versa.

Notes:
- Sizes are logical file lengths and use the same binary KB/MB/GB convention as
  the other BRC tools (1 MB = 1,048,576 bytes).
- Moving data does not carry the source share's inherited NTFS permissions with
  it unless you tick the robocopy permission option. Secure the archive share
  itself; see the README.
- Data Deduplication stores optimized files as reparse points. This tool reads
  the reparse tag rather than the attribute, so deduplicated and compressed
  files are archived as ordinary files; only junctions, symlinks, mount points
  and cloud/HSM tiered stubs stop a folder. Note that copying a deduplicated
  file rehydrates it, so the archive needs room for the logical size.
- This tool changes data. Test it against a copy of a folder tree first.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:CancelRequested = $false
$script:IssuedCodes = New-Object 'System.Collections.Generic.HashSet[string]'
$script:Sep = [System.IO.Path]::DirectorySeparatorChar

# ---------------------------------------------------------------------------
# CSV helpers
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
        [Nullable[datetime]]$DateValue
    )

    if ($null -eq $DateValue) {
        return ""
    }

    return ([datetime]$DateValue).ToString("yyyy-MM-dd HH:mm:ss")
}

function Get-DaysSince {
    param(
        [AllowNull()]
        [Nullable[datetime]]$DateValue
    )

    if ($null -eq $DateValue) {
        return ""
    }

    return [math]::Round(((Get-Date) - ([datetime]$DateValue)).TotalDays, 2)
}

function Format-Plural {
    param(
        [long]$Count,
        [string]$Singular,
        [string]$Plural
    )

    if ($Count -eq 1) {
        return ("{0} {1}" -f $Count, $Singular)
    }

    return ("{0} {1}" -f $Count, $Plural)
}

function Format-SizeFriendly {
    param(
        [long]$Bytes
    )

    if ($Bytes -ge 1TB) { return ("{0:N2} TB" -f ($Bytes / 1TB)) }
    if ($Bytes -ge 1GB) { return ("{0:N2} GB" -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ("{0:N2} MB" -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ("{0:N2} KB" -f ($Bytes / 1KB)) }

    return ("{0} bytes" -f $Bytes)
}

# ---------------------------------------------------------------------------
# Path helpers
# ---------------------------------------------------------------------------

function Get-NormalisedPath {
    param(
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return ""
    }

    try {
        $Full = [System.IO.Path]::GetFullPath($Path)
    } catch {
        $Full = $Path
    }

    if ($Full.Length -gt 3) {
        $Full = $Full.TrimEnd($script:Sep)
    }

    return $Full
}

function Get-RelativePath {
    param(
        [string]$RootPath,
        [string]$CurrentPath
    )

    $Root = Get-NormalisedPath -Path $RootPath
    $Current = Get-NormalisedPath -Path $CurrentPath

    if ($Current.Length -le $Root.Length) {
        return ""
    }

    return $Current.Substring($Root.Length).TrimStart($script:Sep)
}

function Get-DepthFromRoot {
    param(
        [string]$RootPath,
        [string]$CurrentPath
    )

    $Relative = Get-RelativePath -RootPath $RootPath -CurrentPath $CurrentPath

    if ([string]::IsNullOrWhiteSpace($Relative)) {
        return 0
    }

    return (($Relative -split [regex]::Escape($script:Sep)).Count)
}

function Test-PathIsUnder {
    param(
        [string]$ParentPath,
        [string]$ChildPath
    )

    $Parent = Get-NormalisedPath -Path $ParentPath
    $Child = Get-NormalisedPath -Path $ChildPath

    if ([string]::IsNullOrWhiteSpace($Parent) -or [string]::IsNullOrWhiteSpace($Child)) {
        return $false
    }

    if ($Child -ieq $Parent) {
        return $true
    }

    return $Child.StartsWith($Parent.TrimEnd($script:Sep) + $script:Sep, [System.StringComparison]::OrdinalIgnoreCase)
}

function Get-SafeFileNameFragment {
    param(
        [string]$Text
    )

    if ($null -eq $Text) {
        return ""
    }

    $Invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $Builder = New-Object System.Text.StringBuilder

    foreach ($Char in $Text.ToCharArray()) {
        if ($Invalid -contains $Char) {
            [void]$Builder.Append('_')
        } else {
            [void]$Builder.Append($Char)
        }
    }

    return $Builder.ToString().Trim()
}

function Get-NonClashingPath {
    param(
        [string]$DesiredPath
    )

    if (-not (Test-Path -LiteralPath $DesiredPath)) {
        return $DesiredPath
    }

    $Parent = Split-Path -Path $DesiredPath -Parent
    $Leaf = Split-Path -Path $DesiredPath -Leaf
    $Extension = [System.IO.Path]::GetExtension($Leaf)
    $BaseName = [System.IO.Path]::GetFileNameWithoutExtension($Leaf)

    for ($Index = 1; $Index -le 9999; $Index++) {
        $Candidate = Join-Path -Path $Parent -ChildPath ("{0}_{1}{2}" -f $BaseName, $Index, $Extension)

        if (-not (Test-Path -LiteralPath $Candidate)) {
            return $Candidate
        }
    }

    throw "Could not find a free name for '$DesiredPath' after 9999 attempts."
}

# ---------------------------------------------------------------------------
# Threshold and reference code
# ---------------------------------------------------------------------------

function Get-CutoffDate {
    param(
        [int]$ThresholdValue,
        [string]$ThresholdUnit
    )

    $Now = Get-Date

    switch ($ThresholdUnit) {
        "Days"   { return $Now.AddDays(-1 * $ThresholdValue) }
        "Weeks"  { return $Now.AddDays(-7 * $ThresholdValue) }
        "Months" { return $Now.AddMonths(-1 * $ThresholdValue) }
        "Years"  { return $Now.AddYears(-1 * $ThresholdValue) }
        default  { return $Now.AddDays(-1 * $ThresholdValue) }
    }
}

function New-ArchiveReferenceCode {
    # 24 characters from a 32 character unambiguous alphabet (no I, O, 0, 1),
    # drawn from the cryptographic RNG. That is 120 bits of entropy, so a
    # collision across any realistic number of archived folders is not a
    # practical concern. Collisions are still checked for, per run and against
    # the archive root, before a code is used.
    $Alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"
    $Rng = New-Object System.Security.Cryptography.RNGCryptoServiceProvider

    try {
        for ($Attempt = 1; $Attempt -le 50; $Attempt++) {
            $Bytes = New-Object 'byte[]' 24
            $Rng.GetBytes($Bytes)

            $Builder = New-Object System.Text.StringBuilder

            foreach ($Byte in $Bytes) {
                [void]$Builder.Append($Alphabet[$Byte % $Alphabet.Length])
            }

            $Code = $Builder.ToString()

            if (-not $script:IssuedCodes.Contains($Code)) {
                [void]$script:IssuedCodes.Add($Code)
                return $Code
            }
        }

        throw "Could not generate a unique 24 character reference code."
    } finally {
        $Rng.Dispose()
    }
}

function Register-ExistingReferenceCodes {
    # Seeds the in-memory code set from the archive root's master index so that
    # codes stay unique across separate runs, not just within one run.
    param(
        [string]$IndexCsvPath
    )

    if ([string]::IsNullOrWhiteSpace($IndexCsvPath)) {
        return
    }

    if (-not (Test-Path -LiteralPath $IndexCsvPath -PathType Leaf)) {
        return
    }

    try {
        foreach ($Row in (Import-Csv -LiteralPath $IndexCsvPath)) {
            if (-not [string]::IsNullOrWhiteSpace($Row.ReferenceCode)) {
                [void]$script:IssuedCodes.Add($Row.ReferenceCode)
            }
        }
    } catch {
        # A damaged or locked index must not stop an archive run. Uniqueness is
        # still protected by the 120 bits of entropy in each code.
    }
}

# ---------------------------------------------------------------------------
# Reparse points
# ---------------------------------------------------------------------------
#
# "Reparse point" covers two very different things, and the difference decides
# whether a folder is safe to archive:
#
#   A link  - a junction, symbolic link or volume mount point. It stands in for
#             something somewhere else, possibly outside the tree, possibly on
#             another server. Copying it duplicates data nobody asked to
#             archive, and deleting it risks reaching through to the target.
#             These block an archive.
#
#   A stub  - the file really is here, it is just stored differently. Data
#             Deduplication, Windows Overlay compression and Single Instance
#             Storage all work this way. Reads return the real bytes at local
#             speed, so these are ordinary files as far as archiving goes.
#             On a server with dedup enabled this is most of the volume.
#
# Tiered stubs (HSM, Azure File Sync, OneDrive placeholders) sit between the
# two: the data is real but lives elsewhere, and reading it pulls it back,
# which can be slow and can cost money. They get their own switch.
#
# The tag is read with FindFirstFileW, which returns it in dwReserved0 without
# opening the file. If that is unavailable for any reason - Constrained
# Language Mode, a blocked Add-Type, a path that will not enumerate - every
# reparse point is treated as a link, which is the conservative answer and
# matches how this tool behaved before tags were read at all.

$script:ReparseSurrogateBit = [System.Convert]::ToUInt32('20000000', 16)

# The data is here; only its on-disk representation differs.
$script:ReparseTransparentTags = @{
    '80000013' = 'Data Deduplication'
    '80000017' = 'Windows Overlay compression'
    '80000008' = 'Windows Image file'
    '80000007' = 'Single Instance Storage'
}

# The data is real but elsewhere, and reading it recalls it.
$script:ReparseTieredTags = @{
    'C0000004' = 'Hierarchical Storage Management'
    '80000006' = 'Hierarchical Storage Management 2'
    '8000001E' = 'Azure File Sync cloud tiering'
    '80000021' = 'OneDrive placeholder (legacy)'
}

# The item points at something else.
$script:ReparseLinkTags = @{
    'A0000003' = 'Junction or volume mount point'
    'A000000C' = 'Symbolic link'
    '80000014' = 'NFS special file'
    '8000001B' = 'App execution alias'
}

$script:ReparseReaderAvailable = $false

if (-not ('BRC.Native.ReparseReader' -as [type])) {
    try {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

namespace BRC.Native
{
    public static class ReparseReader
    {
        private const int MAX_PATH = 260;
        private const uint FILE_ATTRIBUTE_REPARSE_POINT = 0x400;

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct WIN32_FIND_DATAW
        {
            public uint dwFileAttributes;
            public System.Runtime.InteropServices.ComTypes.FILETIME ftCreationTime;
            public System.Runtime.InteropServices.ComTypes.FILETIME ftLastAccessTime;
            public System.Runtime.InteropServices.ComTypes.FILETIME ftLastWriteTime;
            public uint nFileSizeHigh;
            public uint nFileSizeLow;
            public uint dwReserved0;
            public uint dwReserved1;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = MAX_PATH)]
            public string cFileName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 14)]
            public string cAlternateFileName;
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr FindFirstFileW(string lpFileName, out WIN32_FIND_DATAW lpFindFileData);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool FindClose(IntPtr hFindFile);

        // 0            = not a reparse point
        // uint.MaxValue = the item could not be examined
        // anything else = the reparse tag
        public static uint GetReparseTag(string path)
        {
            WIN32_FIND_DATAW data;
            IntPtr handle = FindFirstFileW(path, out data);

            if (handle == new IntPtr(-1))
            {
                return uint.MaxValue;
            }

            try
            {
                if ((data.dwFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) == 0)
                {
                    return 0;
                }

                return data.dwReserved0;
            }
            finally
            {
                FindClose(handle);
            }
        }
    }
}
'@
        $script:ReparseReaderAvailable = $true
    } catch {
        # Left unavailable on purpose. Get-PathReparseInfo falls back to
        # treating every reparse point as a link.
    }
} else {
    $script:ReparseReaderAvailable = $true
}

function ConvertTo-ExtendedLengthPath {
    param(
        [string]$Path
    )

    # FindFirstFileW stops at 260 characters unless the path is prefixed.
    if ($Path.Length -lt 240) {
        return $Path
    }

    if ($Path.StartsWith('\\?\')) {
        return $Path
    }

    if ($Path.StartsWith('\\')) {
        return '\\?\UNC\' + $Path.Substring(2)
    }

    if ($Path.Length -ge 2 -and $Path[1] -eq ':') {
        return '\\?\' + $Path
    }

    return $Path
}

function Get-ReparseTagKind {
    param(
        [uint32]$Tag
    )

    if ($Tag -eq 0) {
        return "None"
    }

    $Hex = '{0:X8}' -f $Tag

    if ($script:ReparseTransparentTags.ContainsKey($Hex)) { return "Transparent" }
    if ($script:ReparseTieredTags.ContainsKey($Hex))      { return "Tiered" }
    if ($script:ReparseLinkTags.ContainsKey($Hex))        { return "Link" }

    # Cloud files placeholders occupy the range 9000_01A, where the fifth
    # nibble identifies the sync provider.
    if ($Hex.Substring(0, 4) -eq '9000' -and $Hex.Substring(5) -eq '01A') {
        return "Tiered"
    }

    # Documented rule: the name surrogate bit means the reparse point stands in
    # for another named object, which is exactly what a link is.
    if (([uint64]$Tag -band [uint64]$script:ReparseSurrogateBit) -ne 0) {
        return "Link"
    }

    return "Unknown"
}

function Get-ReparseTagName {
    param(
        [uint32]$Tag
    )

    if ($Tag -eq 0) {
        return ""
    }

    $Hex = '{0:X8}' -f $Tag

    foreach ($Table in @($script:ReparseTransparentTags, $script:ReparseTieredTags, $script:ReparseLinkTags)) {
        if ($Table.ContainsKey($Hex)) {
            return ("{0} (tag 0x{1})" -f $Table[$Hex], $Hex)
        }
    }

    if ($Hex.Substring(0, 4) -eq '9000' -and $Hex.Substring(5) -eq '01A') {
        return ("Cloud files placeholder (tag 0x{0})" -f $Hex)
    }

    return ("Unrecognised reparse point (tag 0x{0})" -f $Hex)
}

function Get-PathReparseInfo {
    <#
        Classifies one item. Anything that cannot be read confidently comes
        back as a link, because a link is the answer that stops the archive.
    #>
    param(
        [string]$Path,
        [System.IO.FileAttributes]$Attributes
    )

    if (($Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne [System.IO.FileAttributes]::ReparsePoint) {
        return [pscustomobject]@{
            Kind    = "None"
            Tag     = [uint32]0
            TagName = ""
            Path    = $Path
        }
    }

    if ($script:ReparseReaderAvailable) {
        try {
            $Tag = [BRC.Native.ReparseReader]::GetReparseTag((ConvertTo-ExtendedLengthPath -Path $Path))

            if ($Tag -ne [uint32]::MaxValue -and $Tag -ne 0) {
                return [pscustomobject]@{
                    Kind    = (Get-ReparseTagKind -Tag $Tag)
                    Tag     = $Tag
                    TagName = (Get-ReparseTagName -Tag $Tag)
                    Path    = $Path
                }
            }
        } catch {
            # Fall through to the conservative answer below.
        }
    }

    return [pscustomobject]@{
        Kind    = "Link"
        Tag     = [uint32]0
        TagName = "Reparse point of an unreadable type"
        Path    = $Path
    }
}

function New-ReparseFindingSet {
    return [pscustomobject]@{
        Links       = (New-Object 'System.Collections.Generic.List[object]')
        Tiered      = (New-Object 'System.Collections.Generic.List[object]')
        StubCount   = 0
        StubExample = ""
    }
}

function Add-ReparseFinding {
    <#
        Records anything that is not an ordinary file. Blocking findings are
        named individually and capped, so a pathological tree cannot exhaust
        memory. Transparent stubs are only counted: on a deduplicated volume
        every file is one, and there is nothing to decide about them.
    #>
    param(
        [pscustomobject]$FindingSet,
        [string]$Path,
        [System.IO.FileAttributes]$Attributes
    )

    if (($Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne [System.IO.FileAttributes]::ReparsePoint) {
        return
    }

    $Info = Get-PathReparseInfo -Path $Path -Attributes $Attributes

    switch ($Info.Kind) {
        "Transparent" {
            $FindingSet.StubCount++

            if ([string]::IsNullOrEmpty($FindingSet.StubExample)) {
                $FindingSet.StubExample = $Info.TagName
            }
        }
        "Link" {
            if ($FindingSet.Links.Count -lt 100) {
                $FindingSet.Links.Add($Info)
            }
        }
        "Tiered" {
            if ($FindingSet.Tiered.Count -lt 100) {
                $FindingSet.Tiered.Add($Info)
            }
        }
        "Unknown" {
            if ($FindingSet.Tiered.Count -lt 100) {
                $FindingSet.Tiered.Add($Info)
            }
        }
    }
}

# ---------------------------------------------------------------------------
# Inventory
# ---------------------------------------------------------------------------

function Get-FolderInventory {
    <#
        Walks one candidate folder and returns everything needed to decide on
        it, to describe it in the placeholder .txt, and to verify the copy:
        totals, the newest timestamp anywhere inside it, a relative-path map of
        every file with its length, the listing lines, and any read errors.
    #>
    param(
        [string]$CurrentPath,
        [string]$RootPath,
        [int]$Depth,
        [System.Collections.Generic.List[string]]$Lines,
        [System.Collections.Generic.Dictionary[string, long]]$FileMap,
        [System.Collections.Generic.List[string]]$Errors,
        [pscustomobject]$ReparseSet
    )

    $Indent = "  " * $Depth
    $DirectFiles = New-Object 'System.Collections.Generic.List[System.IO.FileInfo]'
    $ChildFolders = New-Object 'System.Collections.Generic.List[string]'
    $FolderCount = 0
    $FileCount = 0
    $TotalBytes = [long]0
    $NewestWrite = $null

    try {
        $DirInfo = New-Object System.IO.DirectoryInfo($CurrentPath)
        $NewestWrite = $DirInfo.LastWriteTime

        Add-ReparseFinding -FindingSet $ReparseSet -Path $CurrentPath -Attributes $DirInfo.Attributes
    } catch {
        $Errors.Add("Could not read directory info for '$CurrentPath': $($_.Exception.Message)")
    }

    try {
        foreach ($FilePath in [System.IO.Directory]::EnumerateFiles($CurrentPath)) {
            try {
                $DirectFiles.Add((New-Object System.IO.FileInfo($FilePath)))
            } catch {
                $Errors.Add("Could not read file info for '$FilePath': $($_.Exception.Message)")
            }
        }
    } catch {
        $Errors.Add("Could not enumerate files in '$CurrentPath': $($_.Exception.Message)")
    }

    try {
        foreach ($ChildPath in [System.IO.Directory]::EnumerateDirectories($CurrentPath)) {
            $ChildFolders.Add($ChildPath)
        }
    } catch {
        $Errors.Add("Could not enumerate subfolders in '$CurrentPath': $($_.Exception.Message)")
    }

    $RelativeFolder = Get-RelativePath -RootPath $RootPath -CurrentPath $CurrentPath
    $FolderLabel = if ([string]::IsNullOrWhiteSpace($RelativeFolder)) { [string]$script:Sep } else { [string]$script:Sep + $RelativeFolder }

    $HeaderIndex = $Lines.Count
    $Lines.Add("")

    foreach ($FileInfo in ($DirectFiles | Sort-Object Name)) {
        $FileCount++
        $TotalBytes += [long]$FileInfo.Length

        if ($null -eq $NewestWrite -or $FileInfo.LastWriteTime -gt $NewestWrite) {
            $NewestWrite = $FileInfo.LastWriteTime
        }

        $RelativeFile = Get-RelativePath -RootPath $RootPath -CurrentPath $FileInfo.FullName
        $FileMap[$RelativeFile.ToLowerInvariant()] = [long]$FileInfo.Length

        $Lines.Add(("{0}  [FILE] {1,-58} {2,14} {3}" -f `
            $Indent, `
            $FileInfo.Name, `
            (Format-SizeFriendly -Bytes ([long]$FileInfo.Length)), `
            $FileInfo.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")))

        Add-ReparseFinding -FindingSet $ReparseSet -Path $FileInfo.FullName -Attributes $FileInfo.Attributes
    }

    foreach ($ChildPath in ($ChildFolders | Sort-Object)) {
        $FolderCount++

        $ChildResult = Get-FolderInventory `
            -CurrentPath $ChildPath `
            -RootPath $RootPath `
            -Depth ($Depth + 1) `
            -Lines $Lines `
            -FileMap $FileMap `
            -Errors $Errors `
            -ReparseSet $ReparseSet

        $FolderCount += $ChildResult.FolderCount
        $FileCount += $ChildResult.FileCount
        $TotalBytes += $ChildResult.TotalBytes

        if ($null -ne $ChildResult.NewestWrite) {
            if ($null -eq $NewestWrite -or $ChildResult.NewestWrite -gt $NewestWrite) {
                $NewestWrite = $ChildResult.NewestWrite
            }
        }
    }

    $Lines[$HeaderIndex] = ("{0}[DIR ] {1}   ({2}, {3}, {4})" -f `
        $Indent, `
        $FolderLabel, `
        (Format-Plural -Count $FolderCount -Singular "subfolder" -Plural "subfolders"), `
        (Format-Plural -Count $FileCount -Singular "file" -Plural "files"), `
        (Format-SizeFriendly -Bytes $TotalBytes))

    return [pscustomobject]@{
        FolderCount = $FolderCount
        FileCount   = $FileCount
        TotalBytes  = $TotalBytes
        NewestWrite = $NewestWrite
    }
}

function Get-CandidateInventory {
    param(
        [string]$FolderPath
    )

    $Lines = New-Object 'System.Collections.Generic.List[string]'
    $FileMap = New-Object 'System.Collections.Generic.Dictionary[string, long]'
    $Errors = New-Object 'System.Collections.Generic.List[string]'
    $ReparseSet = New-ReparseFindingSet

    $Result = Get-FolderInventory `
        -CurrentPath $FolderPath `
        -RootPath $FolderPath `
        -Depth 0 `
        -Lines $Lines `
        -FileMap $FileMap `
        -Errors $Errors `
        -ReparseSet $ReparseSet

    $FolderLastWrite = $null

    try {
        $FolderLastWrite = (New-Object System.IO.DirectoryInfo($FolderPath)).LastWriteTime
    } catch {
        $Errors.Add("Could not read directory info for '$FolderPath': $($_.Exception.Message)")
    }

    return [pscustomobject]@{
        FolderCount        = $Result.FolderCount
        FileCount          = $Result.FileCount
        TotalBytes         = $Result.TotalBytes
        NewestWrite        = $Result.NewestWrite
        FolderLastWrite    = $FolderLastWrite
        Lines              = $Lines
        FileMap            = $FileMap
        Errors             = $Errors
        LinkFindings       = $ReparseSet.Links
        TieredFindings     = $ReparseSet.Tiered
        StubCount          = $ReparseSet.StubCount
        StubExample        = $ReparseSet.StubExample
    }
}

# ---------------------------------------------------------------------------
# Copy, verify, delete
# ---------------------------------------------------------------------------

function Copy-TreeWithDotNet {
    param(
        [string]$SourcePath,
        [string]$DestinationPath
    )

    $Errors = New-Object 'System.Collections.Generic.List[string]'
    $Stack = New-Object 'System.Collections.Generic.Stack[object]'
    $CopiedFolders = New-Object 'System.Collections.Generic.List[object]'
    $Stack.Push([pscustomobject]@{ Source = $SourcePath; Destination = $DestinationPath })

    while ($Stack.Count -gt 0) {
        $Pair = $Stack.Pop()
        $CopiedFolders.Add($Pair)

        try {
            if (-not (Test-Path -LiteralPath $Pair.Destination -PathType Container)) {
                [void][System.IO.Directory]::CreateDirectory($Pair.Destination)
            }
        } catch {
            $Errors.Add("Could not create '$($Pair.Destination)': $($_.Exception.Message)")
            continue
        }

        try {
            foreach ($FilePath in [System.IO.Directory]::EnumerateFiles($Pair.Source)) {
                $TargetPath = Join-Path -Path $Pair.Destination -ChildPath ([System.IO.Path]::GetFileName($FilePath))

                try {
                    [System.IO.File]::Copy($FilePath, $TargetPath, $true)
                } catch {
                    $Errors.Add("Could not copy '$FilePath': $($_.Exception.Message)")
                }
            }
        } catch {
            $Errors.Add("Could not enumerate files in '$($Pair.Source)': $($_.Exception.Message)")
        }

        try {
            foreach ($ChildPath in [System.IO.Directory]::EnumerateDirectories($Pair.Source)) {
                $ChildTarget = Join-Path -Path $Pair.Destination -ChildPath ([System.IO.Path]::GetFileName($ChildPath))
                $Stack.Push([pscustomobject]@{ Source = $ChildPath; Destination = $ChildTarget })
            }
        } catch {
            $Errors.Add("Could not enumerate subfolders in '$($Pair.Source)': $($_.Exception.Message)")
        }

    }

    # Folder timestamps are applied only once every file is in place, otherwise
    # creating the children would stamp the parent folder again.
    foreach ($Pair in $CopiedFolders) {
        try {
            $SourceInfo = New-Object System.IO.DirectoryInfo($Pair.Source)
            $TargetInfo = New-Object System.IO.DirectoryInfo($Pair.Destination)
            $TargetInfo.CreationTime = $SourceInfo.CreationTime
            $TargetInfo.LastWriteTime = $SourceInfo.LastWriteTime
        } catch {
            $Errors.Add("Could not carry timestamps to '$($Pair.Destination)': $($_.Exception.Message)")
        }
    }

    return [pscustomobject]@{
        Succeeded = ($Errors.Count -eq 0)
        Errors    = $Errors
    }
}

function Copy-TreeWithRobocopy {
    param(
        [string]$SourcePath,
        [string]$DestinationPath,
        [bool]$CopyPermissions
    )

    $Errors = New-Object 'System.Collections.Generic.List[string]'
    $CopyFlags = if ($CopyPermissions) { "/COPY:DATSO" } else { "/COPY:DAT" }

    $Arguments = @(
        $SourcePath,
        $DestinationPath,
        "/E",
        $CopyFlags,
        "/DCOPY:DAT",
        "/R:2",
        "/W:2",
        "/NFL",
        "/NDL",
        "/NP",
        "/NJH",
        "/NJS"
    )

    try {
        $Output = & robocopy.exe @Arguments 2>&1
        $ExitCode = $LASTEXITCODE
    } catch {
        return [pscustomobject]@{
            Succeeded = $false
            Errors    = @("robocopy could not be started: $($_.Exception.Message)")
        }
    }

    # Robocopy exit codes below 8 are success or informational. 8 and above mean
    # at least one file or folder could not be copied.
    if ($ExitCode -ge 8) {
        $Detail = (($Output | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Last 5) -join " | ")
        $Errors.Add("robocopy returned exit code $ExitCode. $Detail")
    }

    return [pscustomobject]@{
        Succeeded = ($Errors.Count -eq 0)
        Errors    = $Errors
    }
}

function Get-FileHashText {
    param(
        [string]$FilePath
    )

    $Stream = $null
    $Sha = $null

    try {
        $Sha = [System.Security.Cryptography.SHA256]::Create()
        $Stream = [System.IO.File]::OpenRead($FilePath)
        $Hash = $Sha.ComputeHash($Stream)
        return ([System.BitConverter]::ToString($Hash).Replace("-", ""))
    } finally {
        if ($Stream) { $Stream.Dispose() }
        if ($Sha) { $Sha.Dispose() }
    }
}

function Test-ArchiveCopy {
    <#
        Compares the source tree against the archived copy: every relative path
        present on both sides, every file the same length, and optionally every
        file the same SHA-256. Returns the differences rather than throwing, so
        a failed check can be reported per folder and the source left in place.
    #>
    param(
        [string]$SourcePath,
        [string]$DestinationPath,
        [System.Collections.Generic.Dictionary[string, long]]$SourceFileMap,
        [bool]$UseHash,
        [System.Windows.Forms.Label]$StatusLabel
    )

    $Differences = New-Object 'System.Collections.Generic.List[string]'
    $DestinationMap = New-Object 'System.Collections.Generic.Dictionary[string, long]'

    try {
        foreach ($FilePath in [System.IO.Directory]::EnumerateFiles($DestinationPath, "*", [System.IO.SearchOption]::AllDirectories)) {
            try {
                $Relative = (Get-RelativePath -RootPath $DestinationPath -CurrentPath $FilePath).ToLowerInvariant()
                $DestinationMap[$Relative] = [long](New-Object System.IO.FileInfo($FilePath)).Length
            } catch {
                $Differences.Add("Could not read archived file '$FilePath': $($_.Exception.Message)")
            }
        }
    } catch {
        $Differences.Add("Could not enumerate the archived copy: $($_.Exception.Message)")
    }

    foreach ($Key in $SourceFileMap.Keys) {
        if (-not $DestinationMap.ContainsKey($Key)) {
            $Differences.Add("Missing from the archive copy: $Key")
            continue
        }

        if ($DestinationMap[$Key] -ne $SourceFileMap[$Key]) {
            $Differences.Add(("Size mismatch for {0}: source {1} bytes, archive {2} bytes" -f $Key, $SourceFileMap[$Key], $DestinationMap[$Key]))
        }
    }

    foreach ($Key in $DestinationMap.Keys) {
        if (-not $SourceFileMap.ContainsKey($Key)) {
            $Differences.Add("Unexpected extra file in the archive copy: $Key")
        }
    }

    if ($UseHash -and $Differences.Count -eq 0) {
        $Checked = 0

        foreach ($Key in $SourceFileMap.Keys) {
            $Checked++

            if (($Checked % 25) -eq 0 -and $null -ne $StatusLabel) {
                $StatusLabel.Text = "Verifying hashes: $Checked of $($SourceFileMap.Count) files..."
                [System.Windows.Forms.Application]::DoEvents()
            }

            try {
                $SourceFile = Join-Path -Path $SourcePath -ChildPath $Key
                $DestinationFile = Join-Path -Path $DestinationPath -ChildPath $Key

                if ((Get-FileHashText -FilePath $SourceFile) -ne (Get-FileHashText -FilePath $DestinationFile)) {
                    $Differences.Add("SHA-256 mismatch for $Key")
                }
            } catch {
                $Differences.Add("Could not hash '$Key': $($_.Exception.Message)")
            }
        }
    }

    return [pscustomobject]@{
        Matches     = ($Differences.Count -eq 0)
        Differences = $Differences
        FileCount   = $DestinationMap.Count
    }
}

function Remove-SourceTree {
    param(
        [string]$FolderPath,
        [bool]$ClearReadOnly,
        [bool]$UseRecycleBin
    )

    $Errors = New-Object 'System.Collections.Generic.List[string]'

    if ($ClearReadOnly) {
        try {
            foreach ($ItemPath in [System.IO.Directory]::EnumerateFiles($FolderPath, "*", [System.IO.SearchOption]::AllDirectories)) {
                try {
                    $FileInfo = New-Object System.IO.FileInfo($ItemPath)

                    if ($FileInfo.IsReadOnly) {
                        $FileInfo.IsReadOnly = $false
                    }
                } catch {
                    $Errors.Add("Could not clear the read-only flag on '$ItemPath': $($_.Exception.Message)")
                }
            }
        } catch {
            $Errors.Add("Could not enumerate '$FolderPath' to clear read-only flags: $($_.Exception.Message)")
        }
    }

    try {
        if ($UseRecycleBin) {
            Add-Type -AssemblyName Microsoft.VisualBasic
            [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory(
                $FolderPath,
                [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs,
                [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin)
        } else {
            [System.IO.Directory]::Delete($FolderPath, $true)
        }
    } catch {
        $Errors.Add("Could not remove the source folder: $($_.Exception.Message)")
    }

    return [pscustomobject]@{
        Succeeded = ($Errors.Count -eq 0)
        Errors    = $Errors
    }
}

# ---------------------------------------------------------------------------
# Placeholder and reference files
# ---------------------------------------------------------------------------

function Write-TextFileUtf8 {
    param(
        [string]$Path,
        [string[]]$Lines
    )

    $Writer = New-Object System.IO.StreamWriter($Path, $false, (New-Object System.Text.UTF8Encoding($false)))

    try {
        foreach ($Line in $Lines) {
            $Writer.WriteLine($Line)
        }

        $Writer.Flush()
    } finally {
        $Writer.Close()
        $Writer.Dispose()
    }
}

function New-PlaceholderContent {
    param(
        [pscustomobject]$Config,
        [string]$OriginalPath,
        [string]$ArchivePath,
        [string]$ReferenceCode,
        [pscustomobject]$Inventory,
        [datetime]$Timestamp
    )

    $AgeRuleText = "{0} {1}" -f $Config.ThresholdValue, $Config.ThresholdUnit
    $Divider = "=" * 78
    $SubDivider = "-" * 78

    $Lines = New-Object 'System.Collections.Generic.List[string]'

    [void]$Lines.Add($Divider)
    [void]$Lines.Add("  THIS CONTENT HAS BEEN ARCHIVED")
    [void]$Lines.Add($Divider)
    [void]$Lines.Add("")
    [void]$Lines.Add("The folder that used to be in this location has been moved to long-term")
    [void]$Lines.Add("archive storage because nothing inside it had been modified for at least")
    [void]$Lines.Add("${AgeRuleText}.")
    [void]$Lines.Add("")
    [void]$Lines.Add("Nothing has been deleted. Every folder and file listed further down this")
    [void]$Lines.Add("file still exists in the archive and can be restored on request.")
    [void]$Lines.Add("")
    [void]$Lines.Add("If you need anything from this content, please contact I.T. and quote the")
    [void]$Lines.Add("archive reference code below. Please do not delete this file - it is the")
    [void]$Lines.Add("only pointer back to where the data went.")
    [void]$Lines.Add("")
    [void]$Lines.Add("  ARCHIVE REFERENCE CODE:  $ReferenceCode")
    [void]$Lines.Add("")

    if (-not [string]::IsNullOrWhiteSpace($Config.ContactText)) {
        foreach ($ContactLine in ($Config.ContactText -split "`r?`n")) {
            [void]$Lines.Add("  Contact:  $ContactLine")
        }

        [void]$Lines.Add("")
    }

    [void]$Lines.Add($SubDivider)
    [void]$Lines.Add("SUMMARY")
    [void]$Lines.Add($SubDivider)
    [void]$Lines.Add(("Original location  : {0}" -f $OriginalPath))

    if ($Config.ShowArchivePathInPlaceholder) {
        [void]$Lines.Add(("Archived to        : {0}" -f $ArchivePath))
    } else {
        [void]$Lines.Add("Archived to        : (held by I.T. - quote the reference code above)")
    }

    [void]$Lines.Add(("Archived on        : {0}" -f $Timestamp.ToString("yyyy-MM-dd HH:mm:ss")))
    [void]$Lines.Add(("Archived by        : {0}\{1} on {2}" -f $env:USERDOMAIN, $env:USERNAME, $env:COMPUTERNAME))
    [void]$Lines.Add(("Age rule applied   : not modified in the last {0} (cut-off {1})" -f $AgeRuleText, $Config.CutoffDate.ToString("yyyy-MM-dd")))
    [void]$Lines.Add(("Age measured by    : {0}" -f $Config.AgeBasis))
    [void]$Lines.Add(("Last modified      : {0}" -f (Format-DateForCsv -DateValue $Inventory.NewestWrite)))
    [void]$Lines.Add(("Folders archived   : {0}" -f ($Inventory.FolderCount + 1)))
    [void]$Lines.Add(("Files archived     : {0}" -f $Inventory.FileCount))
    [void]$Lines.Add(("Total size         : {0} ({1:N0} bytes)" -f (Format-SizeFriendly -Bytes $Inventory.TotalBytes), $Inventory.TotalBytes))
    [void]$Lines.Add("")
    [void]$Lines.Add($SubDivider)
    [void]$Lines.Add("EVERY FOLDER, SUBFOLDER AND FILE THAT WAS ARCHIVED")
    [void]$Lines.Add($SubDivider)
    [void]$Lines.Add("Paths are shown relative to the archived folder itself.")
    [void]$Lines.Add(("Top level: {0}" -f $OriginalPath))

    foreach ($Line in $Inventory.Lines) {
        [void]$Lines.Add($Line)
    }

    [void]$Lines.Add("")
    [void]$Lines.Add($SubDivider)
    [void]$Lines.Add(("End of listing. Archive reference code: {0}" -f $ReferenceCode))
    [void]$Lines.Add(("Generated by BRC Folder Archiver on {0}." -f $Timestamp.ToString("yyyy-MM-dd HH:mm:ss")))
    [void]$Lines.Add($SubDivider)

    return $Lines.ToArray()
}

function Get-PlaceholderFileName {
    param(
        [string]$Template,
        [string]$FolderName,
        [string]$ReferenceCode,
        [datetime]$Timestamp
    )

    $Name = $Template

    if ([string]::IsNullOrWhiteSpace($Name)) {
        $Name = "{FolderName}_ARCHIVED_{Ref}.txt"
    }

    $Name = $Name.Replace("{FolderName}", $FolderName)
    $Name = $Name.Replace("{Ref}", $ReferenceCode)
    $Name = $Name.Replace("{Date}", $Timestamp.ToString("yyyyMMdd"))
    $Name = Get-SafeFileNameFragment -Text $Name

    if (-not $Name.ToLowerInvariant().EndsWith(".txt")) {
        $Name = $Name + ".txt"
    }

    return $Name
}

# ---------------------------------------------------------------------------
# Candidate selection
# ---------------------------------------------------------------------------

function Build-NewestWriteMap {
    <#
        One bottom-up pass over the source tree that records, for every folder,
        the newest LastWriteTime found anywhere beneath it. Doing this once up
        front means "top-most stale folder" mode does not rescan the same
        subtree at every level.
    #>
    param(
        [string]$CurrentPath,
        [System.Collections.Generic.Dictionary[string, object]]$Map,
        [System.Windows.Forms.Label]$StatusLabel,
        [ref]$ScannedCount
    )

    $ScannedCount.Value++

    if (($ScannedCount.Value % 50) -eq 0) {
        $StatusLabel.Text = "Assessing folder ages: $($ScannedCount.Value) folders examined..."
        [System.Windows.Forms.Application]::DoEvents()
    }

    $HadError = $false
    $NewestWrite = $null
    $IsLink = $false

    try {
        $DirInfo = New-Object System.IO.DirectoryInfo($CurrentPath)
        $NewestWrite = $DirInfo.LastWriteTime

        # Only a link is a reason not to walk in. A folder that is merely
        # deduplicated or compressed is an ordinary folder.
        if ((Get-PathReparseInfo -Path $CurrentPath -Attributes $DirInfo.Attributes).Kind -eq "Link") {
            $IsLink = $true
        }
    } catch {
        $HadError = $true
    }

    if (-not $IsLink) {
        try {
            foreach ($FilePath in [System.IO.Directory]::EnumerateFiles($CurrentPath)) {
                try {
                    $FileInfo = New-Object System.IO.FileInfo($FilePath)

                    if ($null -eq $NewestWrite -or $FileInfo.LastWriteTime -gt $NewestWrite) {
                        $NewestWrite = $FileInfo.LastWriteTime
                    }
                } catch {
                    $HadError = $true
                }
            }
        } catch {
            $HadError = $true
        }

        try {
            foreach ($ChildPath in [System.IO.Directory]::EnumerateDirectories($CurrentPath)) {
                if ($script:CancelRequested) {
                    break
                }

                $ChildResult = Build-NewestWriteMap `
                    -CurrentPath $ChildPath `
                    -Map $Map `
                    -StatusLabel $StatusLabel `
                    -ScannedCount $ScannedCount

                if ($ChildResult.HadError) {
                    $HadError = $true
                }

                if ($null -ne $ChildResult.NewestWrite) {
                    if ($null -eq $NewestWrite -or $ChildResult.NewestWrite -gt $NewestWrite) {
                        $NewestWrite = $ChildResult.NewestWrite
                    }
                }
            }
        } catch {
            $HadError = $true
        }
    }

    $Entry = [pscustomobject]@{
        NewestWrite = $NewestWrite
        HadError    = $HadError
        IsLink      = $IsLink
    }

    $Map[(Get-NormalisedPath -Path $CurrentPath).ToLowerInvariant()] = $Entry

    return $Entry
}

function Get-FolderAgeEntry {
    param(
        [pscustomobject]$Config,
        [string]$FolderPath
    )

    if ($Config.AgeBasis -eq "Recursive newest item timestamp") {
        $Key = (Get-NormalisedPath -Path $FolderPath).ToLowerInvariant()

        if ($Config.NewestWriteMap.ContainsKey($Key)) {
            return $Config.NewestWriteMap[$Key]
        }
    }

    $HadError = $false
    $NewestWrite = $null
    $IsLink = $false

    try {
        $DirInfo = New-Object System.IO.DirectoryInfo($FolderPath)
        $NewestWrite = $DirInfo.LastWriteTime

        if ((Get-PathReparseInfo -Path $FolderPath -Attributes $DirInfo.Attributes).Kind -eq "Link") {
            $IsLink = $true
        }
    } catch {
        $HadError = $true
    }

    return [pscustomobject]@{
        NewestWrite = $NewestWrite
        HadError    = $HadError
        IsLink      = $IsLink
    }
}

function Test-FolderIsExcluded {
    param(
        [pscustomobject]$Config,
        [string]$FolderName
    )

    foreach ($Pattern in $Config.ExcludePatterns) {
        if ([string]::IsNullOrWhiteSpace($Pattern)) {
            continue
        }

        if ($FolderName -like $Pattern) {
            return $true
        }
    }

    return $false
}

# ---------------------------------------------------------------------------
# Archiving one folder
# ---------------------------------------------------------------------------

$script:ArchiveCsvColumns = @(
    "Status",
    "ReferenceCode",
    "OriginalPath",
    "ArchivePath",
    "PlaceholderFile",
    "ArchiveCodeFile",
    "RelativePath",
    "FolderName",
    "OriginalParentPath",
    "DepthFromSourceRoot",
    "LastModifiedUsed",
    "DaysSinceModified",
    "CutoffDate",
    "ThresholdValue",
    "ThresholdUnit",
    "AgeBasis",
    "FolderCount",
    "FileCount",
    "TotalSizeBytes",
    "TotalSizeMB",
    "TotalSizeGB",
    "SizeFriendly",
    "CopyMethod",
    "VerifyResult",
    "SourceRemoved",
    "ArchivedLocal",
    "ArchivedUtc",
    "ArchivedBy",
    "RunFromHost",
    "SourceRoot",
    "ArchiveRoot",
    "ErrorCount",
    "Errors"
)

function ConvertTo-ResultRow {
    param(
        [pscustomobject]$Config,
        [pscustomobject]$Result
    )

    return @(
        $Result.Status,
        $Result.ReferenceCode,
        $Result.OriginalPath,
        $Result.ArchivePath,
        $Result.PlaceholderFile,
        $Result.ArchiveCodeFile,
        $Result.RelativePath,
        $Result.FolderName,
        $Result.OriginalParentPath,
        [string]$Result.Depth,
        (Format-DateForCsv -DateValue $Result.LastModifiedUsed),
        [string](Get-DaysSince -DateValue $Result.LastModifiedUsed),
        (Format-DateForCsv -DateValue $Config.CutoffDate),
        [string]$Config.ThresholdValue,
        $Config.ThresholdUnit,
        $Config.AgeBasis,
        [string]$Result.FolderCount,
        [string]$Result.FileCount,
        [string]$Result.TotalBytes,
        [string]([math]::Round($Result.TotalBytes / 1MB, 2)),
        [string]([math]::Round($Result.TotalBytes / 1GB, 3)),
        (Format-SizeFriendly -Bytes $Result.TotalBytes),
        $Result.CopyMethod,
        $Result.VerifyResult,
        [string]$Result.SourceRemoved,
        (Format-DateForCsv -DateValue $Result.Timestamp),
        $Result.Timestamp.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss"),
        ("{0}\{1}" -f $env:USERDOMAIN, $env:USERNAME),
        $env:COMPUTERNAME,
        $Config.SourceRoot,
        $Config.ArchiveRoot,
        [string]$Result.Errors.Count,
        (($Result.Errors | Sort-Object -Unique) -join "; ")
    )
}

function Format-ReparseSkipReason {
    param(
        [System.Collections.Generic.List[object]]$Findings,
        [string]$What,
        [string]$Why
    )

    $Example = $Findings[0]
    $Extra = ""

    if ($Findings.Count -gt 1) {
        $Extra = " and {0} other such {1}" -f ($Findings.Count - 1), $(if ($Findings.Count -eq 2) { "item" } else { "items" })
    }

    return ("Skipped because this folder contains {0}: '{1}' is {2}{3}. Nothing was copied or deleted, because {4}." -f `
        $What, $Example.Path, $Example.TagName, $Extra, $Why)
}

function Invoke-ArchiveCandidate {
    param(
        [pscustomobject]$Config,
        [string]$FolderPath,
        [int]$Depth,
        [System.Windows.Forms.Label]$StatusLabel
    )

    $Timestamp = Get-Date
    $FolderName = Split-Path -Path $FolderPath -Leaf
    $RelativePath = Get-RelativePath -RootPath $Config.SourceRoot -CurrentPath $FolderPath
    $ParentPath = Split-Path -Path $FolderPath -Parent
    $Errors = New-Object 'System.Collections.Generic.List[string]'

    $Result = [pscustomobject]@{
        Status             = "Skipped"
        ReferenceCode      = ""
        OriginalPath       = $FolderPath
        ArchivePath        = ""
        PlaceholderFile    = ""
        ArchiveCodeFile    = ""
        RelativePath       = $RelativePath
        FolderName         = $FolderName
        OriginalParentPath = $ParentPath
        Depth              = $Depth
        LastModifiedUsed   = $null
        FolderCount        = 0
        FileCount          = 0
        TotalBytes         = [long]0
        CopyMethod         = ""
        VerifyResult       = ""
        SourceRemoved      = $false
        Timestamp          = $Timestamp
        Errors             = $Errors
    }

    $StatusLabel.Text = "Reading contents of $FolderPath ..."
    [System.Windows.Forms.Application]::DoEvents()

    $Inventory = Get-CandidateInventory -FolderPath $FolderPath

    $Result.FolderCount = $Inventory.FolderCount + 1
    $Result.FileCount = $Inventory.FileCount
    $Result.TotalBytes = $Inventory.TotalBytes
    $Result.LastModifiedUsed = if ($Config.AgeBasis -eq "Folder timestamp only") { $Inventory.FolderLastWrite } else { $Inventory.NewestWrite }

    if ($Inventory.Errors.Count -gt 0) {
        foreach ($Message in $Inventory.Errors) {
            $Errors.Add($Message)
        }

        $Result.Status = "Skipped"
        $Errors.Add("Skipped because the folder could not be read completely. Nothing was copied or deleted.")
        return $Result
    }

    if ($Config.SkipLinkedFolders -and $Inventory.LinkFindings.Count -gt 0) {
        $Result.Status = "Skipped"
        $Errors.Add((Format-ReparseSkipReason -Findings $Inventory.LinkFindings -What "a junction, symbolic link or mount point" -Why "it points somewhere else, so copying it would pull in data from outside this folder and removing it could reach through to the target"))
        return $Result
    }

    if ($Config.SkipTieredFiles -and $Inventory.TieredFindings.Count -gt 0) {
        $Result.Status = "Skipped"
        $Errors.Add((Format-ReparseSkipReason -Findings $Inventory.TieredFindings -What "cloud-tiered, HSM or unrecognised stub files" -Why "the data lives elsewhere and copying it would recall every file, which can be slow and can cost money"))
        return $Result
    }

    if ($Config.SkipEmptyFolders -and $Inventory.FileCount -eq 0) {
        $Result.Status = "Skipped"
        $Errors.Add("Skipped because the folder contains no files.")
        return $Result
    }

    $ReferenceCode = New-ArchiveReferenceCode
    $Result.ReferenceCode = $ReferenceCode

    if ($Config.PreserveStructure -and -not [string]::IsNullOrWhiteSpace($RelativePath)) {
        $DesiredArchivePath = Join-Path -Path $Config.EffectiveArchiveRoot -ChildPath $RelativePath
    } else {
        $DesiredArchivePath = Join-Path -Path $Config.EffectiveArchiveRoot -ChildPath $FolderName
    }

    if ($Config.PreviewOnly) {
        $Result.Status = "WouldArchive"
        $Result.ArchivePath = $DesiredArchivePath
        $Result.CopyMethod = $Config.CopyMethod
        $Result.VerifyResult = "Not run (preview)"
        $Result.PlaceholderFile = Join-Path -Path $ParentPath -ChildPath (Get-PlaceholderFileName -Template $Config.PlaceholderTemplate -FolderName $FolderName -ReferenceCode $ReferenceCode -Timestamp $Timestamp)
        $Result.ArchiveCodeFile = Join-Path -Path (Split-Path -Path $DesiredArchivePath -Parent) -ChildPath ("{0}.txt" -f $ReferenceCode)
        return $Result
    }

    try {
        $ArchiveParent = Split-Path -Path $DesiredArchivePath -Parent

        if (-not (Test-Path -LiteralPath $ArchiveParent -PathType Container)) {
            [void][System.IO.Directory]::CreateDirectory($ArchiveParent)
        }

        $ArchivePath = Get-NonClashingPath -DesiredPath $DesiredArchivePath
    } catch {
        $Result.Status = "Failed"
        $Errors.Add("Could not prepare the archive location: $($_.Exception.Message)")
        return $Result
    }

    $Result.ArchivePath = $ArchivePath
    $Result.CopyMethod = $Config.CopyMethod

    $StatusLabel.Text = ("Copying {0} ({1} files, {2}) to the archive..." -f $FolderName, $Inventory.FileCount, (Format-SizeFriendly -Bytes $Inventory.TotalBytes))
    [System.Windows.Forms.Application]::DoEvents()

    if ($Config.CopyMethod -eq "Robocopy") {
        $CopyResult = Copy-TreeWithRobocopy -SourcePath $FolderPath -DestinationPath $ArchivePath -CopyPermissions $Config.CopyPermissions
    } else {
        $CopyResult = Copy-TreeWithDotNet -SourcePath $FolderPath -DestinationPath $ArchivePath
    }

    if (-not $CopyResult.Succeeded) {
        foreach ($Message in $CopyResult.Errors) {
            $Errors.Add($Message)
        }

        $Result.Status = "Failed"
        $Result.VerifyResult = "Not run (copy failed)"
        Remove-FailedArchiveCopy -Config $Config -ArchivePath $ArchivePath -Errors $Errors
        return $Result
    }

    $StatusLabel.Text = ("Verifying the archived copy of {0} ..." -f $FolderName)
    [System.Windows.Forms.Application]::DoEvents()

    $Verification = Test-ArchiveCopy `
        -SourcePath $FolderPath `
        -DestinationPath $ArchivePath `
        -SourceFileMap $Inventory.FileMap `
        -UseHash $Config.VerifyWithHash `
        -StatusLabel $StatusLabel

    if (-not $Verification.Matches) {
        $Result.Status = "Failed"
        $Result.VerifyResult = "Failed"

        foreach ($Message in ($Verification.Differences | Select-Object -First 20)) {
            $Errors.Add($Message)
        }

        if ($Verification.Differences.Count -gt 20) {
            $Errors.Add(("...and {0} further differences." -f ($Verification.Differences.Count - 20)))
        }

        $Errors.Add("The source folder was left untouched because the archive copy did not verify.")
        Remove-FailedArchiveCopy -Config $Config -ArchivePath $ArchivePath -Errors $Errors
        return $Result
    }

    $Result.VerifyResult = if ($Config.VerifyWithHash) { "Passed (names, sizes, SHA-256)" } else { "Passed (names and sizes)" }

    # The archive-side reference file: file name and file contents are both the
    # 24 character code, so a single search for the code finds the archive.
    try {
        $CodeFilePath = Join-Path -Path (Split-Path -Path $ArchivePath -Parent) -ChildPath ("{0}.txt" -f $ReferenceCode)
        Write-TextFileUtf8 -Path $CodeFilePath -Lines @($ReferenceCode)
        $Result.ArchiveCodeFile = $CodeFilePath
    } catch {
        $Errors.Add("Could not write the archive-side reference file: $($_.Exception.Message)")
    }

    $PlaceholderLines = New-PlaceholderContent `
        -Config $Config `
        -OriginalPath $FolderPath `
        -ArchivePath $ArchivePath `
        -ReferenceCode $ReferenceCode `
        -Inventory $Inventory `
        -Timestamp $Timestamp

    if ($Config.WriteArchiveSideManifest) {
        try {
            $ManifestPath = Join-Path -Path (Split-Path -Path $ArchivePath -Parent) -ChildPath ("{0}_manifest.txt" -f $ReferenceCode)
            Write-TextFileUtf8 -Path $ManifestPath -Lines $PlaceholderLines
        } catch {
            $Errors.Add("Could not write the archive-side manifest: $($_.Exception.Message)")
        }
    }

    $StatusLabel.Text = ("Removing the original copy of {0} ..." -f $FolderName)
    [System.Windows.Forms.Application]::DoEvents()

    $Removal = Remove-SourceTree -FolderPath $FolderPath -ClearReadOnly $Config.ClearReadOnly -UseRecycleBin $Config.UseRecycleBin

    if (-not $Removal.Succeeded) {
        foreach ($Message in $Removal.Errors) {
            $Errors.Add($Message)
        }

        $Result.Status = "Failed"
        $Errors.Add("The archive copy verified but the original could not be removed. Both copies now exist; remove the original by hand once the cause is cleared.")
        return $Result
    }

    $Result.SourceRemoved = $true

    try {
        $PlaceholderName = Get-PlaceholderFileName -Template $Config.PlaceholderTemplate -FolderName $FolderName -ReferenceCode $ReferenceCode -Timestamp $Timestamp
        $PlaceholderPath = Get-NonClashingPath -DesiredPath (Join-Path -Path $ParentPath -ChildPath $PlaceholderName)
        Write-TextFileUtf8 -Path $PlaceholderPath -Lines $PlaceholderLines
        $Result.PlaceholderFile = $PlaceholderPath
    } catch {
        $Errors.Add("Could not write the placeholder .txt file: $($_.Exception.Message)")
    }

    $Result.Status = if ($Errors.Count -eq 0) { "Archived" } else { "ArchivedWithWarnings" }

    return $Result
}

function Remove-FailedArchiveCopy {
    param(
        [pscustomobject]$Config,
        [string]$ArchivePath,
        [System.Collections.Generic.List[string]]$Errors
    )

    if (-not $Config.CleanUpFailedCopies) {
        $Errors.Add("The partial archive copy was left at '$ArchivePath' for inspection.")
        return
    }

    try {
        if (Test-Path -LiteralPath $ArchivePath -PathType Container) {
            [System.IO.Directory]::Delete($ArchivePath, $true)
            $Errors.Add("The partial archive copy at '$ArchivePath' was removed. The source is untouched.")
        }
    } catch {
        $Errors.Add("Could not remove the partial archive copy at '$ArchivePath': $($_.Exception.Message)")
    }
}

# ---------------------------------------------------------------------------
# Walking the source tree
# ---------------------------------------------------------------------------

function Invoke-ArchiveWalk {
    param(
        [pscustomobject]$Config,
        [string]$CurrentPath,
        [int]$Depth,
        [pscustomobject]$Writers,
        [hashtable]$Counters,
        [System.Windows.Forms.Label]$StatusLabel
    )

    if ($script:CancelRequested -or $Counters.LimitReached) {
        return
    }

    $ChildFolders = New-Object 'System.Collections.Generic.List[string]'

    try {
        foreach ($ChildPath in [System.IO.Directory]::EnumerateDirectories($CurrentPath)) {
            $ChildFolders.Add($ChildPath)
        }
    } catch {
        $Counters.EnumerationErrors++
        return
    }

    foreach ($ChildPath in ($ChildFolders | Sort-Object)) {
        if ($script:CancelRequested -or $Counters.LimitReached) {
            return
        }

        [System.Windows.Forms.Application]::DoEvents()

        $ChildDepth = $Depth + 1
        $ChildName = Split-Path -Path $ChildPath -Leaf
        $Counters.FoldersExamined++

        if (Test-FolderIsExcluded -Config $Config -FolderName $ChildName) {
            $Counters.Excluded++
            continue
        }

        $IsCandidate = $false
        $DescendWhenNotStale = $false

        # Written as if/elseif rather than switch on purpose: inside a switch,
        # continue applies to the switch and not to this foreach.
        if ($Config.SelectionMode -eq "Top-most stale folder") {
            $IsCandidate = $true
            $DescendWhenNotStale = $true
        } elseif ($Config.SelectionMode -eq "Immediate subfolders of the source root only") {
            $IsCandidate = ($ChildDepth -eq 1)
        } elseif ($Config.SelectionMode -eq "Fixed depth below the source root") {
            if ($ChildDepth -eq $Config.FixedDepth) {
                $IsCandidate = $true
            } elseif ($ChildDepth -lt $Config.FixedDepth) {
                Invoke-ArchiveWalk -Config $Config -CurrentPath $ChildPath -Depth $ChildDepth -Writers $Writers -Counters $Counters -StatusLabel $StatusLabel
                continue
            }
        }

        if (-not $IsCandidate) {
            continue
        }

        $AgeEntry = Get-FolderAgeEntry -Config $Config -FolderPath $ChildPath

        if ($AgeEntry.IsLink -and $Config.SkipLinkedFolders) {
            $Counters.Skipped++
            continue
        }

        if ($AgeEntry.HadError) {
            $Counters.Skipped++

            if ($Config.LogSkipped) {
                $SkipResult = [pscustomobject]@{
                    Status             = "Skipped"
                    ReferenceCode      = ""
                    OriginalPath       = $ChildPath
                    ArchivePath        = ""
                    PlaceholderFile    = ""
                    ArchiveCodeFile    = ""
                    RelativePath       = (Get-RelativePath -RootPath $Config.SourceRoot -CurrentPath $ChildPath)
                    FolderName         = $ChildName
                    OriginalParentPath = (Split-Path -Path $ChildPath -Parent)
                    Depth              = $ChildDepth
                    LastModifiedUsed   = $AgeEntry.NewestWrite
                    FolderCount        = 0
                    FileCount          = 0
                    TotalBytes         = [long]0
                    CopyMethod         = ""
                    VerifyResult       = ""
                    SourceRemoved      = $false
                    Timestamp          = (Get-Date)
                    Errors             = (New-Object 'System.Collections.Generic.List[string]')
                }

                $SkipResult.Errors.Add("Skipped because part of this folder could not be read. Nothing was copied or deleted.")
                Write-ArchiveResult -Config $Config -Result $SkipResult -Writers $Writers
            }

            continue
        }

        $IsStale = ($null -ne $AgeEntry.NewestWrite -and $AgeEntry.NewestWrite -lt $Config.CutoffDate)

        if (-not $IsStale) {
            if ($DescendWhenNotStale) {
                Invoke-ArchiveWalk -Config $Config -CurrentPath $ChildPath -Depth $ChildDepth -Writers $Writers -Counters $Counters -StatusLabel $StatusLabel
            }

            continue
        }

        $Result = Invoke-ArchiveCandidate -Config $Config -FolderPath $ChildPath -Depth $ChildDepth -StatusLabel $StatusLabel

        switch ($Result.Status) {
            "Archived" {
                $Counters.Archived++
                $Counters.BytesArchived += $Result.TotalBytes
                $Counters.FilesArchived += $Result.FileCount
            }
            "ArchivedWithWarnings" {
                $Counters.Archived++
                $Counters.Warnings++
                $Counters.BytesArchived += $Result.TotalBytes
                $Counters.FilesArchived += $Result.FileCount
            }
            "WouldArchive" {
                $Counters.WouldArchive++
                $Counters.BytesArchived += $Result.TotalBytes
                $Counters.FilesArchived += $Result.FileCount
            }
            "Failed" {
                $Counters.Failed++
            }
            default {
                $Counters.Skipped++
            }
        }

        if ($Result.Status -ne "Skipped" -or $Config.LogSkipped) {
            Write-ArchiveResult -Config $Config -Result $Result -Writers $Writers
        }

        $StatusLabel.Text = ("Examined {0} folders. Archived {1}. Skipped {2}. Failed {3}. Total {4}." -f `
            $Counters.FoldersExamined, `
            ($Counters.Archived + $Counters.WouldArchive), `
            $Counters.Skipped, `
            $Counters.Failed, `
            (Format-SizeFriendly -Bytes $Counters.BytesArchived))
        [System.Windows.Forms.Application]::DoEvents()

        if ($Config.MaxFolders -gt 0 -and (($Counters.Archived + $Counters.WouldArchive) -ge $Config.MaxFolders)) {
            $Counters.LimitReached = $true
            return
        }

        # A folder that was skipped or failed is still on disk, so in top-most
        # mode its children remain worth considering.
        if ($DescendWhenNotStale -and ($Result.Status -eq "Skipped" -or $Result.Status -eq "Failed")) {
            if (Test-Path -LiteralPath $ChildPath -PathType Container) {
                Invoke-ArchiveWalk -Config $Config -CurrentPath $ChildPath -Depth $ChildDepth -Writers $Writers -Counters $Counters -StatusLabel $StatusLabel
            }
        }
    }
}

function Write-ArchiveResult {
    param(
        [pscustomobject]$Config,
        [pscustomobject]$Result,
        [pscustomobject]$Writers
    )

    $Row = ConvertTo-ResultRow -Config $Config -Result $Result

    Write-CsvRow -Writer $Writers.RunCsv -Values $Row
    $Writers.RunCsv.Flush()

    if ($null -ne $Writers.IndexCsv) {
        Write-CsvRow -Writer $Writers.IndexCsv -Values $Row
        $Writers.IndexCsv.Flush()
    }
}

# ---------------------------------------------------------------------------
# Confirmation and permission checks
# ---------------------------------------------------------------------------

function Show-TypedConfirmation {
    <#
        A live run moves and then deletes real data, so it is gated behind a
        typed word rather than a single mis-clickable Yes button.
    #>
    param(
        [string]$Message,
        [string]$RequiredWord
    )

    $Dialog = New-Object System.Windows.Forms.Form
    $Dialog.Text = "Confirm a live archive run"
    $Dialog.Size = New-Object System.Drawing.Size(640, 400)
    $Dialog.StartPosition = "CenterParent"
    $Dialog.FormBorderStyle = "FixedDialog"
    $Dialog.MaximizeBox = $false
    $Dialog.MinimizeBox = $false

    $MessageBox = New-Object System.Windows.Forms.TextBox
    $MessageBox.Location = New-Object System.Drawing.Point(15, 15)
    $MessageBox.Size = New-Object System.Drawing.Size(595, 250)
    $MessageBox.Multiline = $true
    $MessageBox.ReadOnly = $true
    $MessageBox.ScrollBars = "Vertical"
    $MessageBox.Text = $Message
    $Dialog.Controls.Add($MessageBox)

    $PromptLabel = New-Object System.Windows.Forms.Label
    $PromptLabel.Text = "Type $RequiredWord to continue:"
    $PromptLabel.Location = New-Object System.Drawing.Point(15, 280)
    $PromptLabel.Size = New-Object System.Drawing.Size(200, 22)
    $Dialog.Controls.Add($PromptLabel)

    $InputBox = New-Object System.Windows.Forms.TextBox
    $InputBox.Location = New-Object System.Drawing.Point(220, 277)
    $InputBox.Size = New-Object System.Drawing.Size(180, 24)
    $Dialog.Controls.Add($InputBox)

    $OkButton = New-Object System.Windows.Forms.Button
    $OkButton.Text = "Run it"
    $OkButton.Location = New-Object System.Drawing.Point(410, 320)
    $OkButton.Size = New-Object System.Drawing.Size(95, 30)
    $OkButton.Enabled = $false
    $OkButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $Dialog.Controls.Add($OkButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Text = "Cancel"
    $CancelButton.Location = New-Object System.Drawing.Point(515, 320)
    $CancelButton.Size = New-Object System.Drawing.Size(95, 30)
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $Dialog.Controls.Add($CancelButton)

    $InputBox.Add_TextChanged({
        $OkButton.Enabled = ($InputBox.Text.Trim() -ceq $RequiredWord)
    }.GetNewClosure())

    $Dialog.AcceptButton = $OkButton
    $Dialog.CancelButton = $CancelButton

    $Outcome = $Dialog.ShowDialog()
    $Dialog.Dispose()

    return ($Outcome -eq [System.Windows.Forms.DialogResult]::OK)
}

function Get-OpenArchivePermissions {
    <#
        Archived data is usually the only copy left, so a wide-open archive
        share is a real risk: anyone who can reach it can delete or alter the
        one remaining copy. This reports broad groups that hold write access.
    #>
    param(
        [string]$Path
    )

    $Findings = New-Object 'System.Collections.Generic.List[string]'

    try {
        $Acl = Get-Acl -LiteralPath $Path
    } catch {
        return $Findings
    }

    $BroadIdentities = @("Everyone", "Authenticated Users", "Domain Users", "BUILTIN\Users", "Users")

    # Only the bits that actually let someone change or destroy data. Modify and
    # FullControl both contain these, so they still match, but a plain read or
    # read-and-execute grant does not raise a false alarm.
    $WriteRights = [int][System.Security.AccessControl.FileSystemRights]::WriteData -bor
                   [int][System.Security.AccessControl.FileSystemRights]::AppendData -bor
                   [int][System.Security.AccessControl.FileSystemRights]::WriteExtendedAttributes -bor
                   [int][System.Security.AccessControl.FileSystemRights]::WriteAttributes -bor
                   [int][System.Security.AccessControl.FileSystemRights]::Delete -bor
                   [int][System.Security.AccessControl.FileSystemRights]::DeleteSubdirectoriesAndFiles -bor
                   [int][System.Security.AccessControl.FileSystemRights]::ChangePermissions -bor
                   [int][System.Security.AccessControl.FileSystemRights]::TakeOwnership

    foreach ($Rule in $Acl.Access) {
        if ($Rule.AccessControlType -ne [System.Security.AccessControl.AccessControlType]::Allow) {
            continue
        }

        $Identity = [string]$Rule.IdentityReference
        $IsBroad = $false

        foreach ($Broad in $BroadIdentities) {
            if ($Identity -ieq $Broad -or $Identity.ToLowerInvariant().EndsWith("\" + $Broad.ToLowerInvariant())) {
                $IsBroad = $true
                break
            }
        }

        if (-not $IsBroad) {
            continue
        }

        if (([int]$Rule.FileSystemRights -band $WriteRights) -ne 0) {
            $Findings.Add(("{0} has {1}" -f $Identity, $Rule.FileSystemRights))
        }
    }

    return $Findings
}

# ---------------------------------------------------------------------------
# Run orchestration
# ---------------------------------------------------------------------------

function Start-ArchiveRun {
    param(
        [pscustomobject]$Options,
        [System.Windows.Forms.Label]$StatusLabel,
        [System.Windows.Forms.ProgressBar]$ProgressBar
    )

    $SourceRoot = Get-NormalisedPath -Path $Options.SourceRoot
    $ArchiveRoot = Get-NormalisedPath -Path $Options.ArchiveRoot

    if ([string]::IsNullOrWhiteSpace($SourceRoot)) {
        [System.Windows.Forms.MessageBox]::Show("Please select or enter the source folder to archive from.", "Missing source path", "OK", "Warning") | Out-Null
        return
    }

    if (-not (Test-Path -LiteralPath $SourceRoot -PathType Container)) {
        [System.Windows.Forms.MessageBox]::Show("The source path does not exist or is not accessible:`r`n`r`n$SourceRoot", "Invalid source path", "OK", "Error") | Out-Null
        return
    }

    if ([string]::IsNullOrWhiteSpace($ArchiveRoot)) {
        [System.Windows.Forms.MessageBox]::Show("Please select or enter the archive destination folder.", "Missing archive path", "OK", "Warning") | Out-Null
        return
    }

    if (Test-PathIsUnder -ParentPath $SourceRoot -ChildPath $ArchiveRoot) {
        [System.Windows.Forms.MessageBox]::Show(
            "The archive location is inside the source folder:`r`n`r`nSource:  $SourceRoot`r`nArchive: $ArchiveRoot`r`n`r`nThat would archive the archive. Choose a destination outside the source tree.",
            "Archive location is inside the source",
            "OK",
            "Error") | Out-Null
        return
    }

    if (Test-PathIsUnder -ParentPath $ArchiveRoot -ChildPath $SourceRoot) {
        [System.Windows.Forms.MessageBox]::Show(
            "The source folder is inside the archive location:`r`n`r`nSource:  $SourceRoot`r`nArchive: $ArchiveRoot`r`n`r`nChoose a source outside the archive tree.",
            "Source is inside the archive",
            "OK",
            "Error") | Out-Null
        return
    }

    if ($Options.ThresholdValue -lt 1) {
        [System.Windows.Forms.MessageBox]::Show("The age threshold must be 1 or greater.", "Invalid threshold", "OK", "Warning") | Out-Null
        return
    }

    if ([string]::IsNullOrWhiteSpace($Options.CsvPath)) {
        [System.Windows.Forms.MessageBox]::Show("Please choose where the CSV report should be saved.", "Missing CSV path", "OK", "Warning") | Out-Null
        return
    }

    if (-not (Test-Path -LiteralPath $ArchiveRoot -PathType Container)) {
        if ($Options.PreviewOnly) {
            # A preview creates nothing at all, not even the archive root.
            [System.Windows.Forms.MessageBox]::Show(
                "The archive folder does not exist yet:`r`n`r`n$ArchiveRoot`r`n`r`nThe preview will still run and will show where each folder would go. The folder is created on the first live run.",
                "Archive folder does not exist yet",
                "OK",
                "Information") | Out-Null
        } else {
            $CreateAnswer = [System.Windows.Forms.MessageBox]::Show(
                "The archive folder does not exist yet:`r`n`r`n$ArchiveRoot`r`n`r`nCreate it now?",
                "Create the archive folder?",
                "YesNo",
                "Question")

            if ($CreateAnswer -ne [System.Windows.Forms.DialogResult]::Yes) {
                return
            }

            try {
                [void][System.IO.Directory]::CreateDirectory($ArchiveRoot)
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Could not create the archive folder:`r`n`r`n$($_.Exception.Message)", "Archive folder", "OK", "Error") | Out-Null
                return
            }
        }
    }

    $CutoffDate = Get-CutoffDate -ThresholdValue $Options.ThresholdValue -ThresholdUnit $Options.ThresholdUnit
    $RunTimestamp = Get-Date

    $EffectiveArchiveRoot = $ArchiveRoot

    if ($Options.GroupRunInDatedFolder) {
        $EffectiveArchiveRoot = Join-Path -Path $ArchiveRoot -ChildPath ("ArchiveRun_{0}" -f $RunTimestamp.ToString("yyyyMMdd_HHmmss"))
    }

    $PermissionFindings = Get-OpenArchivePermissions -Path $ArchiveRoot

    if ($PermissionFindings.Count -gt 0) {
        $PermissionAnswer = [System.Windows.Forms.MessageBox]::Show(
            ("Security check on the archive location:`r`n`r`n{0}`r`n`r`nAfter archiving, the copy in the archive is the only copy of this data. Broad write access there means anyone in those groups can delete or alter it, and the placeholder .txt files tell every user where to look.`r`n`r`nRecommended: restrict the archive share and folder to I.T. with Modify, and give end users read-only or no access at all.`r`n`r`nContinue anyway?" -f (($PermissionFindings | Sort-Object -Unique) -join "`r`n")),
            "Archive location permissions",
            "YesNo",
            "Warning")

        if ($PermissionAnswer -ne [System.Windows.Forms.DialogResult]::Yes) {
            return
        }
    }

    if (-not $Options.PreviewOnly) {
        $ConfirmMessage = @"
This is a LIVE archive run. It will move data and then delete the originals.

  Source folder      : $SourceRoot
  Archive destination: $EffectiveArchiveRoot
  Age rule           : not modified in the last $($Options.ThresholdValue) $($Options.ThresholdUnit)
  Cut-off date       : $($CutoffDate.ToString("yyyy-MM-dd HH:mm:ss"))
  Age measured by    : $($Options.AgeBasis)
  Folder selection   : $($Options.SelectionMode)
  Copy engine        : $($Options.CopyMethod)
  Verification       : $(if ($Options.VerifyWithHash) { "names, sizes and SHA-256 of every file" } else { "names and sizes of every file" })
  Delete method      : $(if ($Options.UseRecycleBin) { "send the original to the Recycle Bin" } else { "permanent delete of the original" })
  Reparse points     : $(if ($Options.SkipLinkedFolders) { "links skipped" } else { "LINKS NOT SKIPPED" }); $(if ($Options.SkipTieredFiles) { "tiered/HSM stubs skipped" } else { "TIERED/HSM STUBS NOT SKIPPED" }); dedup and compression treated as ordinary files
  Folder cap         : $(if ($Options.MaxFolders -gt 0) { "$($Options.MaxFolders) folders this run" } else { "no cap" })

For every matching folder the tool will:
  1. copy the folder to the archive,
  2. verify the copy file for file,
  3. delete the original,
  4. leave a .txt placeholder with a 24 character reference code and a full listing.

Anything that fails to copy or verify is left alone. Nothing is deleted before
its copy verifies.

If you have not already run this in preview mode and read the CSV, cancel now
and do that first.
"@

        if (-not (Show-TypedConfirmation -Message $ConfirmMessage -RequiredWord "ARCHIVE")) {
            $StatusLabel.Text = "Cancelled. Nothing was changed."
            return
        }

        $FinalAnswer = [System.Windows.Forms.MessageBox]::Show(
            "Last check.`r`n`r`nOriginals under`r`n$SourceRoot`r`nwill be deleted once their archive copy has verified.`r`n`r`nStart the live run now?",
            "Start the live archive run?",
            "YesNo",
            "Warning")

        if ($FinalAnswer -ne [System.Windows.Forms.DialogResult]::Yes) {
            $StatusLabel.Text = "Cancelled. Nothing was changed."
            return
        }
    }

    $ExcludePatterns = @()

    if (-not [string]::IsNullOrWhiteSpace($Options.ExcludeText)) {
        $ExcludePatterns = @($Options.ExcludeText -split ';' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }

    $Config = [pscustomobject]@{
        SourceRoot                  = $SourceRoot
        ArchiveRoot                 = $ArchiveRoot
        EffectiveArchiveRoot        = $EffectiveArchiveRoot
        ThresholdValue              = $Options.ThresholdValue
        ThresholdUnit               = $Options.ThresholdUnit
        CutoffDate                  = $CutoffDate
        AgeBasis                    = $Options.AgeBasis
        SelectionMode               = $Options.SelectionMode
        FixedDepth                  = $Options.FixedDepth
        PreviewOnly                 = $Options.PreviewOnly
        CopyMethod                  = $Options.CopyMethod
        CopyPermissions             = $Options.CopyPermissions
        VerifyWithHash              = $Options.VerifyWithHash
        UseRecycleBin               = $Options.UseRecycleBin
        ClearReadOnly               = $Options.ClearReadOnly
        SkipLinkedFolders           = $Options.SkipLinkedFolders
        SkipTieredFiles             = $Options.SkipTieredFiles
        SkipEmptyFolders            = $Options.SkipEmptyFolders
        PreserveStructure           = $Options.PreserveStructure
        GroupRunInDatedFolder       = $Options.GroupRunInDatedFolder
        WriteArchiveSideManifest    = $Options.WriteArchiveSideManifest
        ShowArchivePathInPlaceholder = $Options.ShowArchivePathInPlaceholder
        CleanUpFailedCopies         = $Options.CleanUpFailedCopies
        LogSkipped                  = $Options.LogSkipped
        MaxFolders                  = $Options.MaxFolders
        ContactText                 = $Options.ContactText
        PlaceholderTemplate         = $Options.PlaceholderTemplate
        ExcludePatterns             = $ExcludePatterns
        NewestWriteMap              = (New-Object 'System.Collections.Generic.Dictionary[string, object]')
    }

    $script:CancelRequested = $false

    $Counters = @{
        FoldersExamined   = 0
        Archived          = 0
        WouldArchive      = 0
        Skipped           = 0
        Failed            = 0
        Excluded          = 0
        Warnings          = 0
        FilesArchived     = 0
        BytesArchived     = [long]0
        EnumerationErrors = 0
        LimitReached      = $false
    }

    $Writers = [pscustomobject]@{
        RunCsv   = $null
        IndexCsv = $null
    }

    $ProgressBar.Style = "Marquee"
    $ProgressBar.MarqueeAnimationSpeed = 25
    $StatusLabel.Text = "Starting..."
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $CsvDirectory = Split-Path -Path $Options.CsvPath -Parent

        if (-not [string]::IsNullOrWhiteSpace($CsvDirectory) -and -not (Test-Path -LiteralPath $CsvDirectory -PathType Container)) {
            [void][System.IO.Directory]::CreateDirectory($CsvDirectory)
        }

        $Writers.RunCsv = New-Object System.IO.StreamWriter($Options.CsvPath, $false, [System.Text.Encoding]::UTF8)
        Write-CsvRow -Writer $Writers.RunCsv -Values $script:ArchiveCsvColumns

        if ($Options.AppendToIndex -and -not $Options.PreviewOnly) {
            $IndexPath = Join-Path -Path $ArchiveRoot -ChildPath "_BRC_ArchiveIndex.csv"
            Register-ExistingReferenceCodes -IndexCsvPath $IndexPath

            $IndexExists = Test-Path -LiteralPath $IndexPath -PathType Leaf
            $Writers.IndexCsv = New-Object System.IO.StreamWriter($IndexPath, $true, [System.Text.Encoding]::UTF8)

            if (-not $IndexExists) {
                Write-CsvRow -Writer $Writers.IndexCsv -Values $script:ArchiveCsvColumns
            }
        }

        if ($Options.AgeBasis -eq "Recursive newest item timestamp") {
            $ScannedCount = 0
            $ScannedRef = [ref]$ScannedCount

            [void](Build-NewestWriteMap -CurrentPath $SourceRoot -Map $Config.NewestWriteMap -StatusLabel $StatusLabel -ScannedCount $ScannedRef)

            $StatusLabel.Text = "Assessed $($ScannedRef.Value) folders. Working through the matches..."
            [System.Windows.Forms.Application]::DoEvents()
        }

        Invoke-ArchiveWalk -Config $Config -CurrentPath $SourceRoot -Depth 0 -Writers $Writers -Counters $Counters -StatusLabel $StatusLabel

        $Verb = if ($Options.PreviewOnly) { "would be archived" } else { "archived" }
        $ArchivedTotal = $Counters.Archived + $Counters.WouldArchive

        $SummaryLines = New-Object 'System.Collections.Generic.List[string]'

        if ($script:CancelRequested) {
            [void]$SummaryLines.Add("Run cancelled. Everything completed before the cancel is finished and recorded.")
            [void]$SummaryLines.Add("")
        } elseif ($Counters.LimitReached) {
            [void]$SummaryLines.Add("Stopped at the folder cap of $($Options.MaxFolders). Run again to continue.")
            [void]$SummaryLines.Add("")
        }

        if ($Options.PreviewOnly) {
            [void]$SummaryLines.Add("PREVIEW ONLY - nothing was copied, moved or deleted.")
            [void]$SummaryLines.Add("")
        }

        [void]$SummaryLines.Add("Folders examined     : $($Counters.FoldersExamined)")
        [void]$SummaryLines.Add("Folders $Verb : $ArchivedTotal")
        [void]$SummaryLines.Add("Files included       : $($Counters.FilesArchived)")
        [void]$SummaryLines.Add("Data volume          : $(Format-SizeFriendly -Bytes $Counters.BytesArchived)")
        [void]$SummaryLines.Add("Skipped              : $($Counters.Skipped)")
        [void]$SummaryLines.Add("Excluded by name     : $($Counters.Excluded)")
        [void]$SummaryLines.Add("Failed               : $($Counters.Failed)")
        [void]$SummaryLines.Add("Archived with warnings: $($Counters.Warnings)")
        [void]$SummaryLines.Add("")
        [void]$SummaryLines.Add("Cut-off date         : $($CutoffDate.ToString('yyyy-MM-dd HH:mm:ss'))")
        [void]$SummaryLines.Add("")
        [void]$SummaryLines.Add("CSV report saved to:")
        [void]$SummaryLines.Add($Options.CsvPath)

        $StatusLabel.Text = ("Finished. {0} folders {1}. Skipped {2}. Failed {3}." -f $ArchivedTotal, $Verb, $Counters.Skipped, $Counters.Failed)

        [System.Windows.Forms.MessageBox]::Show(
            ($SummaryLines -join "`r`n"),
            $(if ($Options.PreviewOnly) { "Preview complete" } else { "Archive run complete" }),
            "OK",
            $(if ($Counters.Failed -gt 0) { "Warning" } else { "Information" })) | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "The run stopped with an error:`r`n`r`n$($_.Exception.Message)`r`n`r`nAnything already archived is recorded in the CSV.",
            "Run failed",
            "OK",
            "Error") | Out-Null

        $StatusLabel.Text = "Run failed."
    } finally {
        if ($Writers.RunCsv) {
            $Writers.RunCsv.Close()
            $Writers.RunCsv.Dispose()
        }

        if ($Writers.IndexCsv) {
            $Writers.IndexCsv.Close()
            $Writers.IndexCsv.Dispose()
        }

        $ProgressBar.MarqueeAnimationSpeed = 0
        $ProgressBar.Style = "Blocks"
    }

    if ($Options.OpenCsvWhenComplete -and (Test-Path -LiteralPath $Options.CsvPath -PathType Leaf)) {
        try {
            Start-Process -FilePath $Options.CsvPath | Out-Null
        } catch {
            # Opening the CSV is a convenience, never a reason to report failure.
        }
    }
}

# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

[System.Windows.Forms.Application]::EnableVisualStyles()

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "BRC Folder Archiver"
$Form.Size = New-Object System.Drawing.Size(872, 866)
$Form.StartPosition = "CenterScreen"
$Form.MaximizeBox = $false
$Form.FormBorderStyle = "FixedDialog"

# --- Locations -------------------------------------------------------------

$GroupLocations = New-Object System.Windows.Forms.GroupBox
$GroupLocations.Text = "Locations"
$GroupLocations.Location = New-Object System.Drawing.Point(15, 10)
$GroupLocations.Size = New-Object System.Drawing.Size(825, 185)
$Form.Controls.Add($GroupLocations)

$LabelSource = New-Object System.Windows.Forms.Label
$LabelSource.Text = "Source data folder / UNC share to archive from:"
$LabelSource.Location = New-Object System.Drawing.Point(15, 22)
$LabelSource.Size = New-Object System.Drawing.Size(400, 20)
$GroupLocations.Controls.Add($LabelSource)

$TextSource = New-Object System.Windows.Forms.TextBox
$TextSource.Location = New-Object System.Drawing.Point(15, 44)
$TextSource.Size = New-Object System.Drawing.Size(670, 24)
$GroupLocations.Controls.Add($TextSource)

$ButtonBrowseSource = New-Object System.Windows.Forms.Button
$ButtonBrowseSource.Text = "Browse..."
$ButtonBrowseSource.Location = New-Object System.Drawing.Point(700, 42)
$ButtonBrowseSource.Size = New-Object System.Drawing.Size(95, 28)
$GroupLocations.Controls.Add($ButtonBrowseSource)

$LabelArchive = New-Object System.Windows.Forms.Label
$LabelArchive.Text = "Archive destination folder / UNC share:"
$LabelArchive.Location = New-Object System.Drawing.Point(15, 78)
$LabelArchive.Size = New-Object System.Drawing.Size(400, 20)
$GroupLocations.Controls.Add($LabelArchive)

$TextArchive = New-Object System.Windows.Forms.TextBox
$TextArchive.Location = New-Object System.Drawing.Point(15, 100)
$TextArchive.Size = New-Object System.Drawing.Size(670, 24)
$GroupLocations.Controls.Add($TextArchive)

$ButtonBrowseArchive = New-Object System.Windows.Forms.Button
$ButtonBrowseArchive.Text = "Browse..."
$ButtonBrowseArchive.Location = New-Object System.Drawing.Point(700, 98)
$ButtonBrowseArchive.Size = New-Object System.Drawing.Size(95, 28)
$GroupLocations.Controls.Add($ButtonBrowseArchive)

$LabelCsv = New-Object System.Windows.Forms.Label
$LabelCsv.Text = "CSV report of everything archived:"
$LabelCsv.Location = New-Object System.Drawing.Point(15, 134)
$LabelCsv.Size = New-Object System.Drawing.Size(400, 20)
$GroupLocations.Controls.Add($LabelCsv)

$TextCsv = New-Object System.Windows.Forms.TextBox
$TextCsv.Location = New-Object System.Drawing.Point(15, 156)
$TextCsv.Size = New-Object System.Drawing.Size(670, 24)
$DefaultCsvName = "FolderArchive_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss")
$TextCsv.Text = Join-Path -Path ([Environment]::GetFolderPath("Desktop")) -ChildPath $DefaultCsvName
$GroupLocations.Controls.Add($TextCsv)

$ButtonSaveCsv = New-Object System.Windows.Forms.Button
$ButtonSaveCsv.Text = "Save As..."
$ButtonSaveCsv.Location = New-Object System.Drawing.Point(700, 154)
$ButtonSaveCsv.Size = New-Object System.Drawing.Size(95, 28)
$GroupLocations.Controls.Add($ButtonSaveCsv)

# --- What to archive -------------------------------------------------------

$GroupWhat = New-Object System.Windows.Forms.GroupBox
$GroupWhat.Text = "What to archive"
$GroupWhat.Location = New-Object System.Drawing.Point(15, 200)
$GroupWhat.Size = New-Object System.Drawing.Size(825, 160)
$Form.Controls.Add($GroupWhat)

$LabelThreshold = New-Object System.Windows.Forms.Label
$LabelThreshold.Text = "Not modified in the last:"
$LabelThreshold.Location = New-Object System.Drawing.Point(15, 28)
$LabelThreshold.Size = New-Object System.Drawing.Size(160, 20)
$GroupWhat.Controls.Add($LabelThreshold)

$NumericThreshold = New-Object System.Windows.Forms.NumericUpDown
$NumericThreshold.Location = New-Object System.Drawing.Point(180, 26)
$NumericThreshold.Size = New-Object System.Drawing.Size(80, 24)
$NumericThreshold.Minimum = 1
$NumericThreshold.Maximum = 10000
$NumericThreshold.Value = 3
$GroupWhat.Controls.Add($NumericThreshold)

$ComboUnit = New-Object System.Windows.Forms.ComboBox
$ComboUnit.Location = New-Object System.Drawing.Point(270, 26)
$ComboUnit.Size = New-Object System.Drawing.Size(100, 24)
$ComboUnit.DropDownStyle = "DropDownList"
[void]$ComboUnit.Items.Add("Days")
[void]$ComboUnit.Items.Add("Weeks")
[void]$ComboUnit.Items.Add("Months")
[void]$ComboUnit.Items.Add("Years")
$ComboUnit.SelectedItem = "Years"
$GroupWhat.Controls.Add($ComboUnit)

$LabelBasis = New-Object System.Windows.Forms.Label
$LabelBasis.Text = "Age measured by:"
$LabelBasis.Location = New-Object System.Drawing.Point(400, 28)
$LabelBasis.Size = New-Object System.Drawing.Size(120, 20)
$GroupWhat.Controls.Add($LabelBasis)

$ComboBasis = New-Object System.Windows.Forms.ComboBox
$ComboBasis.Location = New-Object System.Drawing.Point(525, 26)
$ComboBasis.Size = New-Object System.Drawing.Size(270, 24)
$ComboBasis.DropDownStyle = "DropDownList"
[void]$ComboBasis.Items.Add("Recursive newest item timestamp")
[void]$ComboBasis.Items.Add("Folder timestamp only")
$ComboBasis.SelectedItem = "Recursive newest item timestamp"
$GroupWhat.Controls.Add($ComboBasis)

$LabelSelection = New-Object System.Windows.Forms.Label
$LabelSelection.Text = "Folder selection:"
$LabelSelection.Location = New-Object System.Drawing.Point(15, 63)
$LabelSelection.Size = New-Object System.Drawing.Size(160, 20)
$GroupWhat.Controls.Add($LabelSelection)

$ComboSelection = New-Object System.Windows.Forms.ComboBox
$ComboSelection.Location = New-Object System.Drawing.Point(180, 61)
$ComboSelection.Size = New-Object System.Drawing.Size(340, 24)
$ComboSelection.DropDownStyle = "DropDownList"
[void]$ComboSelection.Items.Add("Top-most stale folder")
[void]$ComboSelection.Items.Add("Fixed depth below the source root")
[void]$ComboSelection.Items.Add("Immediate subfolders of the source root only")
$ComboSelection.SelectedItem = "Top-most stale folder"
$GroupWhat.Controls.Add($ComboSelection)

$LabelDepth = New-Object System.Windows.Forms.Label
$LabelDepth.Text = "Depth:"
$LabelDepth.Location = New-Object System.Drawing.Point(540, 63)
$LabelDepth.Size = New-Object System.Drawing.Size(50, 20)
$GroupWhat.Controls.Add($LabelDepth)

$NumericDepth = New-Object System.Windows.Forms.NumericUpDown
$NumericDepth.Location = New-Object System.Drawing.Point(595, 61)
$NumericDepth.Size = New-Object System.Drawing.Size(60, 24)
$NumericDepth.Minimum = 1
$NumericDepth.Maximum = 32
$NumericDepth.Value = 2
$NumericDepth.Enabled = $false
$GroupWhat.Controls.Add($NumericDepth)

$LabelSelectionHelp = New-Object System.Windows.Forms.Label
$LabelSelectionHelp.Text = "Top-most: archive the highest folder that is stale and leave its parents alone. Fixed depth 2 under \\server\data archives \\server\data\2020\01 as a unit."
$LabelSelectionHelp.Location = New-Object System.Drawing.Point(15, 90)
$LabelSelectionHelp.Size = New-Object System.Drawing.Size(795, 18)
$GroupWhat.Controls.Add($LabelSelectionHelp)

$LabelExclude = New-Object System.Windows.Forms.Label
$LabelExclude.Text = "Never archive folders named (semicolon separated, wildcards allowed):"
$LabelExclude.Location = New-Object System.Drawing.Point(15, 112)
$LabelExclude.Size = New-Object System.Drawing.Size(440, 20)
$GroupWhat.Controls.Add($LabelExclude)

$TextExclude = New-Object System.Windows.Forms.TextBox
$TextExclude.Location = New-Object System.Drawing.Point(15, 132)
$TextExclude.Size = New-Object System.Drawing.Size(500, 24)
$TextExclude.Text = "`$RECYCLE.BIN;System Volume Information;_Archive*"
$GroupWhat.Controls.Add($TextExclude)

$LabelMaxFolders = New-Object System.Windows.Forms.Label
$LabelMaxFolders.Text = "Stop after this many folders (0 = no cap):"
$LabelMaxFolders.Location = New-Object System.Drawing.Point(535, 112)
$LabelMaxFolders.Size = New-Object System.Drawing.Size(280, 20)
$GroupWhat.Controls.Add($LabelMaxFolders)

$NumericMaxFolders = New-Object System.Windows.Forms.NumericUpDown
$NumericMaxFolders.Location = New-Object System.Drawing.Point(535, 132)
$NumericMaxFolders.Size = New-Object System.Drawing.Size(90, 24)
$NumericMaxFolders.Minimum = 0
$NumericMaxFolders.Maximum = 1000000
$NumericMaxFolders.Value = 0
$GroupWhat.Controls.Add($NumericMaxFolders)

# --- How to archive --------------------------------------------------------

$GroupHow = New-Object System.Windows.Forms.GroupBox
$GroupHow.Text = "How to archive"
$GroupHow.Location = New-Object System.Drawing.Point(15, 365)
$GroupHow.Size = New-Object System.Drawing.Size(825, 178)
$Form.Controls.Add($GroupHow)

$LabelCopyMethod = New-Object System.Windows.Forms.Label
$LabelCopyMethod.Text = "Copy engine:"
$LabelCopyMethod.Location = New-Object System.Drawing.Point(15, 26)
$LabelCopyMethod.Size = New-Object System.Drawing.Size(90, 20)
$GroupHow.Controls.Add($LabelCopyMethod)

$ComboCopyMethod = New-Object System.Windows.Forms.ComboBox
$ComboCopyMethod.Location = New-Object System.Drawing.Point(110, 24)
$ComboCopyMethod.Size = New-Object System.Drawing.Size(190, 24)
$ComboCopyMethod.DropDownStyle = "DropDownList"
[void]$ComboCopyMethod.Items.Add("Robocopy")
[void]$ComboCopyMethod.Items.Add("PowerShell")
$ComboCopyMethod.SelectedItem = "Robocopy"
$GroupHow.Controls.Add($ComboCopyMethod)

$CheckVerifyHash = New-Object System.Windows.Forms.CheckBox
$CheckVerifyHash.Text = "Verify every file with SHA-256 (slow)"
$CheckVerifyHash.Location = New-Object System.Drawing.Point(320, 24)
$CheckVerifyHash.Size = New-Object System.Drawing.Size(250, 24)
$CheckVerifyHash.Checked = $false
$GroupHow.Controls.Add($CheckVerifyHash)

$CheckCopyPermissions = New-Object System.Windows.Forms.CheckBox
$CheckCopyPermissions.Text = "Copy NTFS permissions (robocopy)"
$CheckCopyPermissions.Location = New-Object System.Drawing.Point(580, 24)
$CheckCopyPermissions.Size = New-Object System.Drawing.Size(235, 24)
$CheckCopyPermissions.Checked = $false
$GroupHow.Controls.Add($CheckCopyPermissions)

$CheckPreserveStructure = New-Object System.Windows.Forms.CheckBox
$CheckPreserveStructure.Text = "Mirror the source folder structure in the archive"
$CheckPreserveStructure.Location = New-Object System.Drawing.Point(15, 54)
$CheckPreserveStructure.Size = New-Object System.Drawing.Size(300, 24)
$CheckPreserveStructure.Checked = $true
$GroupHow.Controls.Add($CheckPreserveStructure)

$CheckDatedFolder = New-Object System.Windows.Forms.CheckBox
$CheckDatedFolder.Text = "Group this run in a dated subfolder"
$CheckDatedFolder.Location = New-Object System.Drawing.Point(320, 54)
$CheckDatedFolder.Size = New-Object System.Drawing.Size(250, 24)
$CheckDatedFolder.Checked = $false
$GroupHow.Controls.Add($CheckDatedFolder)

$CheckRecycleBin = New-Object System.Windows.Forms.CheckBox
$CheckRecycleBin.Text = "Send originals to the Recycle Bin"
$CheckRecycleBin.Location = New-Object System.Drawing.Point(580, 54)
$CheckRecycleBin.Size = New-Object System.Drawing.Size(235, 24)
$CheckRecycleBin.Checked = $false
$GroupHow.Controls.Add($CheckRecycleBin)

$CheckClearReadOnly = New-Object System.Windows.Forms.CheckBox
$CheckClearReadOnly.Text = "Clear read-only flags before removing"
$CheckClearReadOnly.Location = New-Object System.Drawing.Point(15, 82)
$CheckClearReadOnly.Size = New-Object System.Drawing.Size(300, 24)
$CheckClearReadOnly.Checked = $true
$GroupHow.Controls.Add($CheckClearReadOnly)

$CheckSkipLinks = New-Object System.Windows.Forms.CheckBox
$CheckSkipLinks.Text = "Skip folders holding junctions/symlinks"
$CheckSkipLinks.Location = New-Object System.Drawing.Point(320, 82)
$CheckSkipLinks.Size = New-Object System.Drawing.Size(250, 24)
$CheckSkipLinks.Checked = $true
$GroupHow.Controls.Add($CheckSkipLinks)

$CheckSkipEmpty = New-Object System.Windows.Forms.CheckBox
$CheckSkipEmpty.Text = "Skip folders that contain no files"
$CheckSkipEmpty.Location = New-Object System.Drawing.Point(580, 82)
$CheckSkipEmpty.Size = New-Object System.Drawing.Size(235, 24)
$CheckSkipEmpty.Checked = $true
$GroupHow.Controls.Add($CheckSkipEmpty)

$CheckCleanUpFailed = New-Object System.Windows.Forms.CheckBox
$CheckCleanUpFailed.Text = "Remove the part-copy if a folder fails"
$CheckCleanUpFailed.Location = New-Object System.Drawing.Point(15, 110)
$CheckCleanUpFailed.Size = New-Object System.Drawing.Size(300, 24)
$CheckCleanUpFailed.Checked = $true
$GroupHow.Controls.Add($CheckCleanUpFailed)

$CheckLogSkipped = New-Object System.Windows.Forms.CheckBox
$CheckLogSkipped.Text = "List skipped folders in the CSV too"
$CheckLogSkipped.Location = New-Object System.Drawing.Point(320, 110)
$CheckLogSkipped.Size = New-Object System.Drawing.Size(250, 24)
$CheckLogSkipped.Checked = $true
$GroupHow.Controls.Add($CheckLogSkipped)

$CheckOpenCsv = New-Object System.Windows.Forms.CheckBox
$CheckOpenCsv.Text = "Open the CSV when finished"
$CheckOpenCsv.Location = New-Object System.Drawing.Point(580, 110)
$CheckOpenCsv.Size = New-Object System.Drawing.Size(235, 24)
$CheckOpenCsv.Checked = $true
$GroupHow.Controls.Add($CheckOpenCsv)

$CheckSkipTiered = New-Object System.Windows.Forms.CheckBox
$CheckSkipTiered.Text = "Skip cloud-tiered, HSM and unrecognised stub files"
$CheckSkipTiered.Location = New-Object System.Drawing.Point(15, 138)
$CheckSkipTiered.Size = New-Object System.Drawing.Size(340, 24)
$CheckSkipTiered.Checked = $true
$GroupHow.Controls.Add($CheckSkipTiered)

$LabelStubHelp = New-Object System.Windows.Forms.Label
$LabelStubHelp.Text = "Deduplicated and compressed files are ordinary files here and never block a run. Only links, and stubs whose data lives elsewhere, do."
$LabelStubHelp.Location = New-Object System.Drawing.Point(365, 141)
$LabelStubHelp.Size = New-Object System.Drawing.Size(450, 32)
$GroupHow.Controls.Add($LabelStubHelp)

# --- Placeholder file ------------------------------------------------------

$GroupPlaceholder = New-Object System.Windows.Forms.GroupBox
$GroupPlaceholder.Text = "The .txt file left behind, and its match in the archive"
$GroupPlaceholder.Location = New-Object System.Drawing.Point(15, 548)
$GroupPlaceholder.Size = New-Object System.Drawing.Size(825, 135)
$Form.Controls.Add($GroupPlaceholder)

$LabelContact = New-Object System.Windows.Forms.Label
$LabelContact.Text = "I.T. contact details to print in every placeholder file:"
$LabelContact.Location = New-Object System.Drawing.Point(15, 22)
$LabelContact.Size = New-Object System.Drawing.Size(400, 20)
$GroupPlaceholder.Controls.Add($LabelContact)

$TextContact = New-Object System.Windows.Forms.TextBox
$TextContact.Location = New-Object System.Drawing.Point(15, 44)
$TextContact.Size = New-Object System.Drawing.Size(470, 70)
$TextContact.Multiline = $true
$TextContact.ScrollBars = "Vertical"
$TextContact.Text = "I.T. Service Desk"
$GroupPlaceholder.Controls.Add($TextContact)

$LabelTemplate = New-Object System.Windows.Forms.Label
$LabelTemplate.Text = "Placeholder file name:"
$LabelTemplate.Location = New-Object System.Drawing.Point(505, 22)
$LabelTemplate.Size = New-Object System.Drawing.Size(300, 20)
$GroupPlaceholder.Controls.Add($LabelTemplate)

$TextTemplate = New-Object System.Windows.Forms.TextBox
$TextTemplate.Location = New-Object System.Drawing.Point(505, 44)
$TextTemplate.Size = New-Object System.Drawing.Size(305, 24)
$TextTemplate.Text = "{FolderName}_ARCHIVED_{Ref}.txt"
$GroupPlaceholder.Controls.Add($TextTemplate)

$LabelTemplateHelp = New-Object System.Windows.Forms.Label
$LabelTemplateHelp.Text = "Tokens: {FolderName}, {Ref}, {Date}. The archive side always gets <Ref>.txt, named and filled with the same 24 character code."
$LabelTemplateHelp.Location = New-Object System.Drawing.Point(505, 70)
$LabelTemplateHelp.Size = New-Object System.Drawing.Size(305, 40)
$GroupPlaceholder.Controls.Add($LabelTemplateHelp)

$CheckArchiveManifest = New-Object System.Windows.Forms.CheckBox
$CheckArchiveManifest.Text = "Also write <Ref>_manifest.txt in the archive"
$CheckArchiveManifest.Location = New-Object System.Drawing.Point(505, 108)
$CheckArchiveManifest.Size = New-Object System.Drawing.Size(310, 22)
$CheckArchiveManifest.Checked = $true
$GroupPlaceholder.Controls.Add($CheckArchiveManifest)

$CheckShowArchivePath = New-Object System.Windows.Forms.CheckBox
$CheckShowArchivePath.Text = "Print the archive path in the placeholder"
$CheckShowArchivePath.Location = New-Object System.Drawing.Point(15, 112)
$CheckShowArchivePath.Size = New-Object System.Drawing.Size(300, 22)
$CheckShowArchivePath.Checked = $false
$GroupPlaceholder.Controls.Add($CheckShowArchivePath)

$CheckAppendIndex = New-Object System.Windows.Forms.CheckBox
$CheckAppendIndex.Text = "Append to _BRC_ArchiveIndex.csv in the archive root"
$CheckAppendIndex.Location = New-Object System.Drawing.Point(320, 112)
$CheckAppendIndex.Size = New-Object System.Drawing.Size(340, 22)
$CheckAppendIndex.Checked = $true
$GroupPlaceholder.Controls.Add($CheckAppendIndex)

# --- Run -------------------------------------------------------------------

$CheckPreview = New-Object System.Windows.Forms.CheckBox
$CheckPreview.Text = "Preview only - list what would be archived and change nothing"
$CheckPreview.Location = New-Object System.Drawing.Point(20, 693)
$CheckPreview.Size = New-Object System.Drawing.Size(500, 24)
$CheckPreview.Checked = $true
$CheckPreview.Font = New-Object System.Drawing.Font($Form.Font, [System.Drawing.FontStyle]::Bold)
$Form.Controls.Add($CheckPreview)

$ProgressBar = New-Object System.Windows.Forms.ProgressBar
$ProgressBar.Location = New-Object System.Drawing.Point(20, 722)
$ProgressBar.Size = New-Object System.Drawing.Size(820, 20)
$ProgressBar.Style = "Blocks"
$Form.Controls.Add($ProgressBar)

$StatusLabel = New-Object System.Windows.Forms.Label
$StatusLabel.Text = "Ready. Run a preview first and read the CSV before you archive for real."
$StatusLabel.Location = New-Object System.Drawing.Point(20, 748)
$StatusLabel.Size = New-Object System.Drawing.Size(820, 20)
$Form.Controls.Add($StatusLabel)

$ButtonStart = New-Object System.Windows.Forms.Button
$ButtonStart.Text = "Run Preview"
$ButtonStart.Location = New-Object System.Drawing.Point(535, 776)
$ButtonStart.Size = New-Object System.Drawing.Size(100, 30)
$Form.Controls.Add($ButtonStart)

$ButtonCancel = New-Object System.Windows.Forms.Button
$ButtonCancel.Text = "Cancel"
$ButtonCancel.Location = New-Object System.Drawing.Point(640, 776)
$ButtonCancel.Size = New-Object System.Drawing.Size(95, 30)
$ButtonCancel.Enabled = $false
$Form.Controls.Add($ButtonCancel)

$ButtonClose = New-Object System.Windows.Forms.Button
$ButtonClose.Text = "Close"
$ButtonClose.Location = New-Object System.Drawing.Point(745, 776)
$ButtonClose.Size = New-Object System.Drawing.Size(95, 30)
$Form.Controls.Add($ButtonClose)

# --- Event handlers --------------------------------------------------------

function Select-FolderInto {
    param(
        [System.Windows.Forms.TextBox]$TargetBox,
        [string]$Description
    )

    $FolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderDialog.Description = $Description
    $FolderDialog.ShowNewFolderButton = $true

    if (-not [string]::IsNullOrWhiteSpace($TargetBox.Text)) {
        try {
            if (Test-Path -LiteralPath $TargetBox.Text -PathType Container) {
                $FolderDialog.SelectedPath = $TargetBox.Text
            }
        } catch {}
    }

    if ($FolderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $TargetBox.Text = $FolderDialog.SelectedPath
    }
}

$ButtonBrowseSource.Add_Click({
    Select-FolderInto -TargetBox $TextSource -Description "Select the live data folder to archive from"
})

$ButtonBrowseArchive.Add_Click({
    Select-FolderInto -TargetBox $TextArchive -Description "Select the archive destination folder"
})

$ButtonSaveCsv.Add_Click({
    $SaveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveDialog.Title = "Save the archive report"
    $SaveDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $SaveDialog.FileName = [System.IO.Path]::GetFileName($TextCsv.Text)

    $InitialDirectory = Split-Path -Path $TextCsv.Text -Parent

    if (-not [string]::IsNullOrWhiteSpace($InitialDirectory)) {
        try {
            if (Test-Path -LiteralPath $InitialDirectory -PathType Container) {
                $SaveDialog.InitialDirectory = $InitialDirectory
            }
        } catch {}
    }

    if ($SaveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $TextCsv.Text = $SaveDialog.FileName
    }
})

$ComboSelection.Add_SelectedIndexChanged({
    $NumericDepth.Enabled = ($ComboSelection.SelectedItem -eq "Fixed depth below the source root")
})

$CheckPreview.Add_CheckedChanged({
    if ($CheckPreview.Checked) {
        $ButtonStart.Text = "Run Preview"
        $StatusLabel.Text = "Preview mode. Nothing will be copied, moved or deleted."
    } else {
        $ButtonStart.Text = "Archive Now"
        $StatusLabel.Text = "LIVE mode. Matching folders will be copied to the archive and the originals deleted."
    }
})

$ButtonCancel.Add_Click({
    $script:CancelRequested = $true
    $StatusLabel.Text = "Cancelling after the folder in progress finishes..."
})

$ButtonClose.Add_Click({
    $Form.Close()
})

$ButtonStart.Add_Click({
    $Options = [pscustomobject]@{
        SourceRoot                   = $TextSource.Text.Trim()
        ArchiveRoot                  = $TextArchive.Text.Trim()
        CsvPath                      = $TextCsv.Text.Trim()
        ThresholdValue               = [int]$NumericThreshold.Value
        ThresholdUnit                = [string]$ComboUnit.SelectedItem
        AgeBasis                     = [string]$ComboBasis.SelectedItem
        SelectionMode                = [string]$ComboSelection.SelectedItem
        FixedDepth                   = [int]$NumericDepth.Value
        ExcludeText                  = $TextExclude.Text
        MaxFolders                   = [int]$NumericMaxFolders.Value
        CopyMethod                   = [string]$ComboCopyMethod.SelectedItem
        CopyPermissions              = $CheckCopyPermissions.Checked
        VerifyWithHash               = $CheckVerifyHash.Checked
        PreserveStructure            = $CheckPreserveStructure.Checked
        GroupRunInDatedFolder        = $CheckDatedFolder.Checked
        UseRecycleBin                = $CheckRecycleBin.Checked
        ClearReadOnly                = $CheckClearReadOnly.Checked
        SkipLinkedFolders            = $CheckSkipLinks.Checked
        SkipTieredFiles              = $CheckSkipTiered.Checked
        SkipEmptyFolders             = $CheckSkipEmpty.Checked
        CleanUpFailedCopies          = $CheckCleanUpFailed.Checked
        LogSkipped                   = $CheckLogSkipped.Checked
        OpenCsvWhenComplete          = $CheckOpenCsv.Checked
        WriteArchiveSideManifest     = $CheckArchiveManifest.Checked
        ShowArchivePathInPlaceholder = $CheckShowArchivePath.Checked
        AppendToIndex                = $CheckAppendIndex.Checked
        ContactText                  = $TextContact.Text
        PlaceholderTemplate          = $TextTemplate.Text.Trim()
        PreviewOnly                  = $CheckPreview.Checked
    }

    $ButtonStart.Enabled = $false
    $ButtonClose.Enabled = $false
    $ButtonCancel.Enabled = $true

    try {
        Start-ArchiveRun -Options $Options -StatusLabel $StatusLabel -ProgressBar $ProgressBar
    } finally {
        $ButtonStart.Enabled = $true
        $ButtonClose.Enabled = $true
        $ButtonCancel.Enabled = $false
        $script:CancelRequested = $false
    }
})

[void]$Form.ShowDialog()
$Form.Dispose()
