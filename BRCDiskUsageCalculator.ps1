<#
BRC Disk Usage Calculator
Windows Server 2019 / Windows PowerShell 5.1

What it does:
- Opens a GUI
- Lets you select or type a folder path / UNC share path
- Walks the tree and measures the size of every directory and every file
- Exports one CSV containing a row per directory and a row per file

Notes:
- Directory rows carry both a direct size (files sitting in that folder) and a
  total size (that folder plus everything underneath it).
- File rows carry the file length in the same size columns so the whole CSV can
  be sorted by SizeBytes to find the biggest items regardless of type.
- Reparse points (junctions and symlinks) are not traversed by default so their
  target data is not counted twice.
- This tool is read-only. It does not modify files or folders.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ------------------------------------------------------------
# Global state
# ------------------------------------------------------------

$script:CancelRequested = $false

# ------------------------------------------------------------
# Helper functions
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
# Pass 1: measure the tree
# ------------------------------------------------------------

function Measure-DirectoryTree {
    param(
        [string]$CurrentPath,
        [string]$RootPath,
        [hashtable]$Map,
        [bool]$Recurse,
        [bool]$FollowReparsePoints,
        [ref]$FolderCount,
        [ref]$FileCount,
        [ref]$ErrorCount,
        [System.Windows.Forms.Label]$StatusLabel
    )

    $FolderCount.Value++

    if (($FolderCount.Value % 25) -eq 0) {
        $StatusLabel.Text = "Measuring... $($FolderCount.Value) folders, $($FileCount.Value) files scanned."
        [System.Windows.Forms.Application]::DoEvents()
    }

    $Errors = New-Object 'System.Collections.Generic.List[string]'

    $DirectSizeBytes = [long]0
    $DirectFileCount = 0
    $DirectSubfolderCount = 0

    $TotalSizeBytes = [long]0
    $TotalFileCount = 0
    $TotalSubfolderCount = 0

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
        $ErrorCount.Value++
    }

    $IsRoot = Test-SamePath -FirstPath $CurrentPath -SecondPath $RootPath
    $Traversed = $true

    if ($IsReparsePoint -and -not $FollowReparsePoints -and -not $IsRoot) {
        $Traversed = $false
        $Errors.Add("Reparse point was not traversed. Enable 'Follow reparse points' to include its contents.")
    }

    $ChildPaths = New-Object 'System.Collections.Generic.List[string]'

    if ($Traversed) {
        try {
            foreach ($FilePath in [System.IO.Directory]::EnumerateFiles($CurrentPath)) {
                if ($script:CancelRequested) {
                    break
                }

                $DirectFileCount++
                $FileCount.Value++

                try {
                    $FileInfo = New-Object System.IO.FileInfo($FilePath)
                    $DirectSizeBytes += [long]$FileInfo.Length
                } catch {
                    $Errors.Add("Could not read file info for '$FilePath': $($_.Exception.Message)")
                    $ErrorCount.Value++
                }
            }
        } catch {
            $Errors.Add("Could not enumerate files: $($_.Exception.Message)")
            $ErrorCount.Value++
        }

        try {
            foreach ($ChildPath in [System.IO.Directory]::EnumerateDirectories($CurrentPath)) {
                $DirectSubfolderCount++
                $ChildPaths.Add($ChildPath)
            }
        } catch {
            $Errors.Add("Could not enumerate subfolders: $($_.Exception.Message)")
            $ErrorCount.Value++
        }
    }

    $TotalSizeBytes = $DirectSizeBytes
    $TotalFileCount = $DirectFileCount
    $TotalSubfolderCount = $DirectSubfolderCount

    if ($Recurse) {
        foreach ($ChildPath in $ChildPaths) {
            if ($script:CancelRequested) {
                break
            }

            $ChildResult = Measure-DirectoryTree `
                -CurrentPath $ChildPath `
                -RootPath $RootPath `
                -Map $Map `
                -Recurse $Recurse `
                -FollowReparsePoints $FollowReparsePoints `
                -FolderCount $FolderCount `
                -FileCount $FileCount `
                -ErrorCount $ErrorCount `
                -StatusLabel $StatusLabel

            $TotalSizeBytes += [long]$ChildResult.TotalSizeBytes
            $TotalFileCount += [int]$ChildResult.TotalFileCount
            $TotalSubfolderCount += [int]$ChildResult.TotalSubfolderCount
        }
    }

    $Map[$CurrentPath] = [pscustomobject]@{
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
        Traversed            = $Traversed
        ChildPaths           = $ChildPaths
        Errors               = $Errors
    }

    return [pscustomobject]@{
        TotalSizeBytes      = $TotalSizeBytes
        TotalFileCount      = $TotalFileCount
        TotalSubfolderCount = $TotalSubfolderCount
    }
}

# ------------------------------------------------------------
# Pass 2: write the CSV
# ------------------------------------------------------------

function Write-DirectoryRow {
    param(
        [System.IO.StreamWriter]$Writer,
        [psobject]$Entry,
        [string]$RootPath,
        [long]$ScanTotalBytes
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
        (Get-PercentOfTotal -Bytes $Entry.TotalSizeBytes -TotalBytes $ScanTotalBytes),
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
        [string]$RootPath,
        [long]$ScanTotalBytes
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
        (Get-PercentOfTotal -Bytes $SizeBytes -TotalBytes $ScanTotalBytes),
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

function Export-DiskUsageRows {
    param(
        [string]$RootPath,
        [hashtable]$Map,
        [long]$ScanTotalBytes,
        [bool]$IncludeRoot,
        [bool]$IncludeDirectories,
        [bool]$IncludeFiles,
        [long]$MinFileSizeBytes,
        [System.IO.StreamWriter]$Writer,
        [System.Windows.Forms.Label]$StatusLabel,
        [ref]$ErrorCount
    )

    $DirectoryRows = 0
    $FileRows = 0
    $Processed = 0

    $Queue = New-Object 'System.Collections.Generic.Queue[string]'
    $Queue.Enqueue($RootPath)

    while ($Queue.Count -gt 0) {
        if ($script:CancelRequested) {
            break
        }

        $CurrentPath = $Queue.Dequeue()
        $Processed++

        if (($Processed % 25) -eq 0) {
            $StatusLabel.Text = "Writing CSV... $DirectoryRows directory rows, $FileRows file rows."
            [System.Windows.Forms.Application]::DoEvents()
        }

        $Entry = $Map[$CurrentPath]

        if ($null -eq $Entry) {
            continue
        }

        $IsRoot = Test-SamePath -FirstPath $CurrentPath -SecondPath $RootPath

        if ($IncludeDirectories -and ($IncludeRoot -or -not $IsRoot)) {
            Write-DirectoryRow `
                -Writer $Writer `
                -Entry $Entry `
                -RootPath $RootPath `
                -ScanTotalBytes $ScanTotalBytes

            $DirectoryRows++
        }

        if ($IncludeFiles -and $Entry.Traversed) {
            try {
                foreach ($FilePath in [System.IO.Directory]::EnumerateFiles($CurrentPath)) {
                    if ($script:CancelRequested) {
                        break
                    }

                    try {
                        $FileInfo = New-Object System.IO.FileInfo($FilePath)

                        if ([long]$FileInfo.Length -lt $MinFileSizeBytes) {
                            continue
                        }

                        Write-FileRow `
                            -Writer $Writer `
                            -FileInfo $FileInfo `
                            -ParentPath $CurrentPath `
                            -RootPath $RootPath `
                            -ScanTotalBytes $ScanTotalBytes

                        $FileRows++

                        if (($FileRows % 500) -eq 0) {
                            $StatusLabel.Text = "Writing CSV... $DirectoryRows directory rows, $FileRows file rows."
                            [System.Windows.Forms.Application]::DoEvents()
                        }
                    } catch {
                        $ErrorCount.Value++

                        Write-ErrorRow `
                            -Writer $Writer `
                            -Path $FilePath `
                            -RootPath $RootPath `
                            -Message "Could not read file info: $($_.Exception.Message)"
                    }
                }
            } catch {
                # A folder that already failed during the measure pass has its
                # error recorded on its own row, so do not report it twice.
                if ($Entry.Errors.Count -eq 0) {
                    $ErrorCount.Value++

                    Write-ErrorRow `
                        -Writer $Writer `
                        -Path $CurrentPath `
                        -RootPath $RootPath `
                        -Message "Could not enumerate files: $($_.Exception.Message)"
                }
            }
        }

        foreach ($ChildPath in $Entry.ChildPaths) {
            if ($Map.ContainsKey($ChildPath)) {
                $Queue.Enqueue($ChildPath)
            }
        }
    }

    return [pscustomobject]@{
        DirectoryRows = $DirectoryRows
        FileRows      = $FileRows
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

    $script:CancelRequested = $false

    $ProgressBar.Style = "Marquee"
    $ProgressBar.MarqueeAnimationSpeed = 25
    $StatusLabel.Text = "Starting scan..."
    [System.Windows.Forms.Application]::DoEvents()

    $Columns = @(
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

    $Writer = $null

    $StartTime = Get-Date
    $FolderCount = 0
    $FileCount = 0
    $ErrorCount = 0

    try {
        $OutputDirectory = Split-Path -Path $CsvPath -Parent

        if (-not [string]::IsNullOrWhiteSpace($OutputDirectory)) {
            if (-not (Test-Path -LiteralPath $OutputDirectory -PathType Container)) {
                New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
            }
        }

        $Map = @{}

        $FolderRef = [ref]$FolderCount
        $FileRef = [ref]$FileCount
        $ErrorRef = [ref]$ErrorCount

        $StatusLabel.Text = "Measuring folder sizes..."
        [System.Windows.Forms.Application]::DoEvents()

        $RootResult = Measure-DirectoryTree `
            -CurrentPath $RootPath `
            -RootPath $RootPath `
            -Map $Map `
            -Recurse $Recurse `
            -FollowReparsePoints $FollowReparsePoints `
            -FolderCount $FolderRef `
            -FileCount $FileRef `
            -ErrorCount $ErrorRef `
            -StatusLabel $StatusLabel

        $FolderCount = $FolderRef.Value
        $FileCount = $FileRef.Value

        $ScanTotalBytes = [long]$RootResult.TotalSizeBytes

        $Writer = New-Object System.IO.StreamWriter($CsvPath, $false, [System.Text.Encoding]::UTF8)
        Write-CsvRow -Writer $Writer -Values $Columns

        $StatusLabel.Text = "Writing CSV..."
        [System.Windows.Forms.Application]::DoEvents()

        $ExportResult = Export-DiskUsageRows `
            -RootPath $RootPath `
            -Map $Map `
            -ScanTotalBytes $ScanTotalBytes `
            -IncludeRoot $IncludeRoot `
            -IncludeDirectories $IncludeDirectories `
            -IncludeFiles $IncludeFiles `
            -MinFileSizeBytes $MinFileSizeBytes `
            -Writer $Writer `
            -StatusLabel $StatusLabel `
            -ErrorCount $ErrorRef

        $ErrorCount = $ErrorRef.Value

        $Writer.Flush()

        $Elapsed = (Get-Date) - $StartTime
        $ElapsedText = "{0:hh\:mm\:ss}" -f $Elapsed

        $CompletionState = if ($script:CancelRequested) { "Cancelled" } else { "Complete" }

        $StatusLabel.Text = "$CompletionState. $FolderCount folders, $FileCount files, $(Format-Bytes -Bytes $ScanTotalBytes). Errors: $ErrorCount"

        [System.Windows.Forms.MessageBox]::Show(
            "Disk usage scan $($CompletionState.ToLower()).`r`n`r`nFolders measured: $FolderCount`r`nFiles measured: $FileCount`r`nTotal size: $(Format-Bytes -Bytes $ScanTotalBytes) ($ScanTotalBytes bytes)`r`n`r`nDirectory rows written: $($ExportResult.DirectoryRows)`r`nFile rows written: $($ExportResult.FileRows)`r`nErrors: $ErrorCount`r`nElapsed: $ElapsedText`r`n`r`nCSV saved to:`r`n$CsvPath",
            $CompletionState,
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
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "The scan failed:`r`n`r`n$($_.Exception.Message)",
            "Scan failed",
            "OK",
            "Error"
        ) | Out-Null

        $StatusLabel.Text = "Scan failed."
    } finally {
        if ($Writer) {
            $Writer.Close()
            $Writer.Dispose()
        }

        $ProgressBar.MarqueeAnimationSpeed = 0
        $ProgressBar.Style = "Blocks"
        $script:CancelRequested = $false
    }
}

# ------------------------------------------------------------
# GUI
# ------------------------------------------------------------

[System.Windows.Forms.Application]::EnableVisualStyles()

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "BRC Disk Usage Calculator"
$Form.Size = New-Object System.Drawing.Size(820, 455)
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

$ModeHelp = New-Object System.Windows.Forms.Label
$ModeHelp.Text = "Tip: directory rows report both DirectSizeBytes (files in that folder only) and TotalSizeBytes (folder plus everything under it). Sort the CSV by SizeBytes to see the largest folders and files together."
$ModeHelp.Location = New-Object System.Drawing.Point(20, 255)
$ModeHelp.Size = New-Object System.Drawing.Size(755, 45)
$Form.Controls.Add($ModeHelp)

$ProgressBar = New-Object System.Windows.Forms.ProgressBar
$ProgressBar.Location = New-Object System.Drawing.Point(20, 310)
$ProgressBar.Size = New-Object System.Drawing.Size(755, 24)
$ProgressBar.Style = "Blocks"
$Form.Controls.Add($ProgressBar)

$StatusLabel = New-Object System.Windows.Forms.Label
$StatusLabel.Text = "Ready."
$StatusLabel.Location = New-Object System.Drawing.Point(20, 345)
$StatusLabel.Size = New-Object System.Drawing.Size(755, 24)
$Form.Controls.Add($StatusLabel)

$ButtonStart = New-Object System.Windows.Forms.Button
$ButtonStart.Text = "Start Scan"
$ButtonStart.Location = New-Object System.Drawing.Point(470, 380)
$ButtonStart.Size = New-Object System.Drawing.Size(95, 30)
$Form.Controls.Add($ButtonStart)

$ButtonCancel = New-Object System.Windows.Forms.Button
$ButtonCancel.Text = "Cancel"
$ButtonCancel.Location = New-Object System.Drawing.Point(575, 380)
$ButtonCancel.Size = New-Object System.Drawing.Size(95, 30)
$ButtonCancel.Enabled = $false
$Form.Controls.Add($ButtonCancel)

$ButtonClose = New-Object System.Windows.Forms.Button
$ButtonClose.Text = "Close"
$ButtonClose.Location = New-Object System.Drawing.Point(680, 380)
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
    $StatusLabel.Text = "Cancelling..."
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
