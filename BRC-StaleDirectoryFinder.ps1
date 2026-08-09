<#
BRC Stale Directory Finder
Windows Server 2019 / Windows PowerShell 5.1

What it does:
- Opens a GUI
- Lets you select or type a folder path / UNC share path
- Lets you choose a stale threshold, such as 90 days, 12 months, or 3 years
- Finds directories that have not been modified within that time
- Supports two scan modes:
    1. Folder timestamp only
    2. Recursive newest item timestamp
- Exports results to CSV

Notes:
- "Folder timestamp only" uses the directory's own LastWriteTime.
- "Recursive newest item timestamp" treats a folder as active if any file or subfolder inside it has been modified recently.
- This tool is read-only. It does not modify files or folders.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

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

function Get-DepthFromRoot {
    param(
        [string]$RootPath,
        [string]$CurrentPath
    )

    try {
        $RootFull = [System.IO.Path]::GetFullPath($RootPath).TrimEnd('\')
        $CurrentFull = [System.IO.Path]::GetFullPath($CurrentPath).TrimEnd('\')

        if ($CurrentFull.Length -le $RootFull.Length) {
            return 0
        }

        $Relative = $CurrentFull.Substring($RootFull.Length).TrimStart('\')

        if ([string]::IsNullOrWhiteSpace($Relative)) {
            return 0
        }

        return (($Relative -split '\\').Count)
    } catch {
        return ""
    }
}

function Get-CutoffDate {
    param(
        [int]$ThresholdValue,
        [string]$ThresholdUnit
    )

    $Now = Get-Date

    switch ($ThresholdUnit) {
        "Days"   { return $Now.AddDays(-1 * $ThresholdValue) }
        "Months" { return $Now.AddMonths(-1 * $ThresholdValue) }
        "Years"  { return $Now.AddYears(-1 * $ThresholdValue) }
        default  { return $Now.AddDays(-1 * $ThresholdValue) }
    }
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

function Scan-FolderTimestampOnly {
    param(
        [string]$RootPath,
        [datetime]$CutoffDate,
        [int]$ThresholdValue,
        [string]$ThresholdUnit,
        [bool]$IncludeRoot,
        [bool]$Recurse,
        [bool]$ExportAllFolders,
        [System.IO.StreamWriter]$Writer,
        [System.Windows.Forms.Label]$StatusLabel
    )

    $Queue = New-Object 'System.Collections.Generic.Queue[string]'
    $Queue.Enqueue($RootPath)

    $ScannedCount = 0
    $MatchedCount = 0
    $ErrorCount = 0

    while ($Queue.Count -gt 0) {
        $CurrentPath = $Queue.Dequeue()
        $ScannedCount++

        if (($ScannedCount % 25) -eq 0) {
            $StatusLabel.Text = "Scanned $ScannedCount folders. Matched $MatchedCount stale folders..."
            [System.Windows.Forms.Application]::DoEvents()
        }

        $Errors = New-Object 'System.Collections.Generic.List[string]'
        $DirectFileCount = 0
        $DirectSubfolderCount = 0
        $FolderLastWriteTime = $null
        $ParentPath = ""
        $Depth = Get-DepthFromRoot -RootPath $RootPath -CurrentPath $CurrentPath

        try {
            $DirInfo = New-Object System.IO.DirectoryInfo($CurrentPath)
            $FolderLastWriteTime = $DirInfo.LastWriteTime
            if ($null -ne $DirInfo.Parent) {
                $ParentPath = $DirInfo.Parent.FullName
            }
        } catch {
            $Errors.Add("Could not read directory info: $($_.Exception.Message)")
            $ErrorCount++
        }

        try {
            foreach ($FilePath in [System.IO.Directory]::EnumerateFiles($CurrentPath)) {
                $DirectFileCount++
            }
        } catch {
            $Errors.Add("Could not enumerate direct files: $($_.Exception.Message)")
            $ErrorCount++
        }

        try {
            foreach ($ChildPath in [System.IO.Directory]::EnumerateDirectories($CurrentPath)) {
                $DirectSubfolderCount++
                if ($Recurse) {
                    $Queue.Enqueue($ChildPath)
                }
            }
        } catch {
            $Errors.Add("Could not enumerate child folders: $($_.Exception.Message)")
            $ErrorCount++
        }

        $IsRoot = ([System.IO.Path]::GetFullPath($CurrentPath).TrimEnd('\') -ieq [System.IO.Path]::GetFullPath($RootPath).TrimEnd('\'))
        $LastModifiedUsed = $FolderLastWriteTime
        $IsStale = $false

        if ($null -ne $LastModifiedUsed -and $LastModifiedUsed -lt $CutoffDate) {
            $IsStale = $true
        }

        if (($IncludeRoot -or -not $IsRoot) -and ($ExportAllFolders -or $IsStale)) {
            if ($IsStale) {
                $MatchedCount++
            }

            $RowValues = @(
                $CurrentPath,
                $ParentPath,
                [string]$Depth,
                (Format-DateForCsv -DateValue $FolderLastWriteTime),
                "",
                (Format-DateForCsv -DateValue $LastModifiedUsed),
                [string](Get-DaysSince -DateValue $LastModifiedUsed),
                (Format-DateForCsv -DateValue $CutoffDate),
                [string]$ThresholdValue,
                $ThresholdUnit,
                "Folder timestamp only",
                [string]$IsStale,
                [string]$DirectFileCount,
                [string]$DirectSubfolderCount,
                "",
                "",
                [string]$Errors.Count,
                (($Errors | Sort-Object -Unique) -join "; ")
            )

            Write-CsvRow -Writer $Writer -Values $RowValues
        }
    }

    return [pscustomobject]@{
        ScannedCount = $ScannedCount
        MatchedCount = $MatchedCount
        ErrorCount = $ErrorCount
    }
}

function Scan-RecursiveNewestItem {
    param(
        [string]$CurrentPath,
        [string]$RootPath,
        [datetime]$CutoffDate,
        [int]$ThresholdValue,
        [string]$ThresholdUnit,
        [bool]$IncludeRoot,
        [bool]$ExportAllFolders,
        [System.IO.StreamWriter]$Writer,
        [ref]$ScannedCount,
        [ref]$MatchedCount,
        [ref]$ErrorCount,
        [System.Windows.Forms.Label]$StatusLabel
    )

    $ScannedCount.Value++

    if (($ScannedCount.Value % 25) -eq 0) {
        $StatusLabel.Text = "Scanned $($ScannedCount.Value) folders. Matched $($MatchedCount.Value) stale folders..."
        [System.Windows.Forms.Application]::DoEvents()
    }

    $Errors = New-Object 'System.Collections.Generic.List[string]'
    $DirectFileCount = 0
    $DirectSubfolderCount = 0
    $TotalFileCount = 0
    $TotalSubfolderCount = 0
    $FolderLastWriteTime = $null
    $NewestItemLastWriteTime = $null
    $ParentPath = ""
    $Depth = Get-DepthFromRoot -RootPath $RootPath -CurrentPath $CurrentPath

    try {
        $DirInfo = New-Object System.IO.DirectoryInfo($CurrentPath)
        $FolderLastWriteTime = $DirInfo.LastWriteTime
        $NewestItemLastWriteTime = $FolderLastWriteTime

        if ($null -ne $DirInfo.Parent) {
            $ParentPath = $DirInfo.Parent.FullName
        }
    } catch {
        $Errors.Add("Could not read directory info: $($_.Exception.Message)")
        $ErrorCount.Value++
    }

    try {
        foreach ($FilePath in [System.IO.Directory]::EnumerateFiles($CurrentPath)) {
            $DirectFileCount++
            $TotalFileCount++

            try {
                $FileInfo = New-Object System.IO.FileInfo($FilePath)

                if ($null -eq $NewestItemLastWriteTime -or $FileInfo.LastWriteTime -gt $NewestItemLastWriteTime) {
                    $NewestItemLastWriteTime = $FileInfo.LastWriteTime
                }
            } catch {
                $Errors.Add("Could not read file info for '$FilePath': $($_.Exception.Message)")
                $ErrorCount.Value++
            }
        }
    } catch {
        $Errors.Add("Could not enumerate direct files: $($_.Exception.Message)")
        $ErrorCount.Value++
    }

    $ChildPaths = New-Object 'System.Collections.Generic.List[string]'

    try {
        foreach ($ChildPath in [System.IO.Directory]::EnumerateDirectories($CurrentPath)) {
            $DirectSubfolderCount++
            $TotalSubfolderCount++
            $ChildPaths.Add($ChildPath)
        }
    } catch {
        $Errors.Add("Could not enumerate child folders: $($_.Exception.Message)")
        $ErrorCount.Value++
    }

    foreach ($ChildPath in $ChildPaths) {
        $ChildResult = Scan-RecursiveNewestItem `
            -CurrentPath $ChildPath `
            -RootPath $RootPath `
            -CutoffDate $CutoffDate `
            -ThresholdValue $ThresholdValue `
            -ThresholdUnit $ThresholdUnit `
            -IncludeRoot $IncludeRoot `
            -ExportAllFolders $ExportAllFolders `
            -Writer $Writer `
            -ScannedCount $ScannedCount `
            -MatchedCount $MatchedCount `
            -ErrorCount $ErrorCount `
            -StatusLabel $StatusLabel

        $TotalFileCount += $ChildResult.TotalFileCount
        $TotalSubfolderCount += $ChildResult.TotalSubfolderCount

        if ($null -ne $ChildResult.NewestItemLastWriteTime) {
            if ($null -eq $NewestItemLastWriteTime -or $ChildResult.NewestItemLastWriteTime -gt $NewestItemLastWriteTime) {
                $NewestItemLastWriteTime = $ChildResult.NewestItemLastWriteTime
            }
        }
    }

    $LastModifiedUsed = $NewestItemLastWriteTime
    $IsStale = $false

    if ($null -ne $LastModifiedUsed -and $LastModifiedUsed -lt $CutoffDate) {
        $IsStale = $true
    }

    $IsRoot = ([System.IO.Path]::GetFullPath($CurrentPath).TrimEnd('\') -ieq [System.IO.Path]::GetFullPath($RootPath).TrimEnd('\'))

    if (($IncludeRoot -or -not $IsRoot) -and ($ExportAllFolders -or $IsStale)) {
        if ($IsStale) {
            $MatchedCount.Value++
        }

        $RowValues = @(
            $CurrentPath,
            $ParentPath,
            [string]$Depth,
            (Format-DateForCsv -DateValue $FolderLastWriteTime),
            (Format-DateForCsv -DateValue $NewestItemLastWriteTime),
            (Format-DateForCsv -DateValue $LastModifiedUsed),
            [string](Get-DaysSince -DateValue $LastModifiedUsed),
            (Format-DateForCsv -DateValue $CutoffDate),
            [string]$ThresholdValue,
            $ThresholdUnit,
            "Recursive newest item timestamp",
            [string]$IsStale,
            [string]$DirectFileCount,
            [string]$DirectSubfolderCount,
            [string]$TotalFileCount,
            [string]$TotalSubfolderCount,
            [string]$Errors.Count,
            (($Errors | Sort-Object -Unique) -join "; ")
        )

        Write-CsvRow -Writer $Writer -Values $RowValues
    }

    return [pscustomobject]@{
        TotalFileCount = $TotalFileCount
        TotalSubfolderCount = $TotalSubfolderCount
        NewestItemLastWriteTime = $NewestItemLastWriteTime
    }
}

function Start-StaleDirectoryScan {
    param(
        [string]$RootPath,
        [string]$CsvPath,
        [int]$ThresholdValue,
        [string]$ThresholdUnit,
        [string]$ScanMode,
        [bool]$IncludeRoot,
        [bool]$Recurse,
        [bool]$ExportAllFolders,
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

    if ($ThresholdValue -lt 1) {
        [System.Windows.Forms.MessageBox]::Show("Threshold value must be 1 or greater.", "Invalid threshold", "OK", "Warning") | Out-Null
        return
    }

    $CutoffDate = Get-CutoffDate -ThresholdValue $ThresholdValue -ThresholdUnit $ThresholdUnit

    $ProgressBar.Style = "Marquee"
    $ProgressBar.MarqueeAnimationSpeed = 25
    $StatusLabel.Text = "Starting scan..."
    [System.Windows.Forms.Application]::DoEvents()

    $Columns = @(
        "Path",
        "ParentPath",
        "Depth",
        "FolderLastWriteTime",
        "NewestItemLastWriteTime",
        "LastModifiedUsed",
        "DaysSinceModified",
        "CutoffDate",
        "ThresholdValue",
        "ThresholdUnit",
        "ScanMode",
        "IsStale",
        "DirectFileCount",
        "DirectSubfolderCount",
        "TotalFileCount",
        "TotalSubfolderCount",
        "ErrorCount",
        "Errors"
    )

    $Writer = $null

    try {
        $OutputDirectory = Split-Path -Path $CsvPath -Parent

        if (-not [string]::IsNullOrWhiteSpace($OutputDirectory)) {
            if (-not (Test-Path -LiteralPath $OutputDirectory -PathType Container)) {
                New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
            }
        }

        $Writer = New-Object System.IO.StreamWriter($CsvPath, $false, [System.Text.Encoding]::UTF8)
        Write-CsvRow -Writer $Writer -Values $Columns

        $ScannedCount = 0
        $MatchedCount = 0
        $ErrorCount = 0

        if ($ScanMode -eq "Folder timestamp only" -or -not $Recurse) {
            $Result = Scan-FolderTimestampOnly `
                -RootPath $RootPath `
                -CutoffDate $CutoffDate `
                -ThresholdValue $ThresholdValue `
                -ThresholdUnit $ThresholdUnit `
                -IncludeRoot $IncludeRoot `
                -Recurse $Recurse `
                -ExportAllFolders $ExportAllFolders `
                -Writer $Writer `
                -StatusLabel $StatusLabel

            $ScannedCount = $Result.ScannedCount
            $MatchedCount = $Result.MatchedCount
            $ErrorCount = $Result.ErrorCount
        } else {
            $ScannedRef = [ref]$ScannedCount
            $MatchedRef = [ref]$MatchedCount
            $ErrorRef = [ref]$ErrorCount

            [void](Scan-RecursiveNewestItem `
                -CurrentPath $RootPath `
                -RootPath $RootPath `
                -CutoffDate $CutoffDate `
                -ThresholdValue $ThresholdValue `
                -ThresholdUnit $ThresholdUnit `
                -IncludeRoot $IncludeRoot `
                -ExportAllFolders $ExportAllFolders `
                -Writer $Writer `
                -ScannedCount $ScannedRef `
                -MatchedCount $MatchedRef `
                -ErrorCount $ErrorRef `
                -StatusLabel $StatusLabel)

            $ScannedCount = $ScannedRef.Value
            $MatchedCount = $MatchedRef.Value
            $ErrorCount = $ErrorRef.Value
        }

        $Writer.Flush()

        $StatusLabel.Text = "Complete. Scanned $ScannedCount folders. Matched $MatchedCount stale folders. Errors: $ErrorCount"

        [System.Windows.Forms.MessageBox]::Show(
            "Stale directory scan complete.`r`n`r`nFolders scanned: $ScannedCount`r`nStale folders matched: $MatchedCount`r`nErrors: $ErrorCount`r`n`r`nCutoff date: $($CutoffDate.ToString("yyyy-MM-dd HH:mm:ss"))`r`n`r`nCSV saved to:`r`n$CsvPath",
            "Complete",
            "OK",
            "Information"
        ) | Out-Null
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
    }
}

[System.Windows.Forms.Application]::EnableVisualStyles()

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "BRC Stale Directory Finder"
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

$DefaultCsvName = "StaleDirectories_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss")
$TextCsv.Text = Join-Path -Path ([Environment]::GetFolderPath("Desktop")) -ChildPath $DefaultCsvName

$Form.Controls.Add($TextCsv)

$ButtonCsv = New-Object System.Windows.Forms.Button
$ButtonCsv.Text = "Save As..."
$ButtonCsv.Location = New-Object System.Drawing.Point(680, 108)
$ButtonCsv.Size = New-Object System.Drawing.Size(95, 28)
$Form.Controls.Add($ButtonCsv)

$LabelThreshold = New-Object System.Windows.Forms.Label
$LabelThreshold.Text = "Not modified in the last:"
$LabelThreshold.Location = New-Object System.Drawing.Point(20, 155)
$LabelThreshold.Size = New-Object System.Drawing.Size(170, 20)
$Form.Controls.Add($LabelThreshold)

$NumericThreshold = New-Object System.Windows.Forms.NumericUpDown
$NumericThreshold.Location = New-Object System.Drawing.Point(190, 153)
$NumericThreshold.Size = New-Object System.Drawing.Size(80, 24)
$NumericThreshold.Minimum = 1
$NumericThreshold.Maximum = 10000
$NumericThreshold.Value = 365
$Form.Controls.Add($NumericThreshold)

$ComboUnit = New-Object System.Windows.Forms.ComboBox
$ComboUnit.Location = New-Object System.Drawing.Point(285, 153)
$ComboUnit.Size = New-Object System.Drawing.Size(100, 24)
$ComboUnit.DropDownStyle = "DropDownList"
[void]$ComboUnit.Items.Add("Days")
[void]$ComboUnit.Items.Add("Months")
[void]$ComboUnit.Items.Add("Years")
$ComboUnit.SelectedItem = "Days"
$Form.Controls.Add($ComboUnit)

$LabelMode = New-Object System.Windows.Forms.Label
$LabelMode.Text = "Scan mode:"
$LabelMode.Location = New-Object System.Drawing.Point(20, 195)
$LabelMode.Size = New-Object System.Drawing.Size(170, 20)
$Form.Controls.Add($LabelMode)

$ComboMode = New-Object System.Windows.Forms.ComboBox
$ComboMode.Location = New-Object System.Drawing.Point(190, 193)
$ComboMode.Size = New-Object System.Drawing.Size(300, 24)
$ComboMode.DropDownStyle = "DropDownList"
[void]$ComboMode.Items.Add("Folder timestamp only")
[void]$ComboMode.Items.Add("Recursive newest item timestamp")
$ComboMode.SelectedItem = "Recursive newest item timestamp"
$Form.Controls.Add($ComboMode)

$CheckRecurse = New-Object System.Windows.Forms.CheckBox
$CheckRecurse.Text = "Recurse subfolders"
$CheckRecurse.Location = New-Object System.Drawing.Point(20, 235)
$CheckRecurse.Size = New-Object System.Drawing.Size(180, 24)
$CheckRecurse.Checked = $true
$Form.Controls.Add($CheckRecurse)

$CheckIncludeRoot = New-Object System.Windows.Forms.CheckBox
$CheckIncludeRoot.Text = "Include selected root folder in output"
$CheckIncludeRoot.Location = New-Object System.Drawing.Point(220, 235)
$CheckIncludeRoot.Size = New-Object System.Drawing.Size(260, 24)
$CheckIncludeRoot.Checked = $false
$Form.Controls.Add($CheckIncludeRoot)

$CheckExportAll = New-Object System.Windows.Forms.CheckBox
$CheckExportAll.Text = "Export all folders, not only stale folders"
$CheckExportAll.Location = New-Object System.Drawing.Point(500, 235)
$CheckExportAll.Size = New-Object System.Drawing.Size(280, 24)
$CheckExportAll.Checked = $false
$Form.Controls.Add($CheckExportAll)

$ModeHelp = New-Object System.Windows.Forms.Label
$ModeHelp.Text = "Tip: Recursive mode is best for archive planning because a folder is only stale if nothing inside it has changed recently."
$ModeHelp.Location = New-Object System.Drawing.Point(20, 270)
$ModeHelp.Size = New-Object System.Drawing.Size(755, 34)
$Form.Controls.Add($ModeHelp)

$ProgressBar = New-Object System.Windows.Forms.ProgressBar
$ProgressBar.Location = New-Object System.Drawing.Point(20, 315)
$ProgressBar.Size = New-Object System.Drawing.Size(755, 24)
$ProgressBar.Style = "Blocks"
$Form.Controls.Add($ProgressBar)

$StatusLabel = New-Object System.Windows.Forms.Label
$StatusLabel.Text = "Ready."
$StatusLabel.Location = New-Object System.Drawing.Point(20, 350)
$StatusLabel.Size = New-Object System.Drawing.Size(755, 24)
$Form.Controls.Add($StatusLabel)

$ButtonStart = New-Object System.Windows.Forms.Button
$ButtonStart.Text = "Start Scan"
$ButtonStart.Location = New-Object System.Drawing.Point(575, 385)
$ButtonStart.Size = New-Object System.Drawing.Size(95, 30)
$Form.Controls.Add($ButtonStart)

$ButtonClose = New-Object System.Windows.Forms.Button
$ButtonClose.Text = "Close"
$ButtonClose.Location = New-Object System.Drawing.Point(680, 385)
$ButtonClose.Size = New-Object System.Drawing.Size(95, 30)
$Form.Controls.Add($ButtonClose)

$ButtonBrowse.Add_Click({
    $FolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderDialog.Description = "Select the folder or share root to scan"
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
    $SaveDialog.Title = "Save stale directory CSV"
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

$ButtonStart.Add_Click({
    $ButtonStart.Enabled = $false
    $ButtonClose.Enabled = $false

    try {
        Start-StaleDirectoryScan `
            -RootPath $TextPath.Text `
            -CsvPath $TextCsv.Text `
            -ThresholdValue ([int]$NumericThreshold.Value) `
            -ThresholdUnit ([string]$ComboUnit.SelectedItem) `
            -ScanMode ([string]$ComboMode.SelectedItem) `
            -IncludeRoot $CheckIncludeRoot.Checked `
            -Recurse $CheckRecurse.Checked `
            -ExportAllFolders $CheckExportAll.Checked `
            -StatusLabel $StatusLabel `
            -ProgressBar $ProgressBar
    } finally {
        $ButtonStart.Enabled = $true
        $ButtonClose.Enabled = $true
    }
})

$ButtonClose.Add_Click({
    $Form.Close()
})

[void]$Form.ShowDialog()
