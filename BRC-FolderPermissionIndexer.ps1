Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ------------------------------------------------------------
# Global caches
# ------------------------------------------------------------

$script:IdentityTypeCache = @{}
$script:ExpandedGroupCache = @{}
$script:HasADModule = $false

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    $script:HasADModule = $true
} catch {
    $script:HasADModule = $false
}

# ------------------------------------------------------------
# Helper functions
# ------------------------------------------------------------

function New-StringSet {
    return New-Object 'System.Collections.Generic.HashSet[string]'
}

function Add-SetValue {
    param(
        [System.Collections.Generic.HashSet[string]]$Set,
        [string]$Value
    )

    if (-not [string]::IsNullOrWhiteSpace($Value)) {
        [void]$Set.Add($Value)
    }
}

function Join-Set {
    param(
        [System.Collections.Generic.HashSet[string]]$Set
    )

    if ($null -eq $Set -or $Set.Count -eq 0) {
        return ""
    }

    return (($Set | Sort-Object) -join "; ")
}

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

function Test-WriteLikeRights {
    param(
        [System.Security.AccessControl.FileSystemRights]$Rights
    )

    $WriteLikeBits = @(
        [System.Security.AccessControl.FileSystemRights]::WriteData,
        [System.Security.AccessControl.FileSystemRights]::AppendData,
        [System.Security.AccessControl.FileSystemRights]::WriteExtendedAttributes,
        [System.Security.AccessControl.FileSystemRights]::WriteAttributes,
        [System.Security.AccessControl.FileSystemRights]::Delete,
        [System.Security.AccessControl.FileSystemRights]::DeleteSubdirectoriesAndFiles,
        [System.Security.AccessControl.FileSystemRights]::ChangePermissions,
        [System.Security.AccessControl.FileSystemRights]::TakeOwnership
    )

    foreach ($Bit in $WriteLikeBits) {
        if (($Rights -band $Bit) -ne 0) {
            return $true
        }
    }

    return $false
}

function Test-ReadLikeRights {
    param(
        [System.Security.AccessControl.FileSystemRights]$Rights
    )

    $ReadLikeBits = @(
        [System.Security.AccessControl.FileSystemRights]::ReadData,
        [System.Security.AccessControl.FileSystemRights]::ReadAttributes,
        [System.Security.AccessControl.FileSystemRights]::ReadExtendedAttributes,
        [System.Security.AccessControl.FileSystemRights]::ReadPermissions,
        [System.Security.AccessControl.FileSystemRights]::ExecuteFile
    )

    foreach ($Bit in $ReadLikeBits) {
        if (($Rights -band $Bit) -ne 0) {
            return $true
        }
    }

    return $false
}

function Get-PermissionClass {
    param(
        [System.Security.AccessControl.FileSystemRights]$Rights
    )

    if (Test-WriteLikeRights -Rights $Rights) {
        return "ReadWrite"
    }

    if (Test-ReadLikeRights -Rights $Rights) {
        return "ReadOnly"
    }

    return "Other"
}

function Get-IdentitySid {
    param(
        [string]$Identity
    )

    try {
        $NtAccount = New-Object System.Security.Principal.NTAccount($Identity)
        $Sid = $NtAccount.Translate([System.Security.Principal.SecurityIdentifier])
        return $Sid.Value
    } catch {
        return $null
    }
}

function Get-IdentityType {
    param(
        [string]$Identity
    )

    if ($script:IdentityTypeCache.ContainsKey($Identity)) {
        return $script:IdentityTypeCache[$Identity]
    }

    $Result = "Other"
    $Lower = $Identity.ToLowerInvariant()

    # Common well-known group-like identities
    $KnownGroupPatterns = @(
        "everyone",
        "authenticated users",
        "domain users",
        "domain admins",
        "enterprise admins",
        "schema admins",
        "administrators",
        "users",
        "backup operators",
        "server operators",
        "account operators",
        "print operators",
        "remote desktop users",
        "network service",
        "interactive",
        "creator owner"
    )

    foreach ($Pattern in $KnownGroupPatterns) {
        if ($Lower -like "*$Pattern*") {
            $script:IdentityTypeCache[$Identity] = "Group"
            return "Group"
        }
    }

    if ($Lower -like "*\system" -or $Lower -eq "nt authority\system") {
        $script:IdentityTypeCache[$Identity] = "Other"
        return "Other"
    }

    if ($Identity -match '\$$') {
        $script:IdentityTypeCache[$Identity] = "Computer"
        return "Computer"
    }

    $SidValue = Get-IdentitySid -Identity $Identity

    if ($SidValue) {
        # Built-in local groups generally live under S-1-5-32
        if ($SidValue -like "S-1-5-32-*") {
            $script:IdentityTypeCache[$Identity] = "Group"
            return "Group"
        }

        if ($script:HasADModule) {
            try {
                $AdObject = Get-ADObject -Filter "objectSid -eq '$SidValue'" -Properties objectClass -ErrorAction Stop

                if ($AdObject.objectClass -eq "group") {
                    $Result = "Group"
                } elseif ($AdObject.objectClass -eq "user") {
                    $Result = "User"
                } elseif ($AdObject.objectClass -eq "computer") {
                    $Result = "Computer"
                } else {
                    $Result = "Other"
                }

                $script:IdentityTypeCache[$Identity] = $Result
                return $Result
            } catch {
                # Fall through to ADSI / name-based fallback
            }
        }
    }

    # ADSI fallback for local/domain accounts
    try {
        $Parts = $Identity -split "\\", 2

        if ($Parts.Count -eq 2) {
            $DomainOrComputer = $Parts[0]
            $Sam = $Parts[1]

            $AdsiPath = "WinNT://$DomainOrComputer/$Sam"
            $AdsiObj = [ADSI]$AdsiPath
            $SchemaClassName = [string]$AdsiObj.SchemaClassName

            if ($SchemaClassName -eq "Group") {
                $Result = "Group"
            } elseif ($SchemaClassName -eq "User") {
                $Result = "User"
            } else {
                $Result = "Other"
            }
        }
    } catch {
        # Final fallback below
    }

    if ($Result -eq "Other") {
        # Most remaining explicit ACL identities are users or service identities.
        # Keep service/special identities as Other where possible.
        if ($Lower -like "nt authority\*" -or $Lower -like "builtin\*") {
            $Result = "Other"
        } else {
            $Result = "User"
        }
    }

    $script:IdentityTypeCache[$Identity] = $Result
    return $Result
}

function Expand-GroupToUsers {
    param(
        [string]$GroupIdentity
    )

    if (-not $script:HasADModule) {
        return @()
    }

    if ($script:ExpandedGroupCache.ContainsKey($GroupIdentity)) {
        return $script:ExpandedGroupCache[$GroupIdentity]
    }

    $Users = @()

    try {
        $SidValue = Get-IdentitySid -Identity $GroupIdentity

        if ($SidValue) {
            $GroupObject = Get-ADObject -Filter "objectSid -eq '$SidValue'" -Properties distinguishedName, objectClass -ErrorAction Stop

            if ($GroupObject.objectClass -eq "group") {
                $Members = Get-ADGroupMember -Identity $GroupObject.DistinguishedName -Recursive -ErrorAction Stop

                $Users = $Members |
                    Where-Object { $_.objectClass -eq "user" } |
                    Select-Object -ExpandProperty SamAccountName |
                    Sort-Object -Unique
            }
        }
    } catch {
        $Users = @()
    }

    $script:ExpandedGroupCache[$GroupIdentity] = $Users
    return $Users
}

function Get-FolderPermissionSummary {
    param(
        [string]$FolderPath,
        [bool]$IncludeInherited,
        [bool]$ExpandAdGroups
    )

    $ReadOnlyGroups = New-StringSet
    $ReadOnlyUsers = New-StringSet
    $ReadOnlyExpandedUsers = New-StringSet

    $ReadWriteGroups = New-StringSet
    $ReadWriteUsers = New-StringSet
    $ReadWriteExpandedUsers = New-StringSet

    $OtherAllowEntries = New-StringSet
    $DenyEntries = New-StringSet
    $ErrorEntries = New-StringSet

    $Owner = ""

    try {
        $Acl = Get-Acl -LiteralPath $FolderPath -ErrorAction Stop
        $Owner = $Acl.Owner

        foreach ($Rule in $Acl.Access) {
            if (-not $IncludeInherited -and $Rule.IsInherited) {
                continue
            }

            $Identity = [string]$Rule.IdentityReference.Value
            $Rights = $Rule.FileSystemRights
            $AccessType = [string]$Rule.AccessControlType
            $InheritedText = if ($Rule.IsInherited) { "Inherited" } else { "Explicit" }

            if ($AccessType -eq "Deny") {
                Add-SetValue -Set $DenyEntries -Value "$Identity [$Rights][$InheritedText]"
                continue
            }

            if ($AccessType -ne "Allow") {
                continue
            }

            $PermissionClass = Get-PermissionClass -Rights $Rights
            $IdentityType = Get-IdentityType -Identity $Identity

            if ($PermissionClass -eq "ReadOnly") {
                if ($IdentityType -eq "Group") {
                    Add-SetValue -Set $ReadOnlyGroups -Value $Identity

                    if ($ExpandAdGroups) {
                        foreach ($ExpandedUser in (Expand-GroupToUsers -GroupIdentity $Identity)) {
                            Add-SetValue -Set $ReadOnlyExpandedUsers -Value $ExpandedUser
                        }
                    }
                } elseif ($IdentityType -eq "User") {
                    Add-SetValue -Set $ReadOnlyUsers -Value $Identity
                } else {
                    Add-SetValue -Set $OtherAllowEntries -Value "$Identity [$Rights][$InheritedText]"
                }
            } elseif ($PermissionClass -eq "ReadWrite") {
                if ($IdentityType -eq "Group") {
                    Add-SetValue -Set $ReadWriteGroups -Value $Identity

                    if ($ExpandAdGroups) {
                        foreach ($ExpandedUser in (Expand-GroupToUsers -GroupIdentity $Identity)) {
                            Add-SetValue -Set $ReadWriteExpandedUsers -Value $ExpandedUser
                        }
                    }
                } elseif ($IdentityType -eq "User") {
                    Add-SetValue -Set $ReadWriteUsers -Value $Identity
                } else {
                    Add-SetValue -Set $OtherAllowEntries -Value "$Identity [$Rights][$InheritedText]"
                }
            } else {
                Add-SetValue -Set $OtherAllowEntries -Value "$Identity [$Rights][$InheritedText]"
            }
        }
    } catch {
        Add-SetValue -Set $ErrorEntries -Value $_.Exception.Message
    }

    return [ordered]@{
        Path                    = $FolderPath
        Owner                   = $Owner
        ReadOnly_Groups         = Join-Set -Set $ReadOnlyGroups
        ReadOnly_Users          = Join-Set -Set $ReadOnlyUsers
        ReadOnly_ExpandedUsers  = Join-Set -Set $ReadOnlyExpandedUsers
        ReadWrite_Groups        = Join-Set -Set $ReadWriteGroups
        ReadWrite_Users         = Join-Set -Set $ReadWriteUsers
        ReadWrite_ExpandedUsers = Join-Set -Set $ReadWriteExpandedUsers
        Deny_Entries            = Join-Set -Set $DenyEntries
        Other_Allow_Entries     = Join-Set -Set $OtherAllowEntries
        Errors                  = Join-Set -Set $ErrorEntries
    }
}

function Start-PermissionIndex {
    param(
        [string]$RootPath,
        [string]$CsvPath,
        [bool]$IncludeInherited,
        [bool]$Recurse,
        [bool]$ExpandAdGroups,
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

    if ($ExpandAdGroups -and -not $script:HasADModule) {
        [System.Windows.Forms.MessageBox]::Show(
            "AD group expansion was selected, but the ActiveDirectory module is not available on this server.`r`n`r`nThe scan will continue, but expanded user columns may be empty.",
            "ActiveDirectory module not found",
            "OK",
            "Warning"
        ) | Out-Null
    }

    $ProgressBar.Style = "Marquee"
    $ProgressBar.MarqueeAnimationSpeed = 25
    $StatusLabel.Text = "Starting scan..."
    [System.Windows.Forms.Application]::DoEvents()

    $Columns = @(
        "Path",
        "Owner",
        "ReadOnly_Groups",
        "ReadOnly_Users",
        "ReadOnly_ExpandedUsers",
        "ReadWrite_Groups",
        "ReadWrite_Users",
        "ReadWrite_ExpandedUsers",
        "Deny_Entries",
        "Other_Allow_Entries",
        "Errors"
    )

    $Writer = $null
    $FolderCount = 0
    $ErrorCount = 0

    try {
        $OutputDirectory = Split-Path -Path $CsvPath -Parent

        if (-not [string]::IsNullOrWhiteSpace($OutputDirectory)) {
            if (-not (Test-Path -LiteralPath $OutputDirectory -PathType Container)) {
                New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
            }
        }

        $Writer = New-Object System.IO.StreamWriter($CsvPath, $false, [System.Text.Encoding]::UTF8)
        Write-CsvRow -Writer $Writer -Values $Columns

        $Queue = New-Object 'System.Collections.Generic.Queue[string]'
        $Queue.Enqueue($RootPath)

        while ($Queue.Count -gt 0) {
            $CurrentPath = $Queue.Dequeue()
            $FolderCount++

            if (($FolderCount % 10) -eq 0) {
                $StatusLabel.Text = "Indexed $FolderCount folders..."
                [System.Windows.Forms.Application]::DoEvents()
            }

            $Summary = Get-FolderPermissionSummary `
                -FolderPath $CurrentPath `
                -IncludeInherited $IncludeInherited `
                -ExpandAdGroups $ExpandAdGroups

            if (-not [string]::IsNullOrWhiteSpace($Summary.Errors)) {
                $ErrorCount++
            }

            $Row = foreach ($Column in $Columns) {
                [string]$Summary[$Column]
            }

            Write-CsvRow -Writer $Writer -Values $Row

            if ($Recurse) {
                try {
                    foreach ($Child in [System.IO.Directory]::EnumerateDirectories($CurrentPath)) {
                        $Queue.Enqueue($Child)
                    }
                } catch {
                    $ErrorCount++

                    $ErrorSummary = [ordered]@{
                        Path                    = $CurrentPath
                        Owner                   = ""
                        ReadOnly_Groups         = ""
                        ReadOnly_Users          = ""
                        ReadOnly_ExpandedUsers  = ""
                        ReadWrite_Groups        = ""
                        ReadWrite_Users         = ""
                        ReadWrite_ExpandedUsers = ""
                        Deny_Entries            = ""
                        Other_Allow_Entries     = ""
                        Errors                  = "Could not enumerate child folders: $($_.Exception.Message)"
                    }

                    $ErrorRow = foreach ($Column in $Columns) {
                        [string]$ErrorSummary[$Column]
                    }

                    Write-CsvRow -Writer $Writer -Values $ErrorRow
                }
            }
        }

        $Writer.Flush()

        $StatusLabel.Text = "Complete. Indexed $FolderCount folders. Errors: $ErrorCount"
        [System.Windows.Forms.MessageBox]::Show(
            "Permission indexing complete.`r`n`r`nFolders indexed: $FolderCount`r`nErrors: $ErrorCount`r`n`r`nCSV saved to:`r`n$CsvPath",
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

# ------------------------------------------------------------
# GUI
# ------------------------------------------------------------

[System.Windows.Forms.Application]::EnableVisualStyles()

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "BRC Folder Permission Indexer"
$Form.Size = New-Object System.Drawing.Size(760, 360)
$Form.StartPosition = "CenterScreen"
$Form.MaximizeBox = $false
$Form.FormBorderStyle = "FixedDialog"

$LabelPath = New-Object System.Windows.Forms.Label
$LabelPath.Text = "Folder path / UNC share path:"
$LabelPath.Location = New-Object System.Drawing.Point(20, 25)
$LabelPath.Size = New-Object System.Drawing.Size(220, 20)
$Form.Controls.Add($LabelPath)

$TextPath = New-Object System.Windows.Forms.TextBox
$TextPath.Location = New-Object System.Drawing.Point(20, 50)
$TextPath.Size = New-Object System.Drawing.Size(600, 24)
$TextPath.Text = ""
$Form.Controls.Add($TextPath)

$ButtonBrowse = New-Object System.Windows.Forms.Button
$ButtonBrowse.Text = "Browse..."
$ButtonBrowse.Location = New-Object System.Drawing.Point(630, 48)
$ButtonBrowse.Size = New-Object System.Drawing.Size(90, 28)
$Form.Controls.Add($ButtonBrowse)

$LabelCsv = New-Object System.Windows.Forms.Label
$LabelCsv.Text = "Output CSV file:"
$LabelCsv.Location = New-Object System.Drawing.Point(20, 90)
$LabelCsv.Size = New-Object System.Drawing.Size(220, 20)
$Form.Controls.Add($LabelCsv)

$TextCsv = New-Object System.Windows.Forms.TextBox
$TextCsv.Location = New-Object System.Drawing.Point(20, 115)
$TextCsv.Size = New-Object System.Drawing.Size(600, 24)

$DefaultCsvName = "FolderPermissions_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss")
$TextCsv.Text = Join-Path -Path ([Environment]::GetFolderPath("Desktop")) -ChildPath $DefaultCsvName

$Form.Controls.Add($TextCsv)

$ButtonCsv = New-Object System.Windows.Forms.Button
$ButtonCsv.Text = "Save As..."
$ButtonCsv.Location = New-Object System.Drawing.Point(630, 113)
$ButtonCsv.Size = New-Object System.Drawing.Size(90, 28)
$Form.Controls.Add($ButtonCsv)

$CheckRecurse = New-Object System.Windows.Forms.CheckBox
$CheckRecurse.Text = "Recurse subfolders"
$CheckRecurse.Location = New-Object System.Drawing.Point(20, 160)
$CheckRecurse.Size = New-Object System.Drawing.Size(180, 24)
$CheckRecurse.Checked = $true
$Form.Controls.Add($CheckRecurse)

$CheckInherited = New-Object System.Windows.Forms.CheckBox
$CheckInherited.Text = "Include inherited permissions"
$CheckInherited.Location = New-Object System.Drawing.Point(220, 160)
$CheckInherited.Size = New-Object System.Drawing.Size(220, 24)
$CheckInherited.Checked = $true
$Form.Controls.Add($CheckInherited)

$CheckExpandGroups = New-Object System.Windows.Forms.CheckBox
$CheckExpandGroups.Text = "Expand AD groups to users"
$CheckExpandGroups.Location = New-Object System.Drawing.Point(470, 160)
$CheckExpandGroups.Size = New-Object System.Drawing.Size(220, 24)
$CheckExpandGroups.Checked = $false
$Form.Controls.Add($CheckExpandGroups)

$ProgressBar = New-Object System.Windows.Forms.ProgressBar
$ProgressBar.Location = New-Object System.Drawing.Point(20, 205)
$ProgressBar.Size = New-Object System.Drawing.Size(700, 24)
$ProgressBar.Style = "Blocks"
$Form.Controls.Add($ProgressBar)

$StatusLabel = New-Object System.Windows.Forms.Label
$StatusLabel.Text = "Ready."
$StatusLabel.Location = New-Object System.Drawing.Point(20, 240)
$StatusLabel.Size = New-Object System.Drawing.Size(700, 24)
$Form.Controls.Add($StatusLabel)

$ButtonStart = New-Object System.Windows.Forms.Button
$ButtonStart.Text = "Start Index"
$ButtonStart.Location = New-Object System.Drawing.Point(520, 275)
$ButtonStart.Size = New-Object System.Drawing.Size(95, 30)
$Form.Controls.Add($ButtonStart)

$ButtonClose = New-Object System.Windows.Forms.Button
$ButtonClose.Text = "Close"
$ButtonClose.Location = New-Object System.Drawing.Point(625, 275)
$ButtonClose.Size = New-Object System.Drawing.Size(95, 30)
$Form.Controls.Add($ButtonClose)

# ------------------------------------------------------------
# GUI events
# ------------------------------------------------------------

$ButtonBrowse.Add_Click({
    $FolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderDialog.Description = "Select the folder or share root to index"
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
    $SaveDialog.Title = "Save permissions CSV"
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
        Start-PermissionIndex `
            -RootPath $TextPath.Text `
            -CsvPath $TextCsv.Text `
            -IncludeInherited $CheckInherited.Checked `
            -Recurse $CheckRecurse.Checked `
            -ExpandAdGroups $CheckExpandGroups.Checked `
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