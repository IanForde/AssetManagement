Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# =============== CONFIGURATION ===============
# Local Git repository path (clone your repo here manually first)
$localRepoPath = "C:\AssetManagement"   # Change if your local repo path is different
$localCsvPath = Join-Path -Path $localRepoPath -ChildPath "AssetList.csv"

# Git branch to push/pull
$gitBranch = 'main'
# =============================================

function Run-GitCommand {
    param (
        [string]$Arguments
    )
    $processInfo = New-Object System.Diagnostics.ProcessStartInfo
    $processInfo.FileName = "git"
    $processInfo.Arguments = $Arguments
    $processInfo.WorkingDirectory = $localRepoPath
    $processInfo.RedirectStandardOutput = $true
    $processInfo.RedirectStandardError = $true
    $processInfo.UseShellExecute = $false
    $processInfo.CreateNoWindow = $true

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $processInfo
    $process.Start() | Out-Null
    $output = $process.StandardOutput.ReadToEnd()
    $errorOutput = $process.StandardError.ReadToEnd()
    $process.WaitForExit()

    return @{ ExitCode = $process.ExitCode; StdOut = $output; StdErr = $errorOutput }
}

function Commit-And-Push {
    try {
        # Stage the CSV
        $addResult = Run-GitCommand "add AssetList.csv"
        if ($addResult.ExitCode -ne 0) {
            throw "Git add failed: $($addResult.StdErr)"
        }

        # Commit changes - allow empty commit to avoid errors if no changes
        $commitMessage = "Update AssetList.csv by script - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        $commitResult = Run-GitCommand "commit -m `"$commitMessage`" --allow-empty"
        if ($commitResult.ExitCode -ne 0 -and ($commitResult.StdErr -notmatch "nothing to commit")) {
            throw "Git commit failed: $($commitResult.StdErr)"
        }

        # Push to remote
        $pushResult = Run-GitCommand "push origin $gitBranch"
        if ($pushResult.ExitCode -ne 0) {
            throw "Git push failed: $($pushResult.StdErr)"
        }

        [System.Windows.Forms.MessageBox]::Show("Changes pushed successfully to GitHub.","Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Git operation failed:`n$_","Git Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

function New-LabeledTextBox($labelText, $top, $width=250) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $labelText
    $label.AutoSize = $true
    $label.Top = $top
    $label.Left = 10

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Top = $top - 3
    $textBox.Left = 110
    $textBox.Width = $width

    return ,@($label, $textBox)
}

function Load-Assets {
    if (-Not (Test-Path $localCsvPath)) {
        "Asset Tag,Model,Serial Number,Assigned To,Description" | Out-File -FilePath $localCsvPath -Encoding UTF8
    }
    $assets = Import-Csv -Path $localCsvPath
    if ($null -eq $assets) { return @() }
    return @($assets)
}

function Save-Assets($assets) {
    try {
        $assets | Export-Csv -Path $localCsvPath -NoTypeInformation -Encoding UTF8
        Commit-And-Push
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error saving asset data: $_","Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

function Show-LookupAssetForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Lookup Asset"
    $form.Size = New-Object System.Drawing.Size(400,200)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Enter Asset Tag or Serial Number:"
    $label.AutoSize = $true
    $label.Top = 20
    $label.Left = 10
    $form.Controls.Add($label)

    $inputBox = New-Object System.Windows.Forms.TextBox
    $inputBox.Top = 50
    $inputBox.Left = 10
    $inputBox.Width = 360
    $form.Controls.Add($inputBox)

    $btnSearch = New-Object System.Windows.Forms.Button
    $btnSearch.Text = "Search"
    $btnSearch.Top = 90
    $btnSearch.Left = 10
    $btnSearch.Width = 80
    $form.Controls.Add($btnSearch)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Top = 90
    $btnCancel.Left = 100
    $btnCancel.Width = 80
    $form.Controls.Add($btnCancel)

    $btnSearch.Add_Click({
        $query = $inputBox.Text.Trim()
        if ([string]::IsNullOrEmpty($query)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter an Asset Tag or Serial Number.","Input Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $assets = Load-Assets
        $match = $assets | Where-Object { `
            $_.'Asset Tag' -like "*$query*" -or $_.'Serial Number' -like "*$query*" }

        if ($match.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No matching asset found.","Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } elseif ($match.Count -eq 1 -or ($match.Count -eq $null -and $match)) {
            $form.Close()
            Show-EditAssetForm -Asset ($match | Select-Object -First 1)
        } else {
            if ($match.Count -eq $null -and $match) { $match = @($match) }
            $form.Close()
            Show-SelectAssetForm -Matches $match
        }
    })

    $btnCancel.Add_Click({ $form.Close() })
    $form.ShowDialog() | Out-Null
}

function Show-SelectAssetForm {
    param(
        [Parameter(Mandatory=$true)]
        [System.Collections.IEnumerable] $Matches
    )
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select Asset"
    $form.Size = New-Object System.Drawing.Size(500,400)
    $form.StartPosition = "CenterScreen"

    $listView = New-Object System.Windows.Forms.ListView
    $listView.View = 'Details'
    $listView.FullRowSelect = $true
    $listView.GridLines = $true
    $listView.MultiSelect = $false
    $listView.Width = 460
    $listView.Height = 300
    $listView.Top = 10
    $listView.Left = 10

    $listView.Columns.Add("Asset Tag", 100) | Out-Null
    $listView.Columns.Add("Model", 100) | Out-Null
    $listView.Columns.Add("Serial Number", 120) | Out-Null
    $listView.Columns.Add("Assigned To", 100) | Out-Null

    foreach ($asset in $Matches) {
        $item = New-Object System.Windows.Forms.ListViewItem($asset.'Asset Tag')
        $item.SubItems.Add($asset.Model) | Out-Null
        $item.SubItems.Add($asset.'Serial Number') | Out-Null
        $item.SubItems.Add($asset.'Assigned To') | Out-Null
        $listView.Items.Add($item) | Out-Null
    }

    $form.Controls.Add($listView)

    $btnSelect = New-Object System.Windows.Forms.Button
    $btnSelect.Text = "Select"
    $btnSelect.Top = 320
    $btnSelect.Left = 10
    $btnSelect.Width = 80
    $form.Controls.Add($btnSelect)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Top = 320
    $btnCancel.Left = 100
    $btnCancel.Width = 80
    $form.Controls.Add($btnCancel)

    $btnSelect.Add_Click({
        if ($listView.SelectedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select an asset from the list.","Selection Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $idx = $listView.SelectedItems[0].Index
        $asset = $Matches[$idx]
        $form.Close()
        Show-EditAssetForm -Asset $asset
    })
    $btnCancel.Add_Click({ $form.Close() })
    $form.ShowDialog() | Out-Null
}

function Show-EditAssetForm {
    param (
        [Parameter(Mandatory=$true)]
        [psobject] $Asset
    )
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Edit Asset"
    $form.Size = New-Object System.Drawing.Size(400,350)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

    $controls = @{}
    $fields = @('Asset Tag','Model','Serial Number','Assigned To','Description')
    for ($i=0; $i -lt $fields.Count; $i++) {
        $lblTextbox = New-LabeledTextBox $fields[$i] (20 + $i*40)
        $form.Controls.AddRange($lblTextbox)
        $controls[$fields[$i]] = $lblTextbox[1]
        $controls[$fields[$i]].Text = $Asset.$($fields[$i])
    }

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = "Save"
    $btnSave.Top = 240
    $btnSave.Left = 80
    $btnSave.Width = 100
    $form.Controls.Add($btnSave)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Top = 240
    $btnCancel.Left = 200
    $btnCancel.Width = 100
    $form.Controls.Add($btnCancel)

    $btnSave.Add_Click({
        if ([string]::IsNullOrEmpty($controls['Asset Tag'].Text.Trim())) {
            [System.Windows.Forms.MessageBox]::Show("Asset Tag cannot be empty.","Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $assets = Load-Assets

        if ($assets -eq $null) { $assets = @() }
        elseif (-not ($assets -is [System.Collections.IEnumerable])) { $assets = @($assets) }

        $updated = $false
        for ($i=0; $i -lt $assets.Count; $i++) {
            if ($assets[$i].'Asset Tag' -eq $Asset.'Asset Tag' -and $assets[$i].'Serial Number' -eq $Asset.'Serial Number') {
                foreach ($fld in $fields) {
                    $assets[$i].$fld = $controls[$fld].Text.Trim()
                }
                $updated = $true
                break
            }
        }
        if (-not $updated) {
            $newObj = New-Object PSObject -Property @{
                'Asset Tag'    = $controls['Asset Tag'].Text.Trim()
                'Model'        = $controls['Model'].Text.Trim()
                'Serial Number'= $controls['Serial Number'].Text.Trim()
                'Assigned To'  = $controls['Assigned To'].Text.Trim()
                'Description'  = $controls['Description'].Text.Trim()
            }
            $assets += $newObj
        }
        Save-Assets $assets
        $form.Close()
    })

    $btnCancel.Add_Click({ $form.Close() })
    $form.ShowDialog() | Out-Null
}

function Show-CreateAssetForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Create New Asset"
    $form.Size = New-Object System.Drawing.Size(400,350)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

    $controls = @{}
    $fields = @('Asset Tag','Model','Serial Number','Assigned To','Description')
    for ($i=0; $i -lt $fields.Count; $i++) {
        $lblTextbox = New-LabeledTextBox $fields[$i] (20 + $i*40)
        $form.Controls.AddRange($lblTextbox)
        $controls[$fields[$i]] = $lblTextbox[1]
    }

    $btnAdd = New-Object System.Windows.Forms.Button
    $btnAdd.Text = "Add Asset"
    $btnAdd.Top = 240
    $btnAdd.Left = 80
    $btnAdd.Width = 100
    $form.Controls.Add($btnAdd)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Top = 240
    $btnCancel.Left = 200
    $btnCancel.Width = 100
    $form.Controls.Add($btnCancel)

    $btnAdd.Add_Click({
        if ([string]::IsNullOrEmpty($controls['Asset Tag'].Text.Trim())) {
            [System.Windows.Forms.MessageBox]::Show("Asset Tag cannot be empty.","Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $assets = Load-Assets

        if ($assets -eq $null) { $assets = @() }
        elseif (-not ($assets -is [System.Collections.IEnumerable])) { $assets = @($assets) }

        $existing = $assets | Where-Object {
            $_.'Asset Tag' -eq $controls['Asset Tag'].Text.Trim() -or
            $_.'Serial Number' -eq $controls['Serial Number'].Text.Trim()
        }
        if ($existing.Count -gt 0) {
            $confirm = [System.Windows.Forms.MessageBox]::Show("Asset Tag or Serial Number already exists. Add anyway?","Duplicate Warning",[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Warning)
            if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        }
        $newObj = New-Object PSObject -Property @{
            'Asset Tag'    = $controls['Asset Tag'].Text.Trim()
            'Model'        = $controls['Model'].Text.Trim()
            'Serial Number'= $controls['Serial Number'].Text.Trim()
            'Assigned To'  = $controls['Assigned To'].Text.Trim()
            'Description'  = $controls['Description'].Text.Trim()
        }
        $assets += $newObj
        Save-Assets $assets
        $form.Close()
    })

    $btnCancel.Add_Click({ $form.Close() })
    $form.ShowDialog() | Out-Null
}

function Show-ViewAllAssets {
    if (Test-Path $localCsvPath) {
        Start-Process notepad.exe $localCsvPath
    } else {
        [System.Windows.Forms.MessageBox]::Show("Local CSV file not found.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

function Show-MainForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Asset Management System"
    $form.Size = New-Object System.Drawing.Size(320,270)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

    $notice = New-Object System.Windows.Forms.Label
    $notice.Text = "Using local Git repo - changes auto-pushed to GitHub."
    $notice.AutoSize = $true
    $notice.ForeColor = "DarkGreen"
    $notice.Top = 8
    $notice.Left = 20
    $form.Controls.Add($notice)

    $btnLookup = New-Object System.Windows.Forms.Button
    $btnLookup.Text = "Lookup Asset"
    $btnLookup.Width = 220
    $btnLookup.Height = 40
    $btnLookup.Top = 40
    $btnLookup.Left = 40
    $form.Controls.Add($btnLookup)

    $btnCreate = New-Object System.Windows.Forms.Button
    $btnCreate.Text = "Create Asset"
    $btnCreate.Width = 220
    $btnCreate.Height = 40
    $btnCreate.Top = 100
    $btnCreate.Left = 40
    $form.Controls.Add($btnCreate)

    $btnViewAll = New-Object System.Windows.Forms.Button
    $btnViewAll.Text = "View All Assets"
    $btnViewAll.Width = 220
    $btnViewAll.Height = 40
    $btnViewAll.Top = 160
    $btnViewAll.Left = 40
    $form.Controls.Add($btnViewAll)

    $btnLookup.Add_Click({ Show-LookupAssetForm })
    $btnCreate.Add_Click({ Show-CreateAssetForm })
    $btnViewAll.Add_Click({ Show-ViewAllAssets })

    $form.ShowDialog() | Out-Null
}


# ========== START APP =============

# Auto pull latest changes before starting UI
$pullResult = Run-GitCommand "pull origin $gitBranch"
if ($pullResult.ExitCode -ne 0) {
    [System.Windows.Forms.MessageBox]::Show("Warning: Git pull failed:`n$($pullResult.StdErr)","Warning",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)
}

Show-MainForm
