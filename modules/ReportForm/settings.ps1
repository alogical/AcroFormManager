<#
.SYNOPSIS
    Settings management form.

.DESCRIPTION
    Windows form for managing the Device module settings.

.NOTES
    Author: Daniel K. Ives
    Email:  daniel.ives@live.com
#>

param(
    $settings = $null
)

Add-Type -AssemblyName System.Windows.Forms

$InvocationPath  = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Definition)

## SECTION 1 : VALIDATION ##

# Settings File #
$settingsPath = Join-Path $InvocationPath 'settings.json'

if (!$settings -and (Test-Path -LiteralPath $settingsPath)) {
    $settings = ConvertFrom-Json ((Get-Content $settingsPath) -join '')
}

if (!$settings) {
    $settings = @{}

    # Non-Configurable Settings #

        # Semantic Version [Major.Minor.Patch]
        # See documentation on Semantic Versioning or <https://semver.org/>
        $settings.version = "1.0.0"

        # Local Database Location
        $settings.localdb = "$InvocationPath\..\..\database\devicedb"

    # Configurable Settings #
    
        # Remote Database Location (shared information source)
        $settings.remotedb = ""

        # TreeView Display Options; these are set by the SortedTreeView
        # settings pannel
        $settings.treeview = @{}
            $settings.treeview.node = [String]::Empty
            $settings.treeview.groups = New-Object System.Collections.ArrayList
}

## SECTION 2 : FORM DEFINITION ##
$form = New-Object System.Windows.Forms.Form
    $form.Width = 600
    $form.Height = 300

    # Non-Resizable dialog
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog

### Remote Database -----------------------------------------------------------
function Dialog-RemotePath ([System.Windows.Forms.Button]$sender, [System.EventArgs]$e) {
    $browseDialog = New-Object System.Windows.Forms.OpenFileDialog
    $browseDialog.Title = "Remote Shared Device Database File Selection"
    $browseDialog.Multiselect = $false
    
    # Currently only CSV flat databases are supported.
    $browseDialog.Filter = "CSV Database (*.csv)|*.csv"
    $browseDialog.FilterIndex = 1

    if ($browseDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $settings.remotedb =  $browseDialog.FileName
        $sender.Display.Text = $browseDialog.FileName
    }
}

# Dialog Control
$sRemoteDatabasePath = & {
    $margin = 10
    $wMax = $form.ClientRectangle.Width
    $wRem = $wMax

    # Elements
    $title = New-Object System.Windows.Forms.Label
        $title.Text = "Remote Device List Database File"

    $label = New-Object System.Windows.Forms.Label
        $label.Text = "Path"

    $button = New-Object System.Windows.Forms.Button
        $button.Text = "Browse"
        $button.Add_Click({Dialog-RemotePath $this $_})

    $display = New-Object System.Windows.Forms.TextBox
        $display.Text = $settings.remotedb

        Add-Member -InputObject $button -MemberType NoteProperty -Name Display -Value $display

    # Sizing
    $title.Width = 200
    $label.Width = 60; $wRem -= 60
    $button.Width = 60; $wRem -= 60
    $display.Width = $wRem - $margin * 2

    # Positioning
    $left = 0

    $title.Location = New-Object System.Drawing.Point($left,0)

    $label.Location = New-Object System.Drawing.Point($left,30)
    $left += $label.Width + $margin

    $display.Location = New-Object System.Drawing.Point($left,30)
    $left += $display.Width + $margin

    $button.Location = New-Object System.Drawing.Point($left,30)
    
    # Settings Pannel
    $panel = New-Object System.Windows.Forms.Panel
        $panel.Dock = [System.Windows.Forms.DockStyle]::Top
        $panel.BackColor = [System.Drawing.Color]::White

        $panel.Width = $wMax
        $panel.Height = 60

        [void]$panel.Controls.Add($title)
        [void]$panel.Controls.Add($label)
        [void]$panel.Controls.Add($display)
        [void]$panel.Controls.Add($button)

    return $panel
}

    [void]$form.Controls.Add($sRemoteDatabasePath)

## SECTION 3 : GET/SAVE SETTINGS ##

[void]$form.ShowDialog()
      $form.Dispose()

ConvertTo-Json $settings > $settingsPath

return $settings