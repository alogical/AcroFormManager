<#
.SYNOPSIS
    DD2875 Report Viewer.

.DESCRIPTION

.NOTES
    Author: Daniel K. Ives
    Email:  daniel.ives@live.com
#>

Add-Type -AssemblyName System.Windows.Forms

$ModuleInvocationPath  = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Definition)

###############################################################################
###############################################################################
## SECTION 01 ## PUBILC FUNCTIONS AND VARIABLES
##
## Pass-thru Export-ModuleMember calls export all functions and variables
## to the global session that were passed to this modules session from nested
## modules.
###############################################################################
###############################################################################

function Show-ManagerConsole {
    param(
        [Parameter(Mandatory = $false)]
        [System.Collections.ArrayList]
            $Items
    )

    $OnLoad = New-Object System.Collections.ArrayList

    ###############################################################################
    # Window Definition
    $MainForm = New-Object System.Windows.Forms.Form
    $MainForm.Text          = "PDF Form Manager"
    $MainForm.Size          = New-Object System.Drawing.Size(1200, 600)
    $MainForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    $MainForm.Add_Load({
        param($sender, $e)
        $TabContainer.BringToFront()

        foreach ($handler in $OnLoad) {
            $handler.Load()
        }
    })

    ###############################################################################
    # Main Menu
    $Menu = New-Object System.Windows.Forms.MenuStrip
        $Menu.Dock = [System.Windows.Forms.DockStyle]::Top
        $Menu.Name = 'MainMenu'

        $MainForm.MainMenuStrip = $Menu
        [void]$MainForm.Controls.Add($Menu)

    ###############################################################################
    # Base Content Tab Control Container Definitions
    $TabContainer = New-Object System.Windows.Forms.TabControl
        $TabContainer.Dock = [System.Windows.Forms.DockStyle]::Fill
    
        [void]$MainForm.Controls.Add( $TabContainer )

    ###############################################################################
    # Register Components
    Initialize-Components $MainForm $TabContainer $Menu $OnLoad

    ###############################################################################
    # Open Window
    [void]$MainForm.ShowDialog()

    $MainForm.Dispose()
}

Export-ModuleMember -Function *

###############################################################################
###############################################################################
## SECTION 02 ## PRIVATE FUNCTIONS AND VARIABLES
##
## No function or variable in this section is exported unless done so by an
## explicit call to Export-ModuleMember
###############################################################################
###############################################################################
$ACRS = "$env:USERPROFILE\documents\WindowsPowerShell\Programs\ACRS\"
Import-Module "$ACRS\modules\DataTreeView\DataTreeView.psm1" -Prefix Flow
Import-Module "$ACRS\modules\DataListView\DataListView.psm1" -Prefix List
Import-Module "$ACRS\modules\DD2875\dd2875.psm1"

$ImagePath = "$ACRS\resources"

###############################################################################
# Static Objects and Scriptblocks
###############################################################################

###############################################################################
# Main Menu Definitions
###############################################################################
### File Menu -------------------------------------------------------------
$Menu = @{}
$Menu.Scan = New-Object System.Windows.Forms.ToolStripMenuItem("Scan Documents", $null, {
    param($sender, $e)
    $this.Enabled = $false
    Set-ViewData @(Scan-Folder) $this.View $this.Component
    $this.Enabled = $true
})
Add-Member -InputObject $Menu.Scan -MemberType NoteProperty -Name View      -Value $null
Add-Member -InputObject $Menu.Scan -MemberType NoteProperty -Name Component -Value $null

$Menu.ChangeView = @{}
$Menu.ChangeView.DropDown = New-Object System.Windows.Forms.ToolStripMenuItem("View")
Add-Member -InputObject $Menu.ChangeView.DropDown -MemberType NoteProperty -Name View      -Value $null
Add-Member -InputObject $Menu.ChangeView.DropDown -MemberType NoteProperty -Name Component -Value $null
Add-Member -InputObject $Menu.ChangeView.DropDown -MemberType NoteProperty -Name Layout    -Value $null

$Menu.ChangeView.TreeView = New-Object System.Windows.Forms.ToolStripMenuItem("TreeView", $null, {

})
Add-Member -InputObject $Menu.ChangeView.TreeView -MemberType NoteProperty -Name Component -Value $null
Add-Member -InputObject $Menu.ChangeView.TreeView -MemberType NoteProperty -Name Layout    -Value $null
[void]$Menu.ChangeView.DropDown.DropDownItems.Add($Menu.ChangeView.TreeView)

$Menu.ChangeView.ListView = New-Object System.Windows.Forms.ToolStripMenuItem("ListView", $null, {

})
Add-Member -InputObject $Menu.ChangeView.ListView -MemberType NoteProperty -Name Component -Value $null
Add-Member -InputObject $Menu.ChangeView.ListView -MemberType NoteProperty -Name Layout    -Value $null
[void]$Menu.ChangeView.DropDown.DropDownItems.Add($Menu.ChangeView.ListView)

$Menu.SaveAsCsv = New-Object System.Windows.Forms.ToolStripMenuItem("CSV", $null, 
    [System.EventHandler]{
    param($sender, $e)

    $Dialog = New-Object System.Windows.Forms.SaveFileDialog
    $Dialog.ShowHelp = $false

    $data = $this.Component.Data
    foreach ($record in $data) {
        [void]$record.PSObject.Properties.Remove('Dirty')
    }

    $Dialog.Filter = "Csv File (*.csv)|*.csv"
    if($Dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
        if (Test-Path -LiteralPath $Dialog.FileName) {
            try {
                Move-Item $Dialog.FileName ("{0}.bak" -f $Dialog.FileName)
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show(
                    "Failed to create back up of existing file before saving to prevent data loss.  Please try again.",
                    "Save Device List",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
                return
            }
        }
        $data | Export-Csv $Dialog.FileName -NoTypeInformation
    }
})
$Menu.SaveAsCsv.Name = 'SaveAsCSV'
Add-Member -InputObject $Menu.SaveAsCsv -MemberType NoteProperty -Name Component -Value $null
Add-Member -InputObject $Menu.SaveAsCsv -MemberType NoteProperty -Name View      -Value $null

$Menu.SaveAs = New-Object System.Windows.Forms.ToolStripMenuItem("SaveAs", $null, @($Menu.SaveAsCsv))
$Menu.SaveAs.Name = 'SaveAs'

$Menu.Open = New-Object System.Windows.Forms.ToolStripMenuItem("Open", $null, 
    [System.EventHandler]{
    param($sender, $e)
    
    $Dialog = New-Object System.Windows.Forms.OpenFileDialog
    
    <# Fix for dialog script hang bug #>
    $Dialog.ShowHelp = $false
        
    # Dialog Configuration
    $Dialog.Filter = "DD2875 Scan Data Csv File (*.csv)|*.csv"
    $Dialog.Multiselect = $false
        
    # Run Selection Dialog
    if($($Dialog.ShowDialog()) -eq "OK") {
        Load-Data -Path $Dialog.FileName -View $this.View -Component $this.Component
    }
    else{
        return
    }
})
$Menu.Open.Name = 'Open'
Add-Member -InputObject $Menu.Open -MemberType NoteProperty -Name Component -Value $null
Add-Member -InputObject $Menu.Open -MemberType NoteProperty -Name View      -Value $null

$Menu.File = New-Object System.Windows.Forms.ToolStripMenuItem("File", $null, @($Menu.SaveAs, $Menu.Open))
$Menu.File.Name = 'File'

$Menu.Settings = New-Object System.Windows.Forms.ToolStripMenuItem("Settings", $null, {
    # Currently only launches the settings dialog window, configuration settings are
    # only used during loading.
    $Settings = & "$SettingsDialog" $Settings
})

###############################################################################
# Initialization
###############################################################################
function Initialize-Components {
    param(
        [Parameter(Mandatory = $true)]
            [System.Windows.Forms.Form]$Window,

        [Parameter(Mandatory = $true)]
            [System.Windows.Forms.TabControl]$Parent,

        [Parameter(Mandatory = $true)]
            [System.Windows.Forms.MenuStrip]$MenuStrip,

        [Parameter(Mandatory = $true)]
            [AllowEmptyCollection()]
            [System.Collections.ArrayList]$OnLoad
    )

    $Container, $Layout, $ComponentMenuStrip = New-Component
    Add-Member -InputObject $Container -Type NoteProperty -Name ComponentMenuStrip -Value $ComponentMenuStrip
    Add-Member -InputObject $Container -Type NoteProperty -Name Window             -Value $Window
    Add-Member -InputObject $Container -Type NoteProperty -Name OnLoad             -Value $OnLoad

    [void]$Parent.TabPages.Add($Container)
    [void]$MenuStrip.Items.Add($Menu.Scan)

    $Menu.Scan.Component = $Container

    $Menu.ChangeView.DropDown.Component = $Container
    $Menu.ChangeView.DropDown.Layout    = $Layout
    $Menu.ChangeView.TreeView.Component = $Container
    $Menu.ChangeView.TreeView.Layout    = $Layout
    $Menu.ChangeView.ListView.Component = $Container
    $Menu.ChangeView.ListView.Layout    = $Layout

    Load-View -Container $Container -Layout $Layout
}

function Load-View {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.Control]
            $Container,

        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.TableLayoutPanel]
            $Layout,

        [Parameter(Mandatory = $false)]
        [ValidateSet('ListView','TreeView')]
        [String]
            $View = 'ListView'
    )
    
    switch ($View) {
        ListView {
            $ViewControl = Initialize-ListComponents -Window $Container.Window -Parent $Container -MenuStrip $Container.ComponentMenuStrip -OnLoad $Container.OnLoad
        }
        TreeView {
            $ViewControl = Initialize-FlowComponents -Window $Container.Window -Parent $Container -MenuStrip $Container.ComponentMenuStrip -OnLoad $Container.OnLoad
        }
    }
    
    $Menu.Scan.View = $ViewControl
    $Menu.ChangeView.DropDown.View = $ViewControl

    [Void]$Layout.Controls.Add($ViewControl, 0, 1)
}

###############################################################################
# WinForm Constructor
###############################################################################
function New-Component() {
    # Container Definitions -------------------------------------------------------
    $Component = New-Object System.Windows.Forms.TabPage
        $Component.Dock = [System.Windows.Forms.DockStyle]::Fill
        $Component.Text = "Account Requests"

        # Attached to Parent Control by Module Component Registration Function

        # Data Source Reference for Component
        Add-Member -InputObject $Component -MemberType NoteProperty -Name Data -Value (New-Object System.Collections.ArrayList)

    $Layout = New-Object System.Windows.Forms.TableLayoutPanel
        $Layout.Dock = [System.Windows.Forms.DockStyle]::Fill
        $Layout.AutoSize = $true
        $Layout.RowCount = 2

    # ToolStrip Section
    [Void]$Layout.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle) )
        $Layout.RowStyles[0].SizeType = [System.Windows.Forms.SizeType]::Absolute
        $Layout.RowStyles[0].Height = 30

    # DataView Section
    [Void]$Layout.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle) )
        $Layout.RowStyles[1].SizeType = [System.Windows.Forms.SizeType]::Percent
        $Layout.RowStyles[1].Height = 100

    [Void]$Component.Controls.Add($Layout)

    $DeviceMenu = New-Object System.Windows.Forms.MenuStrip
    $DeviceMenu.Dock = [System.Windows.Forms.DockStyle]::Fill
        [Void]$Layout.Controls.Add($DeviceMenu, 0, 0)

    return $Component, $Layout, $DeviceMenu
}