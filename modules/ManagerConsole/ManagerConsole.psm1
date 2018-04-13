<#
.SYNOPSIS
    DD2875 Report Viewer.

.DESCRIPTION

.NOTES
    Author: Daniel K. Ives
    Email:  daniel.ives@live.com
#>

Add-Type -AssemblyName System.Windows.Forms

$InvocationPath  = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Definition)

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
    # Register Components
    $Console = Initialize-Components $MainForm $Menu $OnLoad
        [void]$MainForm.Controls.Add($Console)

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
Import-Module "$Global:AppPath\modules\DataTreeView\DataTreeView.psm1" -Prefix Tree
Import-Module "$Global:AppPath\modules\DataListView\DataListView.psm1" -Prefix List
Import-Module "$Global:AppPath\modules\DD2875\dd2875.psm1"
$ImagePath  = "$Global:AppPath\resources"

###############################################################################
# Static Objects and Scriptblocks
###############################################################################

###############################################################################
# Main Menu Definitions
###############################################################################
### File Menu -------------------------------------------------------------
$MenuItem = @{}
$MenuItem.Scan = New-Object System.Windows.Forms.ToolStripMenuItem("Scan Documents", $null, {
    param($sender, $e)
    $this.Enabled = $false
    $this.Layout.View.SetData( @(Scan-Folder), $this.DataStore)
    $this.Enabled = $true
})
Add-Member -InputObject $MenuItem.Scan -MemberType NoteProperty -Name Layout    -Value $null
Add-Member -InputObject $MenuItem.Scan -MemberType NoteProperty -Name DataStore -Value $null

$MenuItem.ChangeView = @{}
$MenuItem.ChangeView.DropDown = New-Object System.Windows.Forms.ToolStripMenuItem("View")

$MenuItem.ChangeView.TreeView = New-Object System.Windows.Forms.ToolStripMenuItem("TreeView", $null, {
    if ($this.Displayed) {
        return
    }
    Load-View -Container $this.Layout -View TreeView
    if ($this.DataStore.Count -gt 0) {
        $this.Layout.View.SetData($this.DataStore.ToArray(), $this.DataStore)
    }
})
Add-Member -InputObject $MenuItem.ChangeView.TreeView -MemberType NoteProperty -Name DataStore -Value $null
Add-Member -InputObject $MenuItem.ChangeView.TreeView -MemberType NoteProperty -Name Layout    -Value $null
Add-Member -InputObject $MenuItem.ChangeView.TreeView -MemberType NoteProperty -Name Displayed  -Value $false
[void]$MenuItem.ChangeView.DropDown.DropDownItems.Add($MenuItem.ChangeView.TreeView)

$MenuItem.ChangeView.ListView = New-Object System.Windows.Forms.ToolStripMenuItem("ListView", $null, {
    if ($this.Displayed) {
        return
    }
    Load-View -Container $this.Layout -View ListView
    if ($this.DataStore.Count -gt 0) {
        $this.Layout.View.SetData($this.DataStore.ToArray(), $this.DataStore)
    }
})
Add-Member -InputObject $MenuItem.ChangeView.ListView -MemberType NoteProperty -Name DataStore -Value $null
Add-Member -InputObject $MenuItem.ChangeView.ListView -MemberType NoteProperty -Name Layout    -Value $null
Add-Member -InputObject $MenuItem.ChangeView.ListView -MemberType NoteProperty -Name Displayed -Value $false
[void]$MenuItem.ChangeView.DropDown.DropDownItems.Add($MenuItem.ChangeView.ListView)

$MenuItem.SaveAsCsv = New-Object System.Windows.Forms.ToolStripMenuItem("CSV", $null, 
    [System.EventHandler]{
    param($sender, $e)

    $Dialog = New-Object System.Windows.Forms.SaveFileDialog
    $Dialog.ShowHelp = $false

    $data = $this.DataStore
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
        $DataStore | Export-Csv $Dialog.FileName -NoTypeInformation
    }
})
$MenuItem.SaveAsCsv.Name = 'SaveAsCSV'
Add-Member -InputObject $MenuItem.SaveAsCsv -MemberType NoteProperty -Name DataStore -Value $null
Add-Member -InputObject $MenuItem.SaveAsCsv -MemberType NoteProperty -Name View      -Value $null

$MenuItem.SaveAs = New-Object System.Windows.Forms.ToolStripMenuItem("SaveAs", $null, @($MenuItem.SaveAsCsv))
$MenuItem.SaveAs.Name = 'SaveAs'

$MenuItem.Open = New-Object System.Windows.Forms.ToolStripMenuItem("Open", $null, 
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
        $this.Layout.View.LoadData($Dialog.FileName, $this.DataStore)
    }
    else{
        return
    }
})
$MenuItem.Open.Name = 'Open'
Add-Member -InputObject $MenuItem.Open -MemberType NoteProperty -Name DataStore -Value $null
Add-Member -InputObject $MenuItem.Open -MemberType NoteProperty -Name Layout    -Value $null

$MenuItem.File = New-Object System.Windows.Forms.ToolStripMenuItem("File", $null, @($MenuItem.SaveAs, $MenuItem.Open))
$MenuItem.File.Name = 'File'

$MenuItem.Settings = New-Object System.Windows.Forms.ToolStripMenuItem("Settings", $null, {
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
            [System.Windows.Forms.MenuStrip]$MenuStrip,

        [Parameter(Mandatory = $true)]
            [AllowEmptyCollection()]
            [System.Collections.ArrayList]$OnLoad
    )

    $Container = New-Component
    $Container.Window = $Window
    $Container.OnLoad = $OnLoad

    [void]$MenuStrip.Items.Add($MenuItem.File)
    [void]$MenuStrip.Items.Add($MenuItem.Scan)
    [void]$MenuStrip.Items.Add($MenuItem.ChangeView.DropDown)
    [void]$MenuStrip.Items.Add($MenuItem.Settings)


    $MenuItem.Scan.DataStore = $Container.DataStore
    $MenuItem.Scan.Layout    = $Container
    $MenuItem.Open.DataStore = $Container.DataStore
    $MenuItem.Open.Layout    = $Container
    $MenuItem.ChangeView.TreeView.DataStore = $Container.DataStore
    $MenuItem.ChangeView.TreeView.Layout    = $Container
    $MenuItem.ChangeView.ListView.DataStore = $Container.DataStore
    $MenuItem.ChangeView.ListView.Layout    = $Container

    Load-View -Container $Container -View TreeView
    
    return $Container
}

function Load-View {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.TableLayoutPanel]
            $Container,

        [Parameter(Mandatory = $false)]
        [ValidateSet('ListView','TreeView')]
        [String]
            $View = 'ListView'
    )
    
    if ($Container.View) {
        [void]$Container.Controls.Remove( ($Container.GetControlFromPosition(0,1)) )
    }

    switch ($View) {
        ListView {
            $ViewControl = Initialize-ListComponents -Window $Container.Window -Parent $Container -MenuStrip $Container.ComponentMenuStrip -OnLoad $Container.OnLoad
            $MenuItem.ChangeView.ListView.Displayed = $true
            $MenuItem.ChangeView.TreeView.Displayed = $false
        }
        TreeView {
            $ViewControl = Initialize-TreeComponents -Window $Container.Window -Parent $Container -MenuStrip $Container.ComponentMenuStrip -OnLoad $Container.OnLoad
            $MenuItem.ChangeView.ListView.Displayed = $false
            $MenuItem.ChangeView.TreeView.Displayed = $true
        }
    }
    
    $replaced       = $Container.View
    $Container.View = $ViewControl

    [Void]$Container.Controls.Add($ViewControl, 0, 1)

    if ($replaced) {
        $replaced.Dispose()
    }
}

###############################################################################
# WinForm Constructor
###############################################################################
function New-Component() {
    $Layout = New-Object System.Windows.Forms.TableLayoutPanel
        $Layout.Dock = [System.Windows.Forms.DockStyle]::Fill
        $Layout.AutoSize = $true
        $Layout.RowCount = 2

        # Reference to ArrayList (Must be directly accessable!)
        Add-Member -InputObject $Layout -MemberType NoteProperty -Name OnLoad    -Value $null
        Add-Member -InputObject $Layout -MemberType NoteProperty -Name DataStore -Value (New-Object System.Collections.ArrayList)
        
        ## Private Storage ------------------------------------------------------------
        # Data Source Reference for Component
        Add-Member -InputObject $Layout -MemberType NoteProperty -Name __ComponentMenuStrip -Value $null
        Add-Member -InputObject $Layout -MemberType NoteProperty -Name __Window             -Value $null
        Add-Member -InputObject $Layout -MemberType NoteProperty -Name __View               -Value $null

        # Accessors
        Add-Member -InputObject $Layout -Type ScriptProperty -Name ComponentMenuStrip -Value {
            # Get Property
            return $this.__ComponentMenuStrip
            }{
            # Set Property
            param(
                [Parameter(Mandatory = $true)]
                [System.Windows.Forms.MenuStrip]
                    $InputObject
                )
            $this.__ComponentMenuStrip = $InputObject
        }
        Add-Member -InputObject $Layout -Type ScriptProperty -Name Window -Value {
            # Get Property
            return $this.__Window
            }{
            # Set Property
            param(
                [Parameter(Mandatory = $true)]
                [System.Windows.Forms.Form]
                    $InputObject
                )
            $this.__Window = $InputObject
        }
        Add-Member -InputObject $Layout -Type ScriptProperty -Name View -Value {
            # Get Property
            return $this.__View
            }{
            # Set Property
            param(
                [Parameter(Mandatory = $true)]
                [System.Windows.Forms.Control]
                    $InputObject
                )
            $this.__View = $InputObject
        }

    # ToolStrip Section
    [Void]$Layout.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle) )
        $Layout.RowStyles[0].SizeType = [System.Windows.Forms.SizeType]::Absolute
        $Layout.RowStyles[0].Height = 30

    # DataView Section
    [Void]$Layout.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle) )
        $Layout.RowStyles[1].SizeType = [System.Windows.Forms.SizeType]::Percent
        $Layout.RowStyles[1].Height = 100

    $DeviceMenu = New-Object System.Windows.Forms.MenuStrip
    $DeviceMenu.Dock = [System.Windows.Forms.DockStyle]::Fill
        [Void]$Layout.Controls.Add($DeviceMenu, 0, 0)
        $Layout.ComponentMenuStrip = $DeviceMenu

    return $Layout
}