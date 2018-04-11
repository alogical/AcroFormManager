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

function Show-Report {
    param(
        [Parameter(Mandatory = $false)]
        [System.Collections.ArrayList]
            $Items
    )

    $OnLoad = New-Object System.Collections.ArrayList

    ###############################################################################
    # Window Definition
    $MainForm = New-Object System.Windows.Forms.Form
    $MainForm.Text          = "DD2875 Account Request Reporter"
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
Import-Module "$ModuleInvocationPath\SingleView.psm1" -Prefix View
Import-Module "$ModuleInvocationPath\..\DD2875\dd2875.psm1"

$ImagePath = "$ModuleInvocationPath\..\..\resources"

###############################################################################
# Static Objects and Scriptblocks


# Main Menu Definitions
#region
### File Menu -------------------------------------------------------------
$MenuItem = @{}
$MenuItem.Scan = New-Object System.Windows.Forms.ToolStripMenuItem("Scan Documents", $null, {
    param($sender, $e)
    $this.Enabled = $false
    Set-ViewData @(Scan-Folder) $this.View $this.Component
    $this.Enabled = $true
})
Add-Member -InputObject $MenuItem.Scan -MemberType NoteProperty -Name View      -Value $null
Add-Member -InputObject $MenuItem.Scan -MemberType NoteProperty -Name Component -Value $null
#endregion

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
    [void]$Parent.TabPages.Add($Container)
    [void]$MenuStrip.Items.Add($MenuItem.Scan)

    $View = Initialize-ViewComponents -Window $Window -Parent $Container -MenuStrip $ComponentMenuStrip -OnLoad $OnLoad
    $MenuItem.Scan.View      = $View
    $MenuItem.Scan.Component = $Container

    [Void]$Layout.Controls.Add($View, 0, 1)
}

function New-Component() {
    ###############################################################################
    # Container Definitions
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

    # Button Section
    [Void]$Layout.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle) )
        $Layout.RowStyles[0].SizeType = [System.Windows.Forms.SizeType]::Absolute
        $Layout.RowStyles[0].Height = 30

    # Rule Detailed Description Span Row
    [Void]$Layout.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle) )
        $Layout.RowStyles[1].SizeType = [System.Windows.Forms.SizeType]::Percent
        $Layout.RowStyles[1].Height = 100

    [Void]$Component.Controls.Add($Layout)

    $DeviceMenu = New-Object System.Windows.Forms.MenuStrip
    $DeviceMenu.Dock = [System.Windows.Forms.DockStyle]::Fill
        [Void]$Layout.Controls.Add($DeviceMenu, 0, 0)

    return $Component, $Layout, $DeviceMenu
}