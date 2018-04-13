<#
.SYNOPSIS
    Data field flow layout display control.

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

function Initialize-Components {
    param(
        [Parameter(Mandatory = $true)]
            [System.Windows.Forms.Form]$Window,

        [Parameter(Mandatory = $true)]
            [System.Windows.Forms.Control]$Parent,

        [Parameter(Mandatory = $true)]
            [System.Windows.Forms.MenuStrip]$MenuStrip,

        [Parameter(Mandatory = $true)]
            [AllowEmptyCollection()]
            [System.Collections.ArrayList]$OnLoad
    )

    # Initialize
    $View = New-ViewControl -Window $Window -DataStore $Parent.DataStore -OnLoad $OnLoad

    $Loader = [PSCustomObject]@{
        Settings = $Settings
        View     = $View
        Parent   = $Parent
    }
    Add-Member -InputObject $Loader -MemberType ScriptMethod -Name Load -Value {
        param($sender, $e)
        if ($this.Settings) {
            if ($this.Settings.remotedb -ne [string]::Empty) {
                $this.Load_Data($this.Settings.remotedb, $this.View, $this.Parent)
            }
        }
    }
    Add-Member -InputObject $Loader -MemberType ScriptMethod -Name Load_Data -Value ${Function:Load-Data}
    [Void]$OnLoad.Add($Loader)

    # Register Component (TableLayout Parent)
    return $View
}

###############################################################################
# Device Data Management Handlers
###############################################################################
function Load-Data {
    param(
        [Parameter(Mandatory = $true)]
            [String]
            $Path,

        [Parameter(Mandatory = $true)]
            [System.Windows.Forms.SplitContainer]
            $View,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
            [System.Collections.ArrayList]
            $DataStore
    )

    $Data = Import-Csv $Path

    if ($Data.Count -eq 0) {
        return
    }

    if ($View.Tree.Display.Nodes.Count -gt 0) {
        $View.Tree.Display.Nodes.Clear()
    }

    Set-Data $Data $View $DataStore
}

$MethodLoadData = {
    param(
        [Parameter(Mandatory = $true)]
            [String]
            $Path,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
            [System.Collections.ArrayList]
            $DataStore
    )

    Load-Data -Path $Path -View $this -DataStore $DataStore
}

function Set-Data {
    param(
        [Parameter(Mandatory = $false)]
            $Data,

        [Parameter(Mandatory = $true)]
            [System.Windows.Forms.SplitContainer]
            $View,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
            [System.Collections.ArrayList]
            $DataStore
    )

    if ($Data.Count -eq 0) {
        return
    }

    if ($Data.Count -gt 0) {
        Write-Debug "Processing $($Data.Count) items:
            Data Variable is [$($Data.GetType())]"

        # Filter first data object field names
        $FieldNames = @( 
            ($Data[0] |
                Get-Member -MemberType NoteProperty |
                    Select-Object -Property Name -Unique |
                        Where-Object {$_.Name -ne 'Dirty'} |
                            % {Write-Output $_.Name})
        )

        $View.Display.Fields.Clear()
        $View.Display.Fields.AddRange($FieldNames)

        # Update Fields Filter Menu Items
        if ($Menu.Fields.HasDropDownItems) {
            $Menu.Fields.DropDownItems.Clear()
        }

        # Top Level Check All | Uncheck All
        $toggle = New-Object System.Windows.Forms.ToolStripMenuItem('Toggle All', $null, {
            $this.Display.Fields.Clear()

            foreach ($item in $this.Items) {
                $item.Checked = $this.Checked
                if ($this.Checked) {
                    $this.Display.Fields.Add($item.Text)
                }
            }

            $this.Display.Redisplay()
        })
        $toggle.CheckOnClick = $true
        $toggle.Checked = $true
        Add-Member -InputObject $toggle -MemberType NoteProperty -Name Items -Value (New-Object System.Collections.ArrayList)
        Add-Member -InputObject $toggle -MemberType NoteProperty -Name Display -Value $View.Display
        [Void]$Menu.Fields.DropDownItems.Add($toggle)
        [Void]$Menu.Fields.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))

        foreach ($field in $FieldNames) {
            $item = New-Object System.Windows.Forms.ToolStripMenuItem($field, $null, {
                if ($this.Checked) {
                    if (!$this.Display.Fields.Contains($this.Text)) {
                        [void]$this.Display.Fields.Add($this.Text)
                    }
                }
                else {
                    if ($this.Display.Fields.Contains($this.Text)) {
                        [void]$this.Display.Fields.Remove($this.Text)
                    }
                }
                $this.Display.Redisplay()
            })
            $item.CheckOnClick = $true
            $item.Checked = $true
            Add-Member -InputObject $item -MemberType NoteProperty -Name Display -Value $View.Display
            [Void]$Menu.Fields.DropDownItems.Add($item)
            [Void]$toggle.Items.Add($item)
        }
        
        # Add state fields
        foreach ($record in $data) {
            Add-Member -InputObject $record -MemberType NoteProperty -Name Dirty -Value $false
        }

        # Saved reference to the data for later export
        [Void]$DataStore.Clear()
        [Void]$DataStore.AddRange($Data)

        # Set TreeView Object Data Source Fields
        $View.Tree.SettingsTab.RegisterFields($FieldNames)
        $View.Tree.FilterTab.RegisterFields($FieldNames)
    }

    if ($View.Tree.SettingsTab.Handler.Valid) {
        $View.Tree.SettingsTab.Handler.Apply()
    }
    else {
        $View.Tree.SettingsTab.PromptUser()
    }
}

$MethodSetData = {
    param(
        [Parameter(Mandatory = $true)]
            $Data,

        [Parameter(Mandatory = $true)]
            [AllowEmptyCollection()]
            [System.Collections.ArrayList]
            $DataStore
    )

    Set-Data -Data $Data -View $this -DataStore $DataStore
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
Import-Module "$Global:AppPath\modules\TreeViewExtended\TreeViewExtended.psm1" -Prefix Tree
$ImagePath  = "$Global:AppPath\resources"

### Settings Management -------------------------------------------------------
$Settings = $null
$SettingsPath   = "$InvocationPath\settings.json"
$SettingsDialog = "$InvocationPath\settings.ps1"

###############################################################################
## Load Settings
if (Test-Path -LiteralPath $SettingsPath -PathType Leaf) {
    $Settings = ConvertFrom-Json ((Get-Content $SettingsPath) -join '')
}

###############################################################################
# Main Menu Definitions
### File Menu -------------------------------------------------------------
$Menu = @{}

# Dynamic Fields Menu
$Menu.Fields = New-Object System.Windows.Forms.ToolStripMenuItem("Fields")
$Menu.Fields.Name = 'Fields'
$Menu.Fields.DropDown.Add_Closing({
    param($sender, $e)
    if ($e.CloseReason -eq [System.Windows.Forms.ToolStripDropDownCloseReason]::ItemClicked -or
        $e.CloseReason -eq [System.Windows.Forms.ToolStripDropDownCloseReason]::AppFocusChange) {
        $e.Cancel = $true
    }
})

###############################################################################
# Control Object Factories

function New-ViewControl {
    param(
        [Parameter(Mandatory = $true)]
            [System.Windows.Forms.Form]$Window,

        [Parameter(Mandatory = $true)]
            [AllowEmptyCollection()]
            [System.Collections.ArrayList]$DataStore,

        [Parameter(Mandatory = $true)]
            [AllowEmptyCollection()]
            [System.Collections.ArrayList]$OnLoad
    )

    # Component Layout
    $View = New-Object System.Windows.Forms.SplitContainer
        $View.Dock = [System.Windows.Forms.DockStyle]::Fill
        $View.Orientation = [System.Windows.Forms.Orientation]::Vertical

        # Attached to Parent Control by Module Component Registration Function
        Add-Member -InputObject $View -MemberType NoteProperty -Name FieldList -Value (New-Object System.Collections.ArrayList)
        Add-Member -InputObject $View -MemberType ScriptMethod -Name LoadData -Value $MethodLoadData
        Add-Member -InputObject $View -MemberType ScriptMethod -Name SetData -Value $MethodSetData
        

    # Device Navigation Panel
        # SortedTreeView component created by intialize function (dependecy on runtime object references)
    $TreeView = Initialize-TreeComponents `
        -Window          $Window              `
        -Parent          $View.Panel1         `
        -MenuStrip       $null                `
        -OnLoad          $OnLoad              `
        -Source          $DataStore           `
        -ImageList       $ImageList           `
        -TreeDefinition  $TreeViewDefinition  `
        -GroupDefinition $GroupNodeDefinition `
        -NodeDefinition  $DataNodeDefinition

        Add-Member -InputObject $View -MemberType NoteProperty -Name Tree -Value $TreeView

    # Device Data Layout Panel
    $DataView = New-DataLayout

        [void]$View.Panel2.Controls.Add( $DataView )
        $DataNodeDefinition.Custom.DataView = $DataView

        Add-Member -InputObject $View -MemberType NoteProperty -Name Display -Value $DataView

    return $View
}

function New-DataLayout {
    # Device Data Layout Panel
    $DataLayout = New-Object System.Windows.Forms.FlowLayoutPanel
        $DataLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
        $DataLayout.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
        $DataLayout.BackColor     = [System.Drawing.Color]::AliceBlue
        #$DataLayout.WrapContents  = $false
        $DataLayout.AutoSize      = $true
        $DataLayout.AutoSizeMode  = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
        $DataLayout.AutoScroll    = $true

    Add-Member -InputObject $DataLayout -MemberType NoteProperty -Name Fields -Value (New-Object System.Collections.ArrayList)

    Add-Member -InputObject $DataLayout -MemberType NoteProperty -Name Record -Value $null

    Add-Member -InputObject $DataLayout -MemberType ScriptMethod -Name SetContent -Value {
        param(
            [Parameter(Mandatory = $true)]
                [PSCustomObject]$record
        )

        $this.SuspendLayout()
        if ($this.Controls.Count -gt 0) {
            $this.Controls.Clear()
        }

        $this.Record = $record

        # Extract field names
        $fields =  @( 
            ($record |
                Get-Member -MemberType NoteProperty |
                    Select-Object -Property Name -Unique |
                        % {Write-Output $_.Name}))

        foreach ($field in $fields) {
            Write-Debug "Generating panel for field ($field)"
            if ($this.Fields.Contains($field)) {
                $panel = New-DataPanel -Title $field -Data $record.($field) -Record $record -MaxWidth $this.Width
            
                [Void]$this.Controls.Add($panel)
            }
        }
        $this.ResumeLayout()
    }

    Add-Member -InputObject $DataLayout -MemberType ScriptMethod -Name Redisplay -Value {
        if ($this.Record -ne $null) {
            $this.SetContent($this.Record)
        }
    }

    return $DataLayout
}

function New-DataPanel {
    param(
        [Parameter(Mandatory = $true)]
            [String]$Title,

        [Parameter(Mandatory = $true)]
            [AllowEmptyString()]
            [String]$Data,

        [Parameter(Mandatory = $true)]
            [PSCustomObject]$Record,

        [Parameter()]
            [Int]$MaxWidth
    )

    $Panel = New-Object System.Windows.Forms.Panel
        #$Panel.AutoSize = $true
        #$Panel.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
        $Panel.Height = 40
        #$Panel.Width  = $MaxWidth
        $Panel.Width = 200

    $TitleLabel = New-Object System.Windows.Forms.Label
        $TitleLabel.Text = $Title
        $TitleLabel.Dock = [System.Windows.Forms.DockStyle]::Top
        $TitleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
        #$TitleLabel.AutoSize = $true
        #$TitleLabel.Width = $MaxWidth
        $TitleLabel.Width = 200

    $DataBox = New-Object System.Windows.Forms.TextBox
        if (![String]::IsNullOrEmpty($Data)) {
            $DataBox.Text = $Data
        }
        $DataBox.Dock = [System.Windows.Forms.DockStyle]::Top
        #$DataBox.AutoSize = $true
        #$DataBox.Width = $MaxWidth
        $DataBox.Width = 200

    [Void]$Panel.Controls.Add($DataBox)
    [Void]$Panel.Controls.Add($TitleLabel)

    Add-Member -InputObject $DataBox -MemberType NoteProperty -Name Record -Value $Record
    Add-Member -InputObject $DataBox -MemberType NoteProperty -Name Field -Value $Title

    $DataBox.Add_TextChanged({
        $this.Record.($this.Field) = $this.Text
        $this.Record.Dirty = $true
    })

    return $Panel
}

###############################################################################
# TreeView Component Static Resources
$ImageList = New-Object System.Windows.Forms.ImageList
$ImageList.ColorDepth = [System.Windows.Forms.ColorDepth]::Depth32Bit
$ImageList.ImageSize  = New-Object System.Drawing.Size(16,16)
$ImageList.Images.Add('group',
    [System.Drawing.Icon]::new("$ImagePath\group.ico"))
$ImageList.Images.Add('signed',
    [System.Drawing.Icon]::new("$ImagePath\tag-blue-add.ico"))
$ImageList.Images.Add('not-signed',
    [System.Drawing.Icon]::new("$ImagePath\tag-blue-delete.ico"))

# Parameter Encapsulation Object
$TreeViewDefinition = [PSCustomObject]@{
    # [System.Windows.Forms.TreeView] Properties
    Properties = @{}

    # ScriptMethod Definitions
    Methods    = @{}

    # [System.Windows.Forms.TreeView] Event Handlers
    Handlers   = @{}
}

$TreeViewDefinition.Methods.GetChecked = {
    $checked = New-Object System.Collections.ArrayList

    if ($this.DataNodes -eq $null) {
        return $checked
    }
        
    foreach ($node in $this.DataNodes) {
        if ($node.Checked) {
            [Void] $checked.Add( $node.Tag )
        }
    }

    return $checked
}

$TreeViewDefinition.Handlers.AfterSelect = {
    param($sender, $e)

    $node = $sender.SelectedNode
    if ($node.Type -eq "Data") {
        $node.ShowDetail()
    }
}

# Parameter Encapsulation Object
$DataNodeDefinition = [PSCustomObject]@{
    # Custom NoteProperties
    Custom     = @{}

    # [System.Windows.Forms.TreeViewNode] Properties
    Properties = @{}

    # ScriptMethod Definitions
    Methods    = @{}

    # [System.Windows.Forms.TreeViewNode] Event Handlers
    Handlers   = @{}

    # SortedTreeView Module TreeNode Processing Methods. Used to customize a TreeNode during creation.
    Processors = @{}
}

# Reference for setting the data view content container
$DataNodeDefinition.Custom.DataView = $null

$DataNodeDefinition.Custom.Type = 'Data'

$DataNodeDefinition.Methods.ShowDetail = {
    $this.DataView.SetContent( $this.Tag )
}

$DataNodeDefinition.Processors.Images = {
    param($node, $record)

    # Images
    if ($record.UnitSecurityManager) {
        $node.ImageKey = "signed"
        $node.SelectedImageKey = "signed"
    }
    else {
        $node.ImageKey = "not-signed"
        $node.SelectedImageKey = "not-signed"
    }
}

$DataNodeDefinition.Properties.ContextMenuStrip = &{
    $context = New-Object System.Windows.Forms.ContextMenuStrip
    [Void]$context.Items.Add( (New-Object System.Windows.Forms.ToolStripMenuItem("Open", $null, {
        param ($sender, $e)
        $Menu = $sender.GetCurrentParent()
        [System.Windows.Forms.TreeView] $TreeView = $Menu.SourceControl
        [System.Windows.Forms.TreeNode] $Node = $TreeView.SelectedNode

        # Invoke File to Open
        Start-Process -FilePath $Node.Tag.Path
    })))
    return $context
}

# Parameter Encapsulation Object
$GroupNodeDefinition = [PSCustomObject]@{
    # Custom Properties
    Custom     = @{}

    # [System.Windows.Forms.TreeViewNode] Properties
    Properties = @{}

    # ScriptMethod Definitions
    Methods    = @{}

    # [System.Windows.Forms.TreeViewNode] Event Handlers
    Handlers   = @{}

    # SortedTreeView Module TreeNode Processing Methods. Used to customize a TreeNode during creation.
    Processors = @{}
}

$GroupNodeDefinition.Processors.Images = {
    param($node, $data)

    $node.ImageKey         = 'group'
    $node.SelectedImageKey = 'group'
}