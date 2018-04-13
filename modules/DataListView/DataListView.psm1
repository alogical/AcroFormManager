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

    # Initialize TableLayout Container with ListView display pane.
    $View = New-ListView

    $Loader = [PSCustomObject]@{
        Settings = $Settings
        View     = $View
        Parent   = $Parent
    }
    Add-Member -InputObject $Loader -MemberType ScriptMethod -Name Load -Value {
        param($sender, $e)
        if ($this.Settings) {
            if ($this.Settings.remotedb -ne [string]::Empty) {
                $this.Load_Data($this.Settings.remotedb, $this.View, $this.Parent.DataStore)
            }
        }
    }
    Add-Member -InputObject $Loader -MemberType ScriptMethod -Name Load_Data -Value ${Function:Load-Data}
    [Void]$OnLoad.Add($Loader)

    # Register Component (TableLayout Parent)
    return $View
}

###############################################################################
### Data Management
###############################################################################
function Load-Data {
    param(
        [Parameter(Mandatory = $true)]
        [String]
            $Path,

        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.Control]
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

    if ($View.Display.Items.Count -gt 0) {
        $View.Display.Items.Clear()
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
        [Parameter(Mandatory = $true)]
        [System.Array]
            $Data,

        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.Control]
            $View,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.ArrayList]
            $DataStore
    )

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

        [Array]::Sort($FieldNames)

        ## Update Fields Dropdown Menu ------------------------------------------------
        $View.Display.Fields.Clear()

        # Update Fields Filter Menu Items
        if ($Menu.Fields.HasDropDownItems) {
            $Menu.Fields.DropDownItems.Clear()
        }

        # Top Level Check All | Uncheck All
        $toggle = New-Object System.Windows.Forms.ToolStripMenuItem('Toggle All', $null, $sbFieldsToggle)
        $toggle.CheckOnClick = $true
        $toggle.Checked      = $true

        Add-Member -InputObject $toggle -MemberType NoteProperty -Name Items   -Value (New-Object System.Collections.ArrayList)
        Add-Member -InputObject $toggle -MemberType NoteProperty -Name Display -Value $View.Display

        [Void]$Menu.Fields.DropDownItems.Add($toggle)
        [Void]$Menu.Fields.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))

        $i = 1
        foreach ($field in $FieldNames) {
            if ($field -eq $PrimaryField) {continue}
            $tag = [PSCustomObject]@{
                Name    = $field
                Visible = $true
                Changed = $true
                Index   = $i++
                Column  = $null
            }
            [void]$view.Display.Fields.Add($tag)

            $item = New-Object System.Windows.Forms.ToolStripMenuItem($field, $null, $sbFieldsItem)
            $item.CheckOnClick = $true
            $item.Checked      = $true
            $item.Tag          = $tag

            # Attach ListView reference for event handlers.
            Add-Member -InputObject $item -MemberType NoteProperty -Name Display -Value $View.Display

            [Void]$Menu.Fields.DropDownItems.Add($item)
            [Void]$toggle.Items.Add($item)
        }
        
        ## Set Data Control Flags -----------------------------------------------------
        # Add state fields
        foreach ($record in $data) {
            Add-Member -InputObject $record -MemberType NoteProperty -Name Dirty -Value $false
        }

        # Saved reference to the data for later export --------------------------------
        [Void]$DataStore.Clear()
        [Void]$DataStore.AddRange($Data)

        $handle = [PSCustomObject]@{
            Control = $View.Display
            Content = $DataStore
        }

        Set-Content -ListView $handle
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

Export-ModuleMember *

###############################################################################
###############################################################################
## SECTION 02 ## PRIVATE FUNCTIONS AND VARIABLES
##
## No function or variable in this section is exported unless done so by an
## explicit call to Export-ModuleMember
###############################################################################
###############################################################################

Import-Module -Name "$Global:AppPath\modules\DataListView\CsExtensions.ps1" -Force

$Script:ResultsPaneSortColumn = -1
$PrimaryField = 'Name'

###############################################################################
### ListView Content Constructor
###############################################################################
Function Set-Content
{
    [CmdletBinding()]
    Param(
        # Thead safe control handle: Syncronized[Hashtable]
        [Parameter(Position = 0, Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
            [Object]$ListView
    )

    Begin
    {

        Write-Debug "Processing $($ListView.Content.Count) items"

        # Map the CLASSES_ROOT registry hive for icon searches
        If ( !(Test-Path "HKCR:") )
        { New-PSDrive -PSProvider registry -Root HKEY_CLASSES_ROOT -Name HKCR | Out-Null }

        # Buffer of ListItems
        $Buffer = New-Object System.Collections.ArrayList
    
        ###########################################################################
        Function Add-ListItem
        {
            [CmdletBinding()]
            Param(
                [Parameter(Position = 0, Mandatory = $True)]
                [ValidateNotNullOrEmpty()]
                    [System.Windows.Forms.ListView]$Ctrl,
                [Parameter(Position = 1, Mandatory = $True)]
                [ValidateNotNullOrEmpty()]
                    [System.Windows.Forms.ListViewItem[]]$Content
                )
    
            If ( $Ctrl.InvokeRequired )
            {
                Write-Debug "Calling ListView control via Dispatcher.Invoke()"
                $Sync = @{}
                $Sync.Control = $Ctrl
                $Sync.Content = $Content
                $Handler = [System.EventHandler]{Param($Sync); $Sync.Control.Items.AddRange($Sync.Content)}
        
                $Ctrl.Invoke($Handler, ($Sync, $Null) )
            }
            Else
            {
                Write-Debug "Calling ListItem control directly"
                $Ctrl.Items.AddRange($Content)
            }
        }

        ###########################################################################
        Function Add-ImageList
        {
            [CmdletBinding()]
            Param(
                [Parameter(Position = 0, Mandatory = $True)]
                [ValidateNotNullOrEmpty()]
                    [System.Windows.Forms.ListView]$Ctrl,
                [Parameter(Position = 1, Mandatory = $True)]
                [ValidateNotNullOrEmpty()]
                    [System.Windows.Forms.ImageList]$ImageList,
                [Parameter()]
                    [Switch]$Large,
                [Parameter()]
                    [Switch]$Small
                )
    
            If ( $Ctrl.InvokeRequired )
            {
                Write-Debug "Calling ListView control via Dispatcher.Invoke()"
                $Sync = [HashTable]::Synchronized(@{})
                $Sync.Control = $Ctrl
                $Sync.ImageList = $ImageList
        
                If ( $PSBoundParameters.ContainsKey('Large') )
                {
                    $Handler = [System.EventHandler]{Param($Sync)
                        $Sync.Control.LargeImageList = $Sync.ImageList
                       }
                }
                ElseIf ( $PSBoundParameters.ContainsKey('Small') )
                {
                    $Handler = [System.EventHandler]{
                        Param($Sync)
                        $Sync.Control.SmallImageList = $Sync.ImageList
                       }
                }
                Else
                {
                    $Handler = [System.EventHandler]{
                        Param($Sync)
                        $Sync.Control.LargeImageList = $Sync.ImageList
                        $Sync.Control.SmallImageList = $Sync.ImageList
                       }
                }
        
                $Ctrl.Invoke($Handler, ($Sync, $Null) )
            }
            Else
            {
                Write-Debug "Calling ListView control directly"
                If ( $PSBoundParameters.ContainsKey('Large') )
                {
                    $Ctrl.LargeImageList = $ImageList
                }
                ElseIf ( $PSBoundParameters.ContainsKey('Small') )
                {
                    $Ctrl.SmallImageList = $ImageList
                }
                Else
                {
                    $Ctrl.LargeImageList = $ImageList
                    $Ctrl.SmallImageList = $ImageList
                }
            }
        }

        #!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        #!! Bug !! ID=#
        #!! Does not handle files that have no extension properly.
        #!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        Function Get-ExtFriendlyName
        {
            [CmdletBinding()]
            Param (
                [Parameter()]
                    [System.String]$Extension
                )

            $Extension = $Extension.Trim()
        
            If ( [System.String]::IsNullOrEmpty($Extension) )
            {
                Write-Debug "Extension is empty or null: <<<< Returning File"
                Return "File"
            }
        
            Write-Debug "Getting registry key for: $Extension"
            $Regkey = Get-Item -LiteralPath "HKCR:\$Extension" -ErrorAction SilentlyContinue
            If ( $Regkey -eq $Null )
            {
                Write-Debug "Registry class info unavailable: <<<< Returning $Extension"
                Return "$($Extension.Remove(0, 1).ToUpper()) File"
            }

            $szCLASS = $Regkey.GetValue("")
            $Regkey = Get-Item -LiteralPath "HKCR:\$szCLASS" -ErrorAction SilentlyContinue
            If ( $Regkey -eq $Null )
            {
                Write-Debug "Registry class info unavailable: <<<< Returning $Extension"
                Return "$($Extension.Remove(0, 1).ToUpper()) File"
            }

            $FriendlyName = $Regkey.GetValue("")
            If ( ($FriendlyName -eq $Null) -or
                 ($FriendlyName -eq "") )
            {
                Write-Debug "Registry class info unavailable: <<<< Returning $Extension"
                Return "$(("$($Extension.Remove(0, 1).ToUpper()) File").Trim())" 
            }

            Return $FriendlyName
        }

    
    } # End Begin Block

    Process
    {
        # Initialize Column Headers -------------------------------------------
        if ($ListView.Control.Columns.Count -gt 0) {
            $ListView.Control.Columns.Clear()
        }
        $widths = Format-ColumnWidth $ListView.Control
        Set-ColumnHeader -Ctrl $ListView.Control -Text $PrimaryField -Width $widths[$PrimaryField] -Align Left | Out-Null
        ForEach ($field in $ListView.Control.Fields) {
            $field.Column  = Set-ColumnHeader -Ctrl $ListView.Control -Text $field.Name -Width $widths[$field.Name] -Align Left
            $field.Changed = $false
        }

        # Initialize List Items -----------------------------------------------
        ForEach ($Item in $ListView.Content) {
            $ListItem = New-Object System.Windows.Forms.ListViewItem
            $ListItem.Text = $Item.($PrimaryField)
            $ListItem.Name = $Item.Path
            $ListItem.Tag  = $Item

            Add-Member -InputObject $ListItem -MemberType NoteProperty -Name Extension -Value (".{0}" -f ($Item.Name.Split("."))[-1])
        
            ### Set Detail Info -----------------------------------------------
            foreach ($field in $ListView.Control.Fields) {
                if ($field.Name -eq $PrimaryField) {
                    continue
                }
                [void]$ListItem.SubItems.Add($Item.($field.Name))
            }

            ### Ectract File Icon ---------------------------------------------
            If ( [System.String]::IsNullOrEmpty($ListItem.Extension) )
            {
                Write-Debug "Extension is a null or empty string: Default file image"
                $ListItem.ImageIndex = 1
                [void]$Buffer.Add($ListItem)
                Continue
            }
            
            If (-Not ($ImageList.Extensions.ContainsKey($ListItem.Extension)))
            {
            
                #******************************************************************
                #** Attempt to extract the icon from file w/.NET based on windows
                #** registry information. For more information see:
                #** http://www.brad-smith.info/blog/archives/164
                #******************************************************************
                Write-Debug "Extracting icon for: $($Item.FullName)
                Discovering default icon source from HKEY_CLASSES_ROOT"
            
                Try
                {
                    Write-Debug "Getting registry key for:
                        File = $($Item.Name)
                        Extension = $(If ($ListItem.Extension.Length -gt 0){$ListItem.Extension}Else{'EmptyString'})"
                    
                    $RegKey = Get-Item -LiteralPath "HKCR:\$($ListItem.Extension)" -ErrorAction Stop
                    $szCLASS = $RegKey.GetValue("")
                
                    Write-Debug "HKEY_CLASSES_ROOT contained the following:
                    RegKey  = $RegKey
                    szCLASS = $szCLASS"

                    Write-Debug "Checking $szCLASS for DefaultIcon"
                    $szTemp = $szCLASS
                    $RegKey = Get-Item -LiteralPath "HKCR:\$szCLASS" -ErrorAction Stop
                    $SubKeys = $RegKey.GetSubKeyNames()
                    Write-Debug "Testing sub-keys of: $RegKey"
                    ForEach ( $SubKey in $SubKeys )
                    {
                        If ($SubKey -eq 'DefaultIcon')
                        {
                            Write-Debug "DefaultIcon key found"
                            $szCLASS = $szTemp
                            Break
                        }
                        If ( $SubKey -eq 'CurVer' )
                        {
                            Write-Debug "CurVer key found"
                            $RegKey = Get-Item -LiteralPath "HKCR:\$szTemp\$SubKey" -ErrorAction Stop
                            $szCLASS = $RegKey.GetValue("")
                        }
                    }
                
                    Write-Debug "Pulling icon source path from:`nHKCR:\$szCLASS\DefaultIcon"
                
                    $RegKey = Get-Item -LiteralPath "HKCR:\$szCLASS\DefaultIcon" -ErrorAction Stop
                    $szInfo = $($RegKey.GetValue("")).Split(",")
                
                    Write-Debug "Icon source path: $($szInfo[0]), $($szInfo[1])"
                
                    $lIcon = [System.IconExtractor]::Extract($szInfo[0], [int]$szInfo[1], $True)
                    $sIcon = [System.IconExtractor]::Extract($szInfo[0], [int]$szInfo[1], $False)
                }
                Catch
                {
                    Write-Warning "DefaultIcon could not be found in windows registry."
                    $lIcon = $Null
                    $sIcon = $Null
                }
            
                #!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                #!! TODO !! ID=0004
                #!! Simple logic to handles unusual icon cases where the default
                #!! icon for a file cannot be found. In this circumstance the
                #!! extension is simply associated with the icon for system file
                #!! types.
                #!!
                #!! This solution is not perfect, as the system may infact have the
                #!! application installed that handles this file type, but it stores
                #!! the DefaultIcon information in a manner that is a-typical and
                #!! not handled gracefully by the code above.
                #!!
                #!! A better solution would be to provide some kind of item tagging
                #!! for extensions with an unknown icon/application and provide a
                #!! select program dialog when trying to open the file. The select
                #!! program dialog may already be provide by start-process for
                #!! unknown file extensions, but that needs to be verified.
                #!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                If ( ($lIcon -eq $Null) -or
                     ($sIcon -eq $Null) )
                {
                    Write-Warning "Icon extraction failed for $($Item.Name)"
                    $ImageList.Extensions.Add($ListItem.Extension, 1)
                }
                Else
                {
                    Write-Debug "Adding Icon to list for extension: $($ListItem.Extension)"
                    $ImageList.Large.Images.Add($lIcon)
                    $ImageList.Small.Images.Add($sIcon)
                    $ImageList.Extensions.Add($ListItem.Extension, $ImageList.Index)
                    $ImageList.Index++
                }
            }
            $ListItem.ImageIndex = $ImageList.Extensions.Get_Item($ListItem.Extension)
        
            [void]$Buffer.Add($ListItem)
        }
    
    } # End Process Block

    End
    {
        Begin-Update $ListView.Control
        Add-ListItem $ListView.Control $Buffer.ToArray()
        Add-ImageList $ListView.Control $ImageList.Large -Large
        Add-ImageList $ListView.Control $ImageList.Small -Small
        End-Update $ListView.Control
    } # End Block
} # End Set-Content

function Format-ColumnWidth ($ListView) {
    $widths  = @{}
    $padding = 10
    if ($ListView.Columns.Count -gt 0) {
        foreach ($column in $ListView.Columns) {
            $widths.Add($column.Text, $column.Width)
        }
    }
    $g = $ListView.CreateGraphics()
    foreach ($field in $ListView.Fields) {
        if (!$widths.ContainsKey($field.Name)) {
            $widths.Add($field.Name, [Int]($g.MeasureString($field.Name, $ListView.Font).Width) + $padding)
        }
    }
    if (!$widths.ContainsKey($PrimaryField)) {
        $widths.Add($PrimaryField, [Int]($g.MeasureString($PrimaryField, $ListView.Font).Width) + $padding)
    }
    $g.Dispose()

    return $widths
}

#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
#!! TODO !! ID=0005
#!! Add functionality to display the number of selected items in the Status
#!! Label
#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

###############################################################################
### Control Event Handlers
###############################################################################
### ListView DoubleClick Event ------------------------------------------------
Function DoubleClick-ListView
{
    [CmdletBinding()]
    Param($sender, $e)

    ForEach($Item in $sender.SelectedItems) {
        # Escape PowerShell Path pattern matching characters.
        $Path = $Item.Tag.Path.Replace(']','`]')
        $Path = $Path.Replace('[','`[')
        Start-Process -FilePath $Path
    }
}

### ListView ViewSelect Menu Click Event ---------------------------------------
Function Click-ViewSelect
{
    Switch ($this.Text) {
        "Details" {
            $this.ListView.View = [System.Windows.Forms.View]::Details
            $this.FieldsMenu.Visible = $true
            }
        "List" {
            $this.ListView.View = [System.Windows.Forms.View]::List
            $this.FieldsMenu.Visible = $false
            }
        "Large Icon" {
            $this.ListView.View = [System.Windows.Forms.View]::LargeIcon
            $this.FieldsMenu.Visible = $false
            }
        "Small Icon" {
            $this.ListView.View = [System.Windows.Forms.View]::SmallIcon
            $this.FieldsMenu.Visible = $false
            }
        Default {
            $this.ListView.View = [System.Windows.Forms.View]::List
            $this.FieldsMenu.Visible = $false
            }
        }
}

Function Click-ListViewColumn
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $True)]
            [Object]$sender,
        [Parameter(Mandatory = $True)]
            [System.EventArgs]$e
        )
    
    # http://msdn.microsoft.com/en-us/library/ms996467.aspx
    
    Write-Debug "Sender allows direct access <Thread Safe> $(-Not($sender.InvokeRequired))
        Source ...... = $sender
        EventArgs ... = $e
        Sort Column . = $($e.Column)
        Sorted Column = $Script:ResultsPaneSortColumn"
    
    If ( $e.Column -ne $Script:ResultsPaneSortColumn )
    {
        $Script:ResultsPaneSortColumn = $e.Column
        $sender.Sorting = [System.Windows.Forms.SortOrder]::Ascending
        Write-Debug "First Sort :: Sort Ascending"
    }
    Else
    {
        If ( $sender.Sorting -eq [System.Windows.Forms.SortOrder]::Ascending )
        {
            Write-Debug "Sort Descending"
            $sender.Sorting = [System.Windows.Forms.SortOrder]::Descending
        }
        Else
        {
            Write-Debug "Sort Ascending"
            $sender.Sorting = [System.Windows.Forms.SortOrder]::Ascending
        }
    }
    
    $sender.Sort()
    $sender.ListViewItemSorter = New-Object ListViewSorter.ItemComparer($e.Column, $sender.Sorting)
}

###############################################################################
### ListView Custom Methods
###############################################################################
$sbRefreshFields = {
    $widths = Format-ColumnWidth $this

    foreach ($field in $this.Fields) {
        if (!$field.Changed) {
            continue
        }

        $field.Changed = $false
        if ($field.Visible) {
            [void]$this.Columns.Insert($field.Index, $field.Column)
        }
        else {
            [void]$this.Columns.Remove($field.Column)
        }
        
    }
}

###############################################################################
### ListView Menu Handlers
###############################################################################
$sbFieldsToggle = {
    foreach ($item in $this.Items) {
        $item.Checked = $this.Checked
        $item.Tag.Visible = $this.Checked
        $item.Tag.Changed = $true
    }

    Suspend-Layout $this.Display
    $this.Display.RefreshFields()
    Resume-Layout  $this.Display
}

$sbFieldsItem = {
    $this.Tag.Visible = $this.Checked
    $this.Tag.Changed = $true

    Suspend-Layout $this.Display
    $this.Display.RefreshFields()
    Resume-Layout  $this.Display
}

###############################################################################
### ListView Method Wrappers
###############################################################################
Function Begin-Update
{
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
            [System.Windows.Forms.ListView]$Ctrl
        )
    
    If ( $Ctrl.InvokeRequired )
    {
        Write-Debug "Calling ListView control via Dispatcher.Invoke()"
        $Handler = [System.EventHandler]{Param($Sync); $Sync.BeginUpdate()}
        $Ctrl.Invoke($Handler, ($Ctrl, $Null) )
    }
    Else
    {
        Write-Debug "Calling ListView control directly"
        $Ctrl.BeginUpdate()
    }
}

Function End-Update
{
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
            [System.Windows.Forms.ListView]$Ctrl
        )
    
    If ( $Ctrl.InvokeRequired )
    {
        Write-Debug "Calling ListView control via Dispatcher.Invoke()"
        $Handler = [System.EventHandler]{Param($Sync); $Sync.EndUpdate()}
        $Ctrl.Invoke($Handler, ($Ctrl, $Null) )
    }
    Else
    {
        Write-Debug "Calling ListView control directly"
        $Ctrl.EndUpdate()
    }
}

Function Suspend-Layout
{
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
            [System.Windows.Forms.ListView]$Ctrl
        )
    
    If ( $Ctrl.InvokeRequired )
    {
        Write-Debug "Calling ListView control via Dispatcher.Invoke()"

        $Handler = [System.EventHandler]{Param($Sync); $Sync.SuspendLayout()}
        $Ctrl.Invoke($Handler, ($Ctrl, $Null) )
    }
    Else
    {
        Write-Debug "Calling ListView control directly"
        $Ctrl.SuspendLayout()
    }
}

Function Resume-Layout
{
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
            [System.Windows.Forms.ListView]$Ctrl
        )
    
    If ( $Ctrl.InvokeRequired )
    {
        Write-Debug "Calling ListView control via Dispatcher.Invoke()"
        $Handler = [System.EventHandler]{Param($Sync); $Sync.ResumeLayout()}
        $Ctrl.Invoke($Handler, ($Ctrl, $Null) )
    }
    Else
    {
        Write-Debug "Calling ListView control directly"
        $Ctrl.ResumeLayout()
    }
}

Function Set-ColumnHeader
{
    [CmdletBinding()]
    Param(
        # ListView control the header will be added to.
        [Parameter(Position = 0, Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [System.Windows.Forms.ListView]
            $Ctrl,
        
        # Dispaly text for the column header.
        [Parameter(Position = 1, Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [String]
            $Text,

        # The width of the column in pixels.
        # To adjust width to the widest item in the column -1. To autosize to heading -2.
        # Unexpected behavior can result if combined width of all columns exceeds 32,768.
        [Parameter(Position = 2, Mandatory = $False)]
        [Int]
            $Width = -2,
        
        # Header title text alignment.
        [Parameter(Position = 3, Mandatory = $False)]
        [System.Windows.Forms.HorizontalAlignment]
            $Align,

        # Column order index.
        [Parameter(Position = 4, Mandatory = $False)]
        [Int]
            $Index
        )
    
    Write-Debug "Control allows direct access <Thread Safe> $(-Not($Ctrl.InvokeRequired))"
    
    $Column = New-Object System.Windows.Forms.ColumnHeader
    $Column.Text = $Text
    $Column.Width = $Width
    If ( $PSBoundParameters.ContainsKey('Align') )
    {
        $Column.TextAlign = $Align
    }
    
    Write-Debug "Column configured:`nObject = $Column"

    if ( $Index -and $Index -lt $Ctrl.Columns.Count ) {
        $Column.DisplayIndex = $Index
    }
    
    If ( $Ctrl.InvokeRequired )
    {
        Write-Debug "Calling control via Dispatcher.Invoke()"
        $Sync = [HashTable]::Synchronized(@{})
        $Sync.Control = $Ctrl
        $Sync.Column  = $Column
        $Handler = [System.EventHandler]{Param($Sync); [void]$Sync.Control.Columns.Add($Sync.Column)}
        
        $Ctrl.Invoke($Handler, ($Sync, $Null) )
    }
    Else
    {
        Write-Debug "Calling ListItem control directly"
        [void]$Ctrl.Columns.Add($Column)
    }
    return $Column
}

###############################################################################
### Context Menus Handlers
###############################################################################
Function Click-DeleteMenu
{
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
            [System.Windows.Forms.ToolStripMenuItem]$sender,
        [Parameter(Mandatory=$True)]
            [System.EventArgs]$e
    )
    
    $Menu = $sender.GetCurrentParent()
    $ListView = $Menu.SourceControl
    $Target = $ListView.SelectedItems
    If ($Target -eq $Null) {
        Write-Debug "No items were selected to target:"
        Return
    }

    # Confirmation Dialog
    $Confirm = [System.Windows.Forms.MessageBox]::Show(
        "Confirm Deletion.`nFiles cannot be recovered after deletion.",
        "Warning",
        [System.Windows.Forms.MessageBoxButtons]::OKCancel,
        [System.Windows.Forms.MessageBoxIcon]::Warning)
    
    If ($Confirm -eq [System.Windows.Forms.DialogResult]::OK) {
        ForEach ($Item in $Target) {
            If ($Item.Tag -eq "Directory") {
                Remove-Item $Item.Name -Recurse
            } Else {
                Remove-Item $Item.Name
            }
        }
    }
}

Function Click-OpenLocationMenu
{
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [System.Windows.Forms.ToolStripMenuItem]$sender,
        [Parameter(Mandatory=$True)]
        [System.EventArgs]$e
    )
    
    If ($sender -ne $Null) {
        $Menu = $sender.GetCurrentParent()
        $ListView = $Menu.SourceControl
        $Target = $ListView.SelectedItems
        If ($Target -eq $Null) {
            Return
        }
    } Else {
        # Break out of the function early
        Return
    }
    
    # Create a collection of folders to open
    $Folders = @{}
    
    ForEach ($Item in $Target) {
        $FileInfo = Get-Item -LiteralPath $Item.Name
        If ($Item.Tag -eq "Directory")
        {
            If ( -Not($Folders.ContainsKey($FileInfo.Parent.FullName)) )
            {
                $Folders.Add($FileInfo.Parent.FullName, 'ParentDirectory')
            }
        }
        Else
        {
            If ( -Not($Folders.ContainsKey($FileInfo.DirectoryName)) )
            {
                $Folders.Add($FileInfo.DirectoryName, 'ParentDirectory')
            }
        }
    }
    
    $FolderCollection = @($Folders.Keys)
    $Response = [System.Windows.Forms.DialogResult]::OK
    Write-Debug "Trying to open $($FolderCollection.Count) folder(s)"
    If ( $FolderCollection.Count -gt 5 )
    {
        # Notify user of potential problem opening many explorer windows.
        $Response = [System.Windows.Forms.MessageBox]::Show(
            "You are about to open $($FolderCollection.Count) explorer windows.`nOpening too many windows can cause windows explorer to crash.",
            "Warning",
            [System.Windows.Forms.MessageBoxButtons]::OKCancel,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        
        Write-Debug "User responded: $Response"
    }
    
    If ( $Response -eq [System.Windows.Forms.DialogResult]::OK )
    {
        ForEach ($Folder in $FolderCollection)
        {
            Explorer.exe $Folder
        }
    }
}

###############################################################################
### Context Menus
###############################################################################
$ContextMenu = @{}

$ContextMenu.DeleteItem   = New-Object System.Windows.Forms.ToolStripMenuItem(
    "Delete",
    $Null,
    [System.EventHandler]{Click-DeleteMenu $This $_}
)
$ContextMenu.OpenLocation = New-Object System.Windows.Forms.ToolStripMenuItem(
    "Open Location",
    $Null,
    [System.EventHandler]{Click-OpenLocationMenu $This $_}
)

$ContextMenu.Menu = New-Object System.Windows.Forms.ContextMenuStrip
$ContextMenu.Menu.Items.AddRange( @($ContextMenu.DeleteItem, $ContextMenu.OpenLocation) )

###############################################################################
### Context Menus
###############################################################################
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
### Image Resources
###############################################################################
$ImageList = [PSCustomObject]@{
    Large      = New-Object System.Windows.Forms.ImageList
    Small      = New-Object System.Windows.Forms.ImageList
    Extensions = @{}
    Index      = 2
}

$ImageList.Large.ColorDepth = [System.Windows.Forms.ColorDepth]::Depth32Bit
$ImageList.Small.ColorDepth = [System.Windows.Forms.ColorDepth]::Depth32Bit

$ImageList.Large.ImageSize = New-Object System.Drawing.Size(32,32)
$ImageList.Small.ImageSize = New-Object System.Drawing.Size(16,16)

# Default Icon for Folders
[System.Drawing.Icon]$Icon = [System.IconExtractor]::Extract("Shell32.dll", 3, $True)
$ImageList.Large.Images.Add($Icon)
[System.Drawing.Icon]$Icon = [System.IconExtractor]::Extract("Shell32.dll", 3, $False)
$ImageList.Small.Images.Add($Icon)

# Default Icon for Unknown Extensions
[System.Drawing.Icon]$Icon = [System.IconExtractor]::Extract("Shell32.dll", 0, $True)
$ImageList.Large.Images.Add($Icon)
[System.Drawing.Icon]$Icon = [System.IconExtractor]::Extract("Shell32.dll", 0, $False)
$ImageList.Small.Images.Add($Icon)

###############################################################################
### Control Definition Designer
###############################################################################
function New-ListView {
    $ResultsPane = New-Object System.Windows.Forms.ListView
    $ResultsPane.Dock = [System.Windows.Forms.DockStyle]::Fill
    $ResultsPane.CheckBoxes       = $False
    $ResultsPane.Scrollable       = $True
    $ResultsPane.MultiSelect      = $True
    $ResultsPane.FullRowSelect    = $True
    $ResultsPane.LabelEdit        = $False
    $ResultsPane.GridLines        = $False
    $ResultsPane.View             = [System.Windows.Forms.View]::Details
    $ResultsPane.ContextMenuStrip = $ContextMenu.Menu

    ### ListView Event Handlers
    $ResultsPane.Add_DoubleClick({DoubleClick-ListView $this $_})
    $ResultsPane.Add_ColumnClick({Click-ListViewColumn $This $_})
    $Script:ResultsPaneSortColumn = -1

    ### ViewSelect Dropdown Selection List ----------------------------------------
    $DropDownButton = New-Object System.Windows.Forms.ToolStripDropDownButton
    $DropDown       = New-Object System.Windows.Forms.ToolStripDropDown
    $ToolStrip      = New-Object System.Windows.Forms.ToolStrip

    $DropDownButton.Text = "Select View"
    $DropDownButton.Dock = [System.Windows.Forms.DockStyle]::Left
    $ToolStrip.Dock      = [System.Windows.Forms.DockStyle]::Fill
    $ToolStrip.GripStyle = "Hidden"

    $OptionButtons = @()
    [System.EventHandler]$Handler = {Click-ViewSelect}
    $OptionButtons += New-Object System.Windows.Forms.ToolStripButton("Details",    $Null, $Handler)
    $OptionButtons += New-Object System.Windows.Forms.ToolStripButton("List",       $Null, $Handler)
    $OptionButtons += New-Object System.Windows.Forms.ToolStripButton("Large Icon", $Null, $Handler)
    $OptionButtons += New-Object System.Windows.Forms.ToolStripButton("Small Icon", $Null, $Handler)

    # Attach ListView reference used in event handler.
    foreach ($option in $OptionButtons) {
        Add-Member -InputObject $option -MemberType NoteProperty -Name ListView   -Value $ResultsPane
        Add-Member -InputObject $option -MemberType NoteProperty -Name FieldsMenu -Value $Menu.Fields
    }

    $DropDown.Items.AddRange($OptionButtons)
    $DropDownButton.DropDown = $DropDown
    [void]$ToolStrip.Items.Add($DropDownButton)
    [void]$ToolStrip.Items.Add($Menu.Fields)

    ### Control Layout Panel ------------------------------------------------------
    $TableLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $TableLayout.RowCount = 2
    $TableLayout.Dock     = [System.Windows.Forms.DockStyle]::Fill
    $TableLayout.Padding  = 0
    $TableLayout.Margin   = 0
    $TableLayout.Controls.Add($ToolStrip)
    $TableLayout.Controls.Add($ResultsPane)
    [void]$TableLayout.SetRow($ToolStrip, 0)
    [void]$TableLayout.SetRow($ResultsPane, 1)

    ### Cell Size Formatting
    $RowStyle0 = New-Object System.Windows.Forms.RowStyle
    $RowStyle0.SizeType = [System.Windows.Forms.SizeType]::Absolute
    $RowStyle0.Height   = 23

    $RowStyle1 = New-Object System.Windows.Forms.RowStyle
    $RowStyle1.SizeType = [System.Windows.Forms.SizeType]::Percent
    $RowStyle1.Height   = 100

    [void]$TableLayout.RowStyles.Add($RowStyle0)
    [void]$TableLayout.RowStyles.Add($RowStyle1)

    # Component Properties
    Add-Member -InputObject $TableLayout -MemberType NoteProperty -Name Display  -Value $ResultsPane

    # Component Methods
    Add-Member -InputObject $TableLayout -MemberType ScriptMethod -Name LoadData -Value $MethodLoadData
    Add-Member -InputObject $TableLayout -MemberType ScriptMethod -Name SetData  -Value $MethodSetData

    # Display Control Properties
    Add-Member -InputObject $ResultsPane -MemberType NoteProperty -Name Fields        -Value (New-Object System.Collections.ArrayList)
    Add-Member -InputObject $ResultsPane -MemberType NoteProperty -Name PrimaryField  -Value $PrimaryField
    Add-Member -InputObject $ResultsPane -MemberType ScriptMethod -Name RefreshFields -Value $sbRefreshFields

    ### Threadproxy----------------------------------------------------------------
    #$TSP.GUI.Controls.Add('ResultsPane', $ResultsPane)

    return $TableLayout
}