<#
.SYNOPSIS
    Custom TreeView control with display and filter settings.

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

function Initialize-Components {
    param(
        [Parameter(Mandatory = $true)]
            [AllowNull()]
            [System.Windows.Forms.Form]
            $Window,

        [Parameter(Mandatory = $true)]
            [System.Windows.Forms.Control]
            $Parent,

        [Parameter(Mandatory = $true)]
            [AllowNull()]
            [System.Windows.Forms.MenuStrip]
            $MenuStrip,

        [Parameter(Mandatory = $true)]
            [AllowEmptyCollection()]
            [System.Collections.ArrayList]
            $OnLoad,

        [Parameter(Mandatory = $true)]
            [AllowEmptyCollection()]
            [System.Collections.ArrayList]
            $Source,

        [Parameter()]
            [System.Windows.Forms.ImageList]
            $ImageList,

        [Parameter()]
            [PSCustomObject]
            $TreeDefinition,

        [Parameter()]
            [PSCustomObject]
            $GroupDefinition,

        [Parameter()]
            [PSCustomObject]
            $NodeDefinition
    )

    $Control = New-SortedTreeView $Window $Parent $Source $ImageList $TreeDefinition $GroupDefinition $NodeDefinition
    Add-Member -InputObject $Control -MemberType NoteProperty -Name Static -Value $Static

    # Register Control With Parent
    [void]$Parent.Controls.Add($Control)

    return $Control
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


###############################################################################
# Static Objects and Scriptblocks used for the creation of an indeterminate num
# of objects as components of those objects.
$Static = @{}
$Static.ProcessDefinition = {
    param([System.Windows.Forms.TreeNode]$node, $record, [Object]$definition)
    # CUSTOM PRE-PROCESSORS
    foreach ($processor in $definition.Processors.Values) {
        & $processor $node $record | Out-Null
    }

    # CUSTOM PROPERTIES
    foreach ($property in $definition.Custom.GetEnumerator()) {
        Add-Member -InputObject $node -MemberType NoteProperty -Name $property.Key -Value $property.Value
    }

    # STANDARD PROPERTIES
    foreach ($property in $definition.Properties.GetEnumerator()) {
        $node.($property.Key) = $property.Value
    }

    # CUSTOM METHODS
    foreach ($method in $definition.Methods.GetEnumerator()) {
        Add-Member -InputObject $node -MemberType ScriptMethod -Name $method.Key -Value $method.Value
    }

    # EVENT HANDLERS - EVENT ADD METHOD CALLED BY LATE BINDING
    foreach ($handler in $definition.Handlers.GetEnumerator()) {
        $node.("Add_$($handler.Key)")($handler.Value)
    }

    Add-Member -InputObject $node -MemberType NoteProperty -Name Filtered -Value $null
    Add-Member -InputObject $node -MemberType ScriptMethod -Name Filter   -Value $Static.Filter
}
$Static.TreeNodeChecked = {
    param($state)

    $this.Checked = $state
    ForEach ($child in $this.Nodes) {
        $child.SetChecked($State)
    }
}
$Static.NewDataBucket = {
    param($Field, [System.Collections.ArrayList]$Buffer, [Object]$Definition)

    Write-Debug "Creating new data bucket"

    $bucket = [PSCustomObject]@{
        Field   = $Field
        Content = New-Object System.Collections.ArrayList
        Buffer  = $Buffer
        Definition = $Definition
    }

    Add-Member -InputObject $bucket -MemberType ScriptMethod -Name Add -Value {
        param($record)
        [void] $this.Content.Add($record)
    }

    Add-Member -InputObject $bucket -MemberType ScriptMethod -Name Build -Value {
        # Process Node Definition for Each Data Record
        ForEach ($record in $this.Content) {
            $node = New-Object System.Windows.Forms.TreeNode($record.($this.Field))
            $node.Tag = $record

            # DEFAULT STATIC METHODS
            Add-Member -InputObject $node -MemberType ScriptMethod -Name SetChecked -Value $Static.TreeNodeChecked

            # CUSTOM PRE-PROCESSORS
            foreach ($processor in $this.Definition.Processors.Values) {
                & $processor $node $record | Out-Null
            }

            # CUSTOM PROPERTIES
            foreach ($property in $this.Definition.Custom.GetEnumerator()) {
                Add-Member -InputObject $node -MemberType NoteProperty -Name $property.Key -Value $property.Value
            }

            # STANDARD PROPERTIES
            foreach ($property in $this.Definition.Properties.GetEnumerator()) {
                $node.($property.Key) = $property.Value
            }

            # CUSTOM METHODS
            foreach ($method in $this.Definition.Methods.GetEnumerator()) {
                Add-Member -InputObject $node -MemberType ScriptMethod -Name $method.Key -Value $method.Value
            }

            Add-Member -InputObject $node -MemberType NoteProperty -Name Filtered -Value $null
            Add-Member -InputObject $node -MemberType ScriptMethod -Name Filter   -Value $Static.Filter

            # Append to Quick Access ArrayList
            [void]$this.Buffer.Add($node)

            Write-Output $node
        }
    }

    return $bucket
}
$Static.NewDataBucketFactory = {
    param($DataLabel, [System.Collections.ArrayList]$Buffer, [Object]$Definition)

    Write-Debug "Creating a new data bucket factory"

    $Factory = [PSCustomObject]@{
        DataLabel  = $DataLabel
        Buffer     = $Buffer
        Definition = $Definition
    }

    Add-Member -InputObject $Factory -MemberType ScriptMethod -Name New -Value {
        Write-Debug "Calling New Data Bucket"
        return & $Static.NewDataBucket $this.DataLabel $this.Buffer $this.Definition
    }

    return $Factory
}
$Static.NewGroupBucket = {
    param($Rule, $Factory, [Object]$Definition)
                
    Write-Debug ("Factory Create [{0}] Bucket" -f $Rule.Name)

    $bucket = [PSCustomObject]@{
        Rule    = $Rule
        Factory = $Factory
        Content = @{}
        Definition = $Definition
    }

    Add-Member -InputObject $bucket -MemberType ScriptMethod -Name Add -Value {
        param($record)

        $value = $record.($this.Rule.Name)

        if (!$this.Content.ContainsKey($value)) {
                    
            Write-Debug "Creating Bucket For ($value)"

            $this.Content.Add($value, $this.Factory.New())
        }

        $this.Content[$value].Add($record)
    }

    Add-Member -InputObject $bucket -MemberType ScriptMethod -Name Build -Value {
        
        $nodes = New-Object System.Collections.ArrayList
        
        # Sort Grouping Nodes
        if ($this.Rule.SortDirection -eq 'Ascending') {
            $Sorted = $this.Content.GetEnumerator() | Sort-Object -Property Key
        }
        else {
            $Sorted = $this.Content.GetEnumerator() | Sort-Object -Property Key -Descending
        }

        # Build the Child Nodes in Sorted Order (Hashtable Key[GroupName]:Value[ChildBucket] Pairs)
        ForEach ($pair in $Sorted) {
            $node = New-Object System.Windows.Forms.TreeNode($pair.Key)

            # CUSTOMIZE OBJECT
            & $Static.ProcessDefinition $node $pair $this.Definition

            # DEFAULT STATIC METHODS
            Add-Member -InputObject $node -MemberType ScriptMethod -Name SetChecked -Value $Static.TreeNodeChecked

            $node.Nodes.AddRange( $pair.Value.Build() )
            [void] $nodes.Add($node)
        }

        return $nodes
    }

    return $bucket
}
$Static.NewGroupFactory = {
    param($Rule, $Next, [Object]$Definition)

    Write-Debug ("Creating new bucket factory [{0}]" -f $Rule.Name)


    $Factory = [PSCustomObject]@{
        Rule = $Rule
        Next = $Next
        Definition = $Definition
    }

    Add-Member -InputObject $Factory -MemberType ScriptMethod -Name New -Value {
        return (& $Static.NewGroupBucket $this.Rule $this.Next $this.Definition)
    }

    return $Factory
}
$Static.NewGroupConstructor = {
    param($GroupBy, $DataLabel, [System.Collections.ArrayList]$Buffer, [Object]$GroupDefinition, [Object]$NodeDefinition)

    Write-Debug ("Creating new TreeView Group Constructor {0}" -f ($GroupBy | Select-Object -Property Name | Out-String))
    Write-Debug "Data Labels Using field [$DataLabel]"

    $FactoryList = New-Object System.Collections.ArrayList

    # Build Initial Grouping Level Buckets
    if ($GroupBy.Count -gt 0) {
        $max = $GroupBy.Count - 1

        for ($i = 0; $i -le $max; $i++) {
                
            # MUST ONLY BE ONE TOP BUCKET
            if ($i -eq 0) {
                $TopBucket = & $Static.NewGroupBucket $GroupBy[$i] $null $GroupDefinition
            }
            else {
                [void] $FactoryList.Add( (& $Static.NewGroupFactory $GroupBy[$i] $null $GroupDefinition) )
            }
        }

        # Back Reference Factories
        $max = $FactoryList.Count - 1
        for ($i = 1; $i -le $max; $i++) {
            $FactoryList[$i - 1].Next = $FactoryList[$i]
        }
    }

    # Handle View Settings with no Grouping Levels Defined
    else {
        Write-Debug "No Grouping Defined, Top Bucket is the Data Bucket"
        $TopBucket = & $Static.NewDataBucket $DataLabel $Buffer $NodeDefinition
    }
            
    # Handle Back References for Multiple Grouping Level Factories
    if ($GroupBy.Count -gt 1) {
        Write-Debug ("Data Bucket Factory Assigned to [{0}] Factory" -f $FactoryList[$max].Rule.Name)
        $FactoryList[$max].Next = & $Static.NewDataBucketFactory $DataLabel $Buffer $NodeDefinition

        # Top Bucket gets the next level bucket factory
        $TopBucket.Factory = $FactoryList[0]
    }

    # Handle Single Grouping Level
    elseif ($GroupBy.Count -eq 1) {
        # If there is only one grouping level then TopBucket gets data buckets
        $TopBucket.Factory = & $Static.NewDataBucketFactory $DataLabel $Buffer $NodeDefinition
    }

    # Build the Constructor Object
    $Constructor = [PSCustomObject]@{
        FactoryList = $FactoryList
        Tree        = $TopBucket
    }

    Add-Member -InputObject $Constructor -MemberType ScriptMethod -Name AddRange -Value {
        param([System.Collections.ArrayList]$Range)
        $tree = $this.Tree
        ForEach ($record in $Range) {
            $tree.Add($record)
        }
    }

    Add-Member -InputObject $Constructor -MemberType ScriptMethod -Name Add -Value {
        param($record)
        $tree.Add($record)
    }

    Add-Member -InputObject $Constructor -MemberType ScriptMethod -Name Build -Value {
        $this.Tree.Build()
    }

    return $Constructor
}

$Static.Filter = {
    param([Bool]$Apply)
    if ($Apply) {
        if ($this.Parent) {
            $this.Filtered = $this.Parent
        }
        else {
            $this.Filtered = $this.TreeView
        }

        if ($this.Filtered) {
            $this.Remove()
        }
        else {
            return
        }

        if ($this.Filtered.Nodes.Count -eq 0 -and $this.Filtered -isnot [System.Windows.Forms.TreeView]) {
            $this.Filtered.Filter($Apply)
        }
    }
    else {
        $parent = $this.Filtered
        $this.Filtered = $null
        if ($parent) {
            [void]$parent.Nodes.Add($this)
        }

        if ($parent -and $parent -isnot [System.Windows.Forms.TreeView]) {
            $parent.Filter($Apply)
        }
    }
}

###############################################################################
# Default Handlers
$default = @{}
$default.AfterCheck = {
    param($sender, $e)

    if($e.Action -eq [System.Windows.Forms.TreeViewAction]::Unknown) {
        return
    }

    ForEach ($child in $e.Node.Nodes) {
        $child.SetChecked($e.Node.Checked)
    }
}
$default.Click = {
    Param(
        # sender: the control initiating the event.
        [Parameter(Mandatory = $true, Position = 0)]
            [System.Windows.Forms.TreeView]
            $sender,

        # e: event arguments passed by the sender.
        [Parameter(Mandatory = $true, Position = 1)]
            [system.EventArgs]
            $e
    )

    # Ignore $e.Location as the coordinates are clipped to the client area of the treeview,
    # but treeview.GetNodeAt() expects full screen area coordinates.  Seems like an unusual
    # way to implement that functionality...

    # Get the TreeViewNode that was clicked (Right or Left)
    $Target = $sender.GetNodeAt($sender.PointToClient([System.Windows.Forms.Control]::MousePosition))

    if ($Target -ne $null) {
        $sender.SelectedNode = $Target
    }
}


function New-SortedTreeView {
    param (
        # Window reference for docking support (undock window parent).
        [Parameter(Mandatory = $true)]
            [System.Windows.Forms.Form]
            $Window,

        # Parent layout control that hosts this SortedTreeView for docking support (dock target).
        [Parameter(Mandatory = $true)]
            [System.Windows.Forms.Control]
            $Parent,

        # TreeView data source used to create the TreeNodes.
        [Parameter(Mandatory = $true)]
            [AllowEmptyCollection()]
            [System.Collections.ArrayList]
            $Source,
        
        # Collection of images used by TreeNodes.
        [Parameter(Mandatory = $true)]
            [System.Windows.Forms.ImageList]
            $ImageList,

        # Object containing the property overrides, event handlers, and custom properties/methods for the TreeView layout container.
        [Parameter(Mandatory = $false)]
            [PSCustomObject]
            $TreeDefinition,

        # Object containing the property overrides, event handlers, and custom properties/methods for TreeNodes used for data grouping (branches).
        [Parameter(Mandatory = $false)]
            [PSCustomObject]
            $GroupDefinition,
        
        # Object containing the property overrides, event handlers, and custom properties/methods for TreeNodes used for data representation (leaves)
        [Parameter(Mandatory = $false)]
            [PSCustomObject]
            $NodeDefinition
    )

    # Top Level Container
    $Container = New-Object System.Windows.Forms.TabControl
        $Container.Dock = [System.Windows.Forms.DockStyle]::Fill
        
        # Attached to Parent Control by Module Component Registration Function

    $TreeViewTab = New-Object System.Windows.Forms.TabPage
        $TreeViewTab.Dock = [System.Windows.Forms.DockStyle]::Fill
        $TreeViewTab.Text = "Explorer"

        [void]$Container.TabPages.Add( $TreeViewTab )

    # TreeView Component
    #region
    $TreeViewControl = New-Object System.Windows.Forms.TreeView
    &{
        $TreeViewControl.Dock = [System.Windows.Forms.DockStyle]::Fill
        $TreeViewControl.CheckBoxes = $true
        $TreeViewControl.ImageList = $ImageList

        # Late binding (Dynamic)
        foreach ($property in $TreeDefinition.Properties.GetEnumerator()) {
            $TreeViewControl.($property.Key) = $property.Value
        }

        # Late binding (Dynamic)
        foreach ($handler in $TreeDefinition.Handlers.GetEnumerator()) {
            $TreeViewControl."Add_$($handler.Key)"($handler.Value)
        }

        # Default Handlers
        if (!$TreeDefinition.Handlers.ContainsKey('AfterCheck')) {
            [void]$TreeViewControl.Add_AfterCheck($default.AfterCheck)
        }
        if (!$TreeDefinition.Handlers.ContainsKey('Click')) {
            [void]$TreeViewControl.Add_Click($default.Click)
        }

        Add-Member -InputObject $TreeViewControl -MemberType NoteProperty -Name Static -Value $Static

        Add-Member -InputObject $TreeViewControl -MemberType NoteProperty -Name Source -Value $Source

        Add-Member -InputObject $TreeViewControl -MemberType NoteProperty -Name Excluded -Value (New-Object System.Collections.ArrayList)

        Add-Member -InputObject $TreeViewControl -MemberType NoteProperty -Name DataNodes -Value $null

        Add-Member -InputObject $TreeViewControl -MemberType ScriptMethod -Name FindNode -Value {
            param([String]$NodeName, [System.Windows.Forms.TreeNode]$StartNode)

            # Top Level search using the TreeView as root
            if (!$StartNode) {
                foreach ($node in $this.Nodes) {
                    $seek = $this.FindNode($NodeName, $node)
                    if ($seek) {
                        return $seek
                    }
                }
            }

            # Recursive search start node
            foreach ($node in $StartNode.Nodes) {
                if ($node.Name -eq $NodeName) {
                    return $node
                }
                $seek = $this.FindNode($NodeName, $node)
                if ($seek) {
                    return $seek
                }
            }

            # Default output
            return $null
        }

        foreach ($method in $TreeDefinition.Methods.GetEnumerator()) {
            Add-Member -InputObject $TreeViewcontrol -MemberType ScriptMethod -Name $method.Key -Value $method.Value
        }

        [void]$TreeViewTab.Controls.Add( $TreeViewControl )
    }

    $SettingsTab = New-Object System.Windows.Forms.TabPage
    &{
        $SettingsTab.Dock = [System.Windows.Forms.DockStyle]::Fill
        $SettingsTab.Text = "View Settings"

        Add-Member -InputObject $SettingsTab -MemberType NoteProperty -Name Fields -Value $null

        Add-Member -InputObject $SettingsTab -MemberType NoteProperty -Name SortFields -Value (New-Object System.Collections.ArrayList)

        Add-Member -InputObject $SettingsTab -MemberType NoteProperty -Name GroupFields -Value (New-Object System.Collections.ArrayList)

        Add-Member -InputObject $SettingsTab -MemberType ScriptMethod -Name PromptUser -Value {
            [System.Windows.Forms.MessageBox]::Show(
                "Update View Settings",
                "Data fields or format changed.`r`nUpdate View Settings.",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Exclamation
            )

            $this.Parent.SelectedIndex = 1
        }

        Add-Member -InputObject $SettingsTab -MemberType ScriptMethod -Name RegisterFields -Value {
            param($fields)
            # Prompt Flag if the View Settings needs to be updated by the user
            $RequireUpdate = $false

            if ($this.Fields.Count -ne $fields.Count) {
                $RequiredUpdate = $true
                $this.Fields = $fields
                $this.DataLabel.DataSource = @([String]::Empty) + @($fields)
            }

            # Verify DataSource is set for the DataLabel Selector
            elseif (!$this.DataLabel.DataSource) {
                $RequireUpdate = $true
                $this.Fields = $fields
                $this.DataLabel.DataSource = @([String]::Empty) + @($fields)
            }

            # Verify DataSource fields matches input
            else {
                foreach ($f in $fields) {
                    if ($this.DataLabel.DataSource -notcontains $f) {
                        $RequireUpdate = $true
                        $this.DataLabel.DataSource = @([String]::Empty) + @($fields)
                        break
                    }
                }
            }

            # Validate grouped fields exist on data objects
            foreach ($field in $this.GroupFields) {
                if ($fields -notcontains $field.Name) {
                    $field.Setting.Unregister()
                    $this.DataLabel.DataSource = @([String]::Empty) + @($fields)
                    $RequireUpdate = $true
                }
            }

            if ($RequireUpdate) {
                $this.Handler.Valid = $false
            }
        }

        [void]$Container.TabPages.Add( $SettingsTab )
    }

    # View Settings Flow Panel
    $SettingsPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    &{
        $SettingsPanel.Dock          = [System.Windows.Forms.DockStyle]::Fill
        $SettingsPanel.BackColor     = [System.Drawing.Color]::White
        $SettingsPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
        $SettingsPanel.WrapContents  = $false
        $SettingsPanel.AutoScroll    = $true

        # Data Node Configuration (Apply and Undock Buttons)
        $DataNodePanel = New-Object System.Windows.Forms.Panel
        &{
            $DataNodePanel.BackColor = [System.Drawing.Color]::White

            $Registration = [PSCustomObject]@{
                Name = [String]::Empty
                Setting = $DataNodePanel
            }

            # Static property for the sort level settings to check when removing registrations
            Add-Member -InputObject $DataNodePanel -MemberType NoteProperty -Name SettingType -Value 'DataNode'

            Add-Member -InputObject $DataNodePanel -MemberType NoteProperty -Name SettingsTab -Value $SettingsTab

            Add-Member -InputObject $DataNodePanel -MemberType NoteProperty -Name Registered -Value $false

            Add-Member -InputObject $DataNodePanel -MemberType ScriptMethod -Name Unregister -Value {
                if(!$this.Registered) {
                    return
                }
                $this.Registered = $false
                Write-Debug ("Unregistered: {0}`t {1}" -f $this.Registration.Name, $this.GetHashCode())
                [void]$this.SettingsTab.GroupFields.Remove($this.Registration)
                $this.Controls['DisplayFieldSelectionBox'].SelectedIndex = 0
            }

            Add-Member -InputObject $DataNodePanel -MemberType NoteProperty -Name Registration -Value $Registration

            # Undock Navigation TreeView Button
            [void]$DataNodePanel.Controls.Add((&{
                # UNDOCK Button
                $Button = New-Object System.Windows.Forms.Button
                $Button.Dock = [System.Windows.Forms.DockStyle]::Left
                $Button.Text = "Undock"
                $Button.Height = 21
                $Button.Add_Click({
                    $this.Undock()
                })

                Add-Member -InputObject $Button -MemberType NoteProperty -Name Window -Value $Window

                Add-Member -InputObject $Button -MemberType NoteProperty -Name DockPackage -Value $Container

                Add-Member -InputObject $Button -MemberType NoteProperty -Name DockTarget -Value $Parent

                Add-Member -InputObject $Button -MemberType ScriptMethod -Name Undock -Value {
                    $form = New-Object System.Windows.Forms.Form
                    $form.Text          = "Baseline Security Scan Viewer"
                    $form.Size          = New-Object System.Drawing.Size(300, 600)
                    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

                    $form.Add_Closing({
                        param($sender, $e)
                        $this.Redock()
                    })

                    Add-Member -InputObject $form -MemberType NoteProperty -Name DockPackage -Value $this.DockPackage
                    Add-Member -InputObject $form -MemberType NoteProperty -Name DockTarget -Value $this.DockTarget
                    Add-Member -InputObject $form -MemberType NoteProperty -Name DockButton -Value $this
                    Add-Member -InputObject $form -MemberType ScriptMethod -Name Redock -Value {
                        $this.SuspendLayout()
                        $this.DockPackage.SuspendLayout()
                        $this.DockTarget.SuspendLayout()

                        $this.Controls.Clear()
                        $this.DockTarget.Controls.Add($this.DockPackage)
                        $this.DockTarget.Parent.Panel1Collapsed = $false
                        $this.DockButton.Enabled = $true

                        $this.DockTarget.ResumeLayout()
                        $this.DockPackage.ResumeLayout()
                        $this.DockTarget.PerformLayout()
                    }

                    $this.DockPackage.SuspendLayout()
                    $this.DockTarget.SuspendLayout()

                    $this.DockTarget.Controls.Clear()
                    $this.DockTarget.Parent.Panel1Collapsed = $true
                    $this.Enabled = $false

                    $form.Controls.Add($this.DockPackage)

                    $this.DockTarget.ResumeLayout()
                    $this.DockPackage.ResumeLayout()
                    [void]$form.Show($this.Window)
                }

                # Add Button to Parent Control
                return $Button
            }))

            # Apply Settings Button
            [void]$DataNodePanel.Controls.Add( 
            ( &{
                $Apply = New-Object System.Windows.Forms.Button

                $Handler = {
                    Write-Debug "Applying View Settings..."

                    # Data Node Label Field
                    $DataLabel = $this.Parent.Controls['DisplayFieldSelectionBox'].SelectedValue

                    if ([String]::IsNullOrEmpty($DataLabel)) {
                        [System.Windows.Forms.MessageBox]::Show(
                            "You must select a [Data Node Label] field!",
                            "Compliance View Settings",
                            [System.Windows.Forms.MessageBoxButtons]::OK,
                            [System.Windows.Forms.MessageBoxIcon]::Exclamation
                        )
                        return
                    }

                    $this.Valid = $true

                    # Gather Sorting Data
                    #$SortBy = New-Object System.Collections.ArrayList
                    #$SortSettings = $this.SettingsContainer.Controls['SortByOptions']
                    #for ($i = 1; $i -lt $SortSettings.Controls.Count; $i++) {
                    #    [void]$SortBy.Add($SortSettings.Controls[$i].Registration.Name)
                    #}
                    #$SortBy.TrimToSize()
                    #
                    #Write-Debug "Building Sort Buckets: $SortBy"

                    # Gather Grouping Data
                    $GroupBy = New-Object System.Collections.ArrayList
                    $GroupSettings = $this.SettingsContainer.Controls['GroupByOptions']
                    for ($i = 1; $i -lt $GroupSettings.Controls.Count; $i++) {
                        [void]$GroupBy.Add($GroupSettings.Controls[$i].Registration)
                    }
                    $GroupBy.TrimToSize()

                    Write-Debug "Building Group Buckets: $GroupBy"

                    # Quick Access ArrayList for all of the data nodes
                    $DataNodes = New-Object System.Collections.ArrayList

                    $Constructor = & $this.Tree.Static.NewGroupConstructor $GroupBy $DataLabel $DataNodes $this.GroupDefinition $this.NodeDefinition
                    $Constructor.AddRange($this.Tree.Source)

                    $this.Tree.SuspendLayout()

                    $this.Tree.Nodes.Clear()
                    $this.Tree.Nodes.AddRange( $Constructor.Build() )

                    # Append Quick Access Data Node List
                    $this.Tree.DataNodes = $DataNodes
                    $this.Component.FilterTab.Handler.Reset($this.Component.FilterTab.Handler.Applied)

                    $this.Tree.ResumeLayout()

                    $this.Component.SelectedIndex = 0
                }

                # TREEVIEW SETTINGS -------------------------------------------
                Add-Member -InputObject $Apply -MemberType NoteProperty -Name SettingsContainer -Value $SettingsPanel

                # PARENT AND DATA SOURCE CONTAINERS/OBJECTS -------------------
                Add-Member -InputObject $Apply -MemberType NoteProperty -Name Tree -Value $TreeViewControl
                Add-Member -InputObject $Apply -MemberType NoteProperty -Name Component -Value $Container

                # TREEVIEW NODE CUSTOMIZATION DEFINITIONS ---------------------
                Add-Member -InputObject $Apply -MemberType NoteProperty -Name GroupDefinition -Value $GroupDefinition
                Add-Member -InputObject $Apply -MemberType NoteProperty -Name NodeDefinition -Value $NodeDefinition

                # REMOTE CALL HANDLE ------------------------------------------
                Add-Member -InputObject $Apply -MemberType ScriptMethod -Name Apply -Value $Handler

                # Flag for remote callers set by Apply.Handler and SettingsTab.RegisterFields
                Add-Member -InputObject $Apply -MemberType NoteProperty -Name Valid -Value $false
                Add-Member -InputObject $SettingsTab -MemberType NoteProperty -Name Handler -Value $Apply

                $Apply.Dock   = [System.Windows.Forms.DockStyle]::Left
                $Apply.Height = 21
                $Apply.Width  = 85
                $Apply.Text   = "Apply Settings"
                $Apply.Add_Click($Handler)

                return $Apply
        
                })
            )

            # Select Data Label Combo Box
            [void]$DataNodePanel.Controls.Add(
            (&{
                $LabelSelector = New-Object System.Windows.Forms.ComboBox
                $LabelSelector.Name = "DisplayFieldSelectionBox"
                $LabelSelector.Dock = [System.Windows.Forms.DockStyle]::Left
                $LabelSelector.Width = 85
                $LabelSelector.Add_SelectedValueChanged({
                    $Current = $this.Parent # Settings Panel

                    if ($this.SelectedValue -eq [String]::Empty) {
                        $Current.Unregister()
                        return
                    }

                    $registered = $this.SettingsTab.GroupFields.ToArray()
                    foreach ($field in $registered) {
                        if (!$field.Setting.Equals($Current) -and $field.Name -eq $this.SelectedValue) {
                            if ($field.Setting.SettingType -eq 'Group') {
                                $field.Setting.Unregister()
                            }
                        }
                    }
        
                    Write-Debug ("Field Registered: {0}`t{1}" -f $Current.Registration.Name, $Current.GetHashCode())
                    $Current.Registration.Name = $this.SelectedValue

                    if (!$Current.Registered) {
                        [void]$this.SettingsTab.GroupFields.Add($Current.Registration)
                        $Current.Registered = $true
                    }
                })

                Add-Member -InputObject $SettingsTab -MemberType NoteProperty -Name DataLabel -Value $LabelSelector

                Add-Member -InputObject $LabelSelector -MemberType NoteProperty -Name SettingsTab -Value $SettingsTab

                return $LabelSelector
            }) )

            # Panel Setting Label
            [void]$DataNodePanel.Controls.Add(
            (&{
                $Label= New-Object System.Windows.Forms.Label
                $Label.Dock = [System.Windows.Forms.DockStyle]::Left
                $Label.Text = "Data Node Label:"
                $Label.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft

                return $Label
            }) )

            $TmpWidth = 0
            foreach ($TmpItem in $DataNodePanel.Controls) {
                $TmpWidth += $TmpItem.width
            }
            $DataNodePanel.Width = $TmpWidth + 10
            $DataNodePanel.Height = 22

            [void]$SettingsPanel.Controls.Add( $DataNodePanel )
        }

        # Sort By Options Flow Panel
        #region
    #    $OptionsFlowPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    #    $OptionsFlowPanel.Name = 'SortByOptions'
    #    $OptionsFlowPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
    #    $OptionsFlowPanel.WrapContents = $false
    #    $OptionsFlowPanel.BackColor = [System.Drawing.Color]::White
    #    $OptionsFlowPanel.AutoSize = $true
    #    $OptionsFlowPanel.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    #
    #    $SettingsPanel.Controls.Add( $OptionsFlowPanel )
    #
    #    # Add Sort Level Static Panel
    #    #region
    #    $OptionsFlowPanel.Controls.Add( (New-Object System.Windows.Forms.Panel) )
    #    $OptionsFlowPanel.Controls[0].Height = 22
    #
    #    $OptionsFlowPanel.Controls[0].Controls.Add( (New-Object System.Windows.Forms.Label) )
    #    $OptionsFlowPanel.Controls[0].Controls[0].Dock = [System.Windows.Forms.DockStyle]::Left
    #    $OptionsFlowPanel.Controls[0].Controls[0].Text = "Sort Level"
    #    $OptionsFlowPanel.Controls[0].Controls[0].TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    #
    #    $OptionsFlowPanel.Controls[0].Controls.Add( (New-Object System.Windows.Forms.Button) )
    #    $OptionsFlowPanel.Controls[0].Controls[1].Text = "Add"
    #    $OptionsFlowPanel.Controls[0].Controls[1].Dock = [System.Windows.Forms.DockStyle]::Left
    #    $OptionsFlowPanel.Controls[0].Controls[1].Height = 21
    #    $OptionsFlowPanel.Controls[0].Controls[1].Width  = 50
    #    $OptionsFlowPanel.Controls[0].Controls[1].Add_Click({
    #        if ($SettingsTab.GroupFields.Count -eq $SettingsTab.Fields.Count -1) {
    #            [System.Windows.Forms.MessageBox]::Show(
    #                "Maximum number of group/sort levels reached!",
    #                "View Settings",
    #                [System.Windows.Forms.MessageBoxButtons]::OK,
    #                [System.Windows.Forms.MessageBoxIcon]::Exclamation
    #            )
    #            return
    #        }
    #
    #        Write-Debug ("Registered Fields:`r`n{0}" -f ($SettingsTab.GroupFields | Format-Table | Out-String))
    #        
    #        # Find first unregistered field name
    #        $filtered = New-Object System.Collections.ArrayList
    #        [void]$filtered.AddRange($SettingsTab.Fields)
    #        foreach ($registration in $SettingsTab.GroupFields) {
    #            [void]$filtered.Remove($registration.Name)
    #        }
    #        $available = $filtered[0]
    #        
    #        Write-Debug ("Available Field:`r`n{0}" -f $available)
    #
    #        $RemoveButton = New-Object System.Windows.Forms.Button
    #        $RemoveButton.Dock = [System.Windows.Forms.DockStyle]::Left
    #        $RemoveButton.Text = "Remove"
    #        $RemoveButton.Height = 21
    #        $RemoveButton.Width = 75
    #        [void]$RemoveButton.Add_Click({
    #            $this.Parent.Unregister()
    #        })
    #
    #        $SettingLabel = New-Object System.Windows.Forms.Label
    #        $SettingLabel.Dock = [System.Windows.Forms.DockStyle]::Left
    #        $SettingLabel.Width = 60
    #        $SettingLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    #        $SettingLabel.Text = "Sort by:"
    #
    #        $FieldSelectBox = New-Object System.Windows.Forms.ComboBox
    #        $FieldSelectBox.Dock = [System.Windows.Forms.DockStyle]::Left
    #        $FieldSelectBox.DataSource = @($SettingsTab.Fields)
    #        $FieldSelectBox.Width = 85
    #        
    #
    #        $SortDirection = New-Object System.Windows.Forms.ComboBox
    #        $SortDirection.Dock = [System.Windows.Forms.DockStyle]::Left
    #        $SortDirection.DataSource = @("Ascending", "Descending")
    #        $SortDirection.Width = 80
    #        $SortDirection.Add_SelectedValueChanged({
    #            $this.Parent.Registration.SortDirection = $this.SelectedValue
    #        })
    #
    #        $SettingPanel = New-Object System.Windows.Forms.Panel
    #        $SettingPanel.Height = 22
    #        $SettingPanel.Width = $RemoveButton.Width + $SortDirection.Width + $SettingLabel.Width + $FieldSelectBox.Width
    #
    #        [void]$SettingPanel.Controls.Add($SortDirection)
    #        [void]$SettingPanel.Controls.Add($FieldSelectBox)
    #        [void]$SettingPanel.Controls.Add($SettingLabel)
    #        [void]$SettingPanel.Controls.Add($RemoveButton)
    #
    #        $registration = [PSCustomObject]@{
    #            Name = $available
    #            SortDirection = 'Ascending'
    #            Setting = $SettingPanel
    #        }
    #
    #        Add-Member -InputObject $SettingPanel -MemberType NoteProperty -Name Registration -Value $registration
    #
    #        Add-Member -InputObject $SettingPanel -MemberType NoteProperty -Name SettingType -Value "Sort"
    #
    #        Add-Member -InputObject $SettingPanel -MemberType ScriptMethod -Name Unregister -Value {
    #            [void]$this.Parent.Controls.Remove($this)
    #            $SettingsTab.GroupFields.Remove($this.Registration)
    #            Write-Debug ("Unregistered: {0}`t {1}" -f $this.Registration.Name, $this.GetHashCode())
    #        }
    #
    #        [void]$SettingsTab.GroupFields.Add($registration)
    #
    #        Write-Debug ("Field Registered: {0}`t{1}" -f $registration.Name, $registration.Setting.GetHashCode())
    #
    #        $GroupSortPanel = $this.Parent.Parent
    #        [void]$GroupSortPanel.Controls.Add($SettingPanel)
    #        
    #        # After registering the field set the combo box selected value
    #        $FieldSelectBox.SelectedItem = $available
    #
    #        # Adding the event handler last to prevent issues with the initial loading of the setting panel
    #        $FieldSelectBox.Add_SelectedValueChanged({
    #            $Current = $this.Parent
    #
    #            $registered = $SettingsTab.GroupFields.ToArray()
    #            foreach ($field in $registered) {
    #                if (!$field.Setting.Equals($Current) -and $field.Name -eq $this.SelectedValue) {
    #
    #                    # Sorting Levels do not unregister the data node display field
    #                    if ($field.Setting.SettingType -eq 'DataNode') {
    #                        continue
    #                    }
    #                    $field.Setting.Unregister()
    #                }
    #            }
    #
    #            $Current.Registration.Name = $this.SelectedValue
    #            Write-Debug ("Field Registered: {0}`t{1}" -f $Current.Registration.Name, $Current.GetHashCode())
    #        })
    #    })
    #    #endregion
        #endregion

        # Group By Options Flow Panel
        $OptionsFlowPanel = New-Object System.Windows.Forms.FlowLayoutPanel
        &{
            $OptionsFlowPanel.Name          = 'GroupByOptions'
            $OptionsFlowPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
            $OptionsFlowPanel.WrapContents  = $false
            $OptionsFlowPanel.BackColor     = [System.Drawing.Color]::White
            $OptionsFlowPanel.AutoSize      = $true
            $OptionsFlowPanel.AutoSizeMode  = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink

            [void]$SettingsPanel.Controls.Add( $OptionsFlowPanel )
        }

        # Add Group Level Static Panel
        &{
            $GroupStaticPanel = New-Object System.Windows.Forms.Panel
            $GroupStaticPanel.Height = 22
                
                [void]$OptionsFlowPanel.Controls.Add( $GroupStaticPanel )

            $Label = New-Object System.Windows.Forms.Label
                $Label.Dock = [System.Windows.Forms.DockStyle]::Left
                $Label.Text = "Group Level"
                $Label.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft

                [void]$GroupStaticPanel.Controls.Add( $Label )
            
            $GroupButton = New-Object System.Windows.Forms.Button
                $GroupButton.Text = "Add"
                $GroupButton.Dock = [System.Windows.Forms.DockStyle]::Left
                $GroupButton.Height = 21
                $GroupButton.Width  = 50
                
                [void]$GroupStaticPanel.Controls.Add( $GroupButton )

                Add-Member -InputObject $GroupButton -MemberType NoteProperty -Name SettingsTab -Value $SettingsTab

                Add-Member -InputObject $GroupButton -MemberType NoteProperty -Name OptionsPanel -Value $OptionsFlowPanel

                [void]$GroupButton.Add_Click({
                    if ($this.SettingsTab.GroupFields.Count -eq $this.SettingsTab.Fields.Count -1) {
                        [System.Windows.Forms.MessageBox]::Show(
                            "Maximum number of group/sort levels reached!",
                            "View Settings",
                            [System.Windows.Forms.MessageBoxButtons]::OK,
                            [System.Windows.Forms.MessageBoxIcon]::Exclamation
                        )
                        return
                    }

                    Write-Debug ("Registered Fields:`r`n{0}" -f ($this.SettingsTab.GroupFields | Format-Table | Out-String))
        
                    # Find first unregistered field name
                    $filtered = New-Object System.Collections.ArrayList
                    [void]$filtered.AddRange($this.SettingsTab.Fields)
                    foreach ($registration in $this.SettingsTab.GroupFields) {
                        [void]$filtered.Remove($registration.Name)
                    }
                    $available = $filtered[0]
        
                    Write-Debug ("Avaliable Field:`r`n{0}" -f $available)

                    $RemoveButton = New-Object System.Windows.Forms.Button
                    $RemoveButton.Dock = [System.Windows.Forms.DockStyle]::Left
                    $RemoveButton.Text = "Remove"
                    $RemoveButton.Height = 21
                    $RemoveButton.Width = 75
                    [void]$RemoveButton.Add_Click({
                        $this.Parent.Unregister()
                    })

                    $SettingLabel = New-Object System.Windows.Forms.Label
                    $SettingLabel.Dock = [System.Windows.Forms.DockStyle]::Left
                    $SettingLabel.Width = 60
                    $SettingLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
                    $SettingLabel.Text = "Group by:"

                    $FieldSelectBox = New-Object System.Windows.Forms.ComboBox
                    $FieldSelectBox.Dock = [System.Windows.Forms.DockStyle]::Left
                    $FieldSelectBox.DataSource = @($this.SettingsTab.Fields)
                    $FieldSelectBox.Width = 85

                    Add-Member -InputObject $FieldSelectBox -MemberType NoteProperty -Name SettingsTab -Value $this.SettingsTab

                    $SortDirection = New-Object System.Windows.Forms.ComboBox
                    $SortDirection.Dock = [System.Windows.Forms.DockStyle]::Left
                    $SortDirection.DataSource = @("Ascending", "Descending")
                    $SortDirection.Width = 80
                    $SortDirection.Add_SelectedValueChanged({
                        $this.Parent.Registration.SortDirection = $this.SelectedValue
                    })

                    $SettingPanel = New-Object System.Windows.Forms.Panel
                    $SettingPanel.Height = 22
                    $SettingPanel.Width = $RemoveButton.Width + $SortDirection.Width + $SettingLabel.Width + $FieldSelectBox.Width

                    [void]$SettingPanel.Controls.Add($SortDirection)
                    [void]$SettingPanel.Controls.Add($FieldSelectBox)
                    [void]$SettingPanel.Controls.Add($SettingLabel)
                    [void]$SettingPanel.Controls.Add($RemoveButton)

                    $registration = [PSCustomObject]@{
                        Name = $available
                        SortDirection = "Ascending"
                        Setting = $SettingPanel
                    }

                    Add-Member -InputObject $SettingPanel -MemberType NoteProperty -Name SettingsTab -Value $this.SettingsTab

                    Add-Member -InputObject $SettingPanel -MemberType NoteProperty -Name Registration -Value $registration

                    Add-Member -InputObject $SettingPanel -MemberType NoteProperty -Name SettingType -Value "Group"

                    Add-Member -InputObject $SettingPanel -MemberType ScriptMethod -Name Unregister -Value {
                        [void]$this.Parent.Controls.Remove($this)
                        $this.SettingsTab.GroupFields.Remove($this.Registration)
                        Write-Debug ("Unregistered: {0}`t {1}" -f $this.Registration.Name, $this.GetHashCode())
                    }

                    # Register the setting
                    [void]$this.SettingsTab.GroupFields.Add($registration)

                    Write-Debug ("Field Registered: {0}`t{1}" -f $registration.Name, $registration.Setting.GetHashCode())

                    [void]$this.OptionsPanel.Controls.Add($SettingPanel)
        
                    # After registering the field set the combo box selected value
                    $FieldSelectBox.SelectedItem = $available

                    # Adding the event handler last to prevent issues with the initial loading of the setting panel
                    $FieldSelectBox.Add_SelectedValueChanged({
                        $Current = $this.Parent # Setting Panel

                        $registered = $this.SettingsTab.GroupFields.ToArray()
                        foreach ($field in $registered) {
                            if (!$field.Setting.Equals($Current) -and $field.Name -eq $this.SelectedValue) {
                                $field.Setting.Unregister()
                            }
                        }

                        $Current.Registration.Name = $this.SelectedValue
                        Write-Debug ("Field Registered: {0}`t{1}" -f $Current.Registration.Name, $Current.GetHashCode())
                    })
                }) # End Button.Add_Click
        }

        # Add the TreeView Settings Panel
        [void]$SettingsTab.Controls.Add( $SettingsPanel )
    }

    $FilterTab = New-Object System.Windows.Forms.TabPage
    &{
        $FilterTab.Dock = [System.Windows.Forms.DockStyle]::Fill
        $FilterTab.Text = "Filter Settings"

        Add-Member -InputObject $FilterTab -MemberType NoteProperty -Name Fields -Value $null

        Add-Member -InputObject $FilterTab -MemberType NoteProperty -Name FilterFields -Value (New-Object System.Collections.ArrayList)

        Add-Member -InputObject $FilterTab -MemberType ScriptMethod -Name PromptUser -Value {
            [System.Windows.Forms.MessageBox]::Show(
                "Update View Settings",
                "Data fields or format changed.`r`nUpdate View Settings.",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Exclamation
            )

            $this.Parent.SelectedIndex = 1
        }

        Add-Member -InputObject $FilterTab -MemberType ScriptMethod -Name RegisterFields -Value {
            param($fields)
            # Prompt Flag if the View Settings needs to be updated by the user
            $RequireUpdate = $false

            if ($this.Fields.Count -ne $fields.Count) {
                $RequiredUpdate = $true
                $this.Fields = $fields
            }

            # Validate grouped fields exist on data objects
            foreach ($field in $this.FilterFields) {
                if ($fields -notcontains $field.Name) {
                    $field.Setting.Unregister()
                    $this.DataLabel.DataSource = @([String]::Empty) + @($fields)
                    $RequireUpdate = $true
                }
            }

            if ($RequireUpdate) {
                $this.Handler.Valid = $false
            }
        }

        [void]$Container.TabPages.Add( $FilterTab )
    }

    # Filter Settings Flow Panel
    $FilterPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    &{
        $FilterPanel.Dock          = [System.Windows.Forms.DockStyle]::Fill
        $FilterPanel.BackColor     = [System.Drawing.Color]::White
        $FilterPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
        $FilterPanel.WrapContents  = $false
        $FilterPanel.AutoScroll    = $true

        # (Apply)
        $FilterControlPanel = New-Object System.Windows.Forms.Panel
        &{
            $FilterControlPanel.BackColor = [System.Drawing.Color]::White

            Add-Member -InputObject $FilterControlPanel -MemberType NoteProperty -Name FilterTab -Value $FilterTab

            # Apply Settings Button
            [void]$FilterControlPanel.Controls.Add( 
            ( &{
                $Apply = New-Object System.Windows.Forms.Button

                $Reset   = {
                    param([bool]$apply = $false)
                    $this.Applied = $false
                    $this.Valid   = $false
                    $this.Parsed  = $false
                    $this.Text    = "Apply"
                    $this.Tree.Excluded.Clear()
                    if ($apply) {
                        $this.Apply()
                    }
                }

                $Handler = {
                    Write-Debug "Applying Filter Settings..."
                    
                    $this.Valid = $true

                    # Unapply
                    if ($this.Applied) {
                        $this.Applied = !$this.Applied
                        $this.Text = "Apply"

                        $this.Tree.SuspendLayout()
                        foreach ($node in $this.Tree.Excluded) {
                            [void]$node.Filter($false)
                        }
                        $this.Tree.ResumeLayout()
                        $this.Component.SelectedIndex = 0
                        return
                    }

                    # Select Nodes to Filter
                    if ($this.Valid -and !$this.Parsed) {
                        $this.Parsed = $true
                        $this.Tree.Excluded.Clear()
                        $excluded = $this.Tree.Excluded

                        # Gather Filter Data
                        $FilterBy = New-Object System.Collections.ArrayList
                        $FilterSettings = $this.SettingsContainer.Controls['FilterByOptions']
                        for ($i = 1; $i -lt $FilterSettings.Controls.Count; $i++) {
                            [void]$FilterBy.Add($FilterSettings.Controls[$i].Registration)
                        }
                        $FilterBy.TrimToSize()
                        Write-Debug "Filtering: $FilterBy"
                        
                        foreach ($node in $this.Tree.DataNodes) {
                            foreach($filter in $FilterBy) {
                                Switch ($filter.FilterType) {
                                    Include {
                                        if ($node.Tag.($filter.Name) -notmatch $filter.FilterValue) {
                                            if (!$excluded.Contains($node)) {
                                                [void]$excluded.Add($node)
                                            }
                                        }
                                    }

                                    Exclude {
                                        if ($node.Tag.($filter.Name) -match $filter.FilterValue) {
                                            if (!$excluded.Contains($node)) {
                                                [void]$excluded.Add($node)
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    # Apply Filter
                    if (!$this.Applied) {
                        $this.Applied = !$this.Applied
                        $this.Text = "Unapply"
                        $this.Tree.SuspendLayout()
                        foreach ($node in $this.Tree.Excluded) {
                            $node.Filter($true)
                        }
                        $this.Tree.ResumeLayout()
                        $this.Component.SelectedIndex = 0
                    }
                }

                # TREEVIEW SETTINGS -------------------------------------------
                Add-Member -InputObject $Apply -MemberType NoteProperty -Name SettingsContainer -Value $FilterPanel

                # PARENT AND DATA SOURCE CONTAINERS/OBJECTS -------------------
                Add-Member -InputObject $Apply -MemberType NoteProperty -Name Tree -Value $TreeViewControl
                Add-Member -InputObject $Apply -MemberType NoteProperty -Name Component -Value $Container

                # REMOTE CALL HANDLE ------------------------------------------
                Add-Member -InputObject $Apply -MemberType ScriptMethod -Name Apply -Value $Handler
                Add-Member -InputObject $Apply -MemberType ScriptMethod -Name Reset -Value $Reset

                # Flag for remote callers set by Apply.Handler and FilterTab.RegisterFields
                Add-Member -InputObject $Apply -MemberType NoteProperty -Name Valid -Value $false
                Add-Member -InputObject $Apply -MemberType NoteProperty -Name Applied -Value $false
                Add-Member -InputObject $Apply -MemberType NoteProperty -Name Parsed -Value $false
                Add-Member -InputObject $FilterTab -MemberType NoteProperty -Name Handler -Value $Apply

                $Apply.Dock   = [System.Windows.Forms.DockStyle]::Left
                $Apply.Height = 21
                $Apply.Width  = 85
                $Apply.Text   = "Apply"
                $Apply.Add_Click($Handler)

                return $Apply
        
                })
            )

            $TmpWidth = 0
            foreach ($TmpItem in $FilterControlPanel.Controls) {
                $TmpWidth += $TmpItem.width
            }
            $FilterControlPanel.Width = $TmpWidth + 10
            $FilterControlPanel.Height = 22

            [void]$FilterPanel.Controls.Add( $FilterControlPanel )
        }

        # Filter By Options Flow Panel
        $OptionsFlowPanel = New-Object System.Windows.Forms.FlowLayoutPanel
        &{
            $OptionsFlowPanel.Name          = 'FilterByOptions'
            $OptionsFlowPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
            $OptionsFlowPanel.WrapContents  = $false
            $OptionsFlowPanel.BackColor     = [System.Drawing.Color]::White
            $OptionsFlowPanel.AutoSize      = $true
            $OptionsFlowPanel.AutoSizeMode  = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink

            [void]$FilterPanel.Controls.Add( $OptionsFlowPanel )
        }

        # Add Filter Level Static Panel
        &{
            $FilterStaticPanel = New-Object System.Windows.Forms.Panel
            $FilterStaticPanel.Height = 22
                
                [void]$OptionsFlowPanel.Controls.Add( $FilterStaticPanel )

            $Label = New-Object System.Windows.Forms.Label
                $Label.Dock = [System.Windows.Forms.DockStyle]::Left
                $Label.Text = "Filters"
                $Label.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft

                [void]$FilterStaticPanel.Controls.Add( $Label )
            
            $AddFilter = New-Object System.Windows.Forms.Button
                $AddFilter.Text = "Add"
                $AddFilter.Dock = [System.Windows.Forms.DockStyle]::Left
                $AddFilter.Height = 21
                $AddFilter.Width  = 50
                
                [void]$FilterStaticPanel.Controls.Add( $AddFilter )

                Add-Member -InputObject $AddFilter -MemberType NoteProperty -Name FilterTab -Value $FilterTab

                Add-Member -InputObject $AddFilter -MemberType NoteProperty -Name OptionsPanel -Value $OptionsFlowPanel

                [void]$AddFilter.Add_Click({
                    if ($this.FilterTab.FilterFields.Count -eq $this.FilterTab.Fields.Count -1) {
                        [System.Windows.Forms.MessageBox]::Show(
                            "Maximum number of group/sort levels reached!",
                            "View Settings",
                            [System.Windows.Forms.MessageBoxButtons]::OK,
                            [System.Windows.Forms.MessageBoxIcon]::Exclamation
                        )
                        return
                    }

                    Write-Debug ("Registered Fields:`r`n{0}" -f ($this.FilterTab.FilterFields | Format-Table | Out-String))
                    $this.FilterTab.Handler.Reset()

                    # Find first unregistered field name
                    $filtered = New-Object System.Collections.ArrayList
                    [void]$filtered.AddRange($this.FilterTab.Fields)
                    foreach ($registration in $this.FilterTab.FilterFields) {
                        [void]$filtered.Remove($registration.Name)
                    }
                    $available = $filtered[0]
        
                    Write-Debug ("Avaliable Field:`r`n{0}" -f $available)

                    $RemoveButton = New-Object System.Windows.Forms.Button
                    $RemoveButton.Dock = [System.Windows.Forms.DockStyle]::Left
                    $RemoveButton.Text = "Remove"
                    $RemoveButton.Height = 21
                    $RemoveButton.Width = 75
                    [void]$RemoveButton.Add_Click({
                        $this.Parent.Unregister()
                    })

                    $SettingLabel = New-Object System.Windows.Forms.Label
                    $SettingLabel.Dock = [System.Windows.Forms.DockStyle]::Left
                    $SettingLabel.Width = 60
                    $SettingLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
                    $SettingLabel.Text = "Filter by:"

                    $FieldSelectBox = New-Object System.Windows.Forms.ComboBox
                    $FieldSelectBox.Dock = [System.Windows.Forms.DockStyle]::Left
                    $FieldSelectBox.DataSource = @($this.FilterTab.Fields)
                    $FieldSelectBox.Width = 85
                    
                    Add-Member -InputObject $FieldSelectBox -MemberType NoteProperty -Name FilterTab -Value $this.FilterTab

                    $FilterType = New-Object System.Windows.Forms.ComboBox
                    $FilterType.Dock = [System.Windows.Forms.DockStyle]::Left
                    $FilterType.DataSource = @("Include", "Exclude")
                    $FilterType.Width = 80

                    # Register Changes
                    $FilterType.Add_SelectedValueChanged({
                        $this.Parent.Registration.FilterType = $this.SelectedValue
                        $this.FilterTab.Handler.Parsed = $false
                    })

                    Add-Member -InputObject $FilterType -MemberType NoteProperty -Name FilterTab -Value $this.FilterTab

                    $FilterValue = New-Object System.Windows.Forms.TextBox
                    $FilterValue.Dock = [System.Windows.Forms.DockStyle]::Bottom
                    $FilterValue.Width = $RemoveButton.Width + $FilterType.Width + $SettingLabel.Width + $FieldSelectBox.Width
                    $FilterValue.Add_GotFocus({
                        $this.Previous = $this.Text
                    })

                    # Register Changes
                    $FilterValue.Add_LostFocus({
                        if ($this.Previous -ne $this.Text) {
                            $this.Parent.Registration.FilterValue = $this.Text
                            $this.FilterTab.Handler.Parsed = $false
                        }
                    })
                    Add-Member -InputObject $FilterValue -MemberType NoteProperty -Name Previous  -Value $null
                    Add-Member -InputObject $FilterValue -MemberType NoteProperty -Name FilterTab -Value $this.FilterTab

                    $SettingPanel = New-Object System.Windows.Forms.Panel
                    $SettingPanel.Height = 44
                    $SettingPanel.Width = $RemoveButton.Width + $FilterType.Width + $SettingLabel.Width + $FieldSelectBox.Width

                    [void]$SettingPanel.Controls.Add($FilterType)
                    [void]$SettingPanel.Controls.Add($FieldSelectBox)
                    [void]$SettingPanel.Controls.Add($SettingLabel)
                    [void]$SettingPanel.Controls.Add($RemoveButton)
                    [void]$SettingPanel.Controls.Add($FilterValue)

                    $registration = [PSCustomObject]@{
                        Name = $available
                        FilterType = "Include"
                        FilterValue = [String]::Empty
                        Setting = $SettingPanel
                    }

                    Add-Member -InputObject $SettingPanel -MemberType NoteProperty -Name FilterTab -Value $this.FilterTab

                    Add-Member -InputObject $SettingPanel -MemberType NoteProperty -Name Registration -Value $registration

                    Add-Member -InputObject $SettingPanel -MemberType NoteProperty -Name SettingType -Value "Filter"

                    Add-Member -InputObject $SettingPanel -MemberType ScriptMethod -Name Unregister -Value {
                        [void]$this.Parent.Controls.Remove($this)
                        $this.FilterTab.FilterFields.Remove($this.Registration)
                        Write-Debug ("Unregistered: {0}`t {1}" -f $this.Registration.Name, $this.GetHashCode())
                    }

                    # Register the setting
                    [void]$this.FilterTab.FilterFields.Add($registration)

                    Write-Debug ("Field Registered: {0}`t{1}" -f $registration.Name, $registration.Setting.GetHashCode())

                    [void]$this.OptionsPanel.Controls.Add($SettingPanel)
        
                    # After registering the field set the combo box selected value
                    $FieldSelectBox.SelectedItem = $available

                    # Adding the event handler last to prevent issues with the initial loading of the setting panel
                    $FieldSelectBox.Add_SelectedValueChanged({
                        $Current = $this.Parent # Setting Panel

                        $this.FilterTab.Handler.Parsed = $false

                        $registered = $this.FilterTab.FilterFields.ToArray()
                        foreach ($field in $registered) {
                            if (!$field.Setting.Equals($Current) -and $field.Name -eq $this.SelectedValue) {
                                $field.Setting.Unregister()
                            }
                        }

                        $Current.Registration.Name = $this.SelectedValue
                        Write-Debug ("Field Registered: {0}`t{1}" -f $Current.Registration.Name, $Current.GetHashCode())
                    })
                }) # End Button.Add_Click
        }

        # Add the TreeView Settings Panel
        [void]$FilterTab.Controls.Add( $FilterPanel )
    }

    Add-Member -InputObject $Container -MemberType NoteProperty -Name SettingsTab -Value $SettingsTab

    Add-Member -InputObject $Container -MemberType NoteProperty -Name FilterTab -Value $FilterTab

    Add-Member -InputObject $Container -MemberType NoteProperty -Name Display -Value $TreeViewControl

    return $Container
}