Add-Type -AssemblyName System.Windows.Forms
Import-Module iTextSharp

function Form-OnLoad ($sender, $e) {
    # DOCUMENT OPEN DIALOG
    $dialog = New-Object System.Windows.Forms.OpenFiledialog
    $dialog.Title = "Open DD2875, System Access Authorization Request"
    $dialog.Multiselect = $false
    
    # Currently only CSV flat databases are supported.
    $dialog.Filter = "DD2875 (*.pdf)|*.pdf"
    $dialog.FilterIndex = 1

    if ($dialog.Showdialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $fpath = $dialog.FileName
    }
    else {
        return
    }

    $this.Reader = New-Object iTextSharp.text.pdf.PdfReader -ArgumentList $fpath
    $this.TreeView.Reader = $this.Reader

    $pdfViewer = Get-Viewer $this.Reader
    $pdfViewer.Dock = [System.Windows.Forms.DockStyle]::Fill

    $form.Viewer     = $pdfViewer
    $treeview.Viewer = $pdfViewer

    Set-TreeViewContent $treeview (Split-Path $fpath -Leaf) $pdfViewer.Fields

    [void]$layout.Panel2.Controls.Add($pdfViewer)
}

function Set-TreeViewContent ($treeview, $file, $controls) {
    $dnode = New-Object System.Windows.Forms.TreeNode
    $dnode.Text = $file
    $dnode.ContextMenuStrip = $DocumentNode_ContextMenuStrip
    $dnode.Tag  = [PSCustomObject]@{
        Type = "Document"
    }
    Add-Member -InputObject $dnode -MemberType ScriptMethod -Name Accepts -Value $DocumentNode_Accepts
    [void]$treeview.Nodes.Add($dnode)
    $treeview.DocumentNode = $dnode

    foreach ($ctrl in $controls) {
        $node = New-Object System.Windows.Forms.TreeNode
        $node.Text = $ctrl.Name
        $node.Tag  = [PSCustomObject]@{
            Type  = "Field"
            Field = $ctrl
        }

        # Set TreeNode select callback
        $ctrl.Add_Click($FieldItem_Select)
        Add-Member -InputObject $ctrl -MemberType NoteProperty -Name Node -Value $node

        # Set Viewer field value popup
        $tip = New-Object System.Windows.Forms.ToolTip
        $tip.SetToolTip($ctrl, $ctrl.Text)
        $tip.IsBalloon = $true
        Add-Member -InputObject $ctrl -MemberType NoteProperty -Name ToolTip -Value $tip

        # Node select handler to update viewer
        Add-Member -InputObject $node -MemberType ScriptMethod -Name ToggleHighlight -Value $FieldNode_ToggleHighlight

        # TreeView DragDrop support
        Add-Member -InputObject $node -MemberType ScriptMethod -Name Accepts -Value $FieldNode_Accepts
        [void]$dnode.Nodes.Add($node)
    }
}

function Add-ProcessNode ($sender, $e) {
    $Menu = $sender.GetCurrentParent()
    [System.Windows.Forms.TreeView] $TreeView = $Menu.SourceControl

    $n = New-Object System.Windows.Forms.TreeNode
    $n.Tag = [PSCustomObject]@{
        Type = 'Process'
    }
    Add-Member -InputObject $n -MemberType ScriptMethod -Name Accepts -Value $ProcessNode_Accepts

    $f = New-Object System.Windows.Forms.Form
    $f.KeyPreview = $true
    $f.Add_Closing($ProcessNodeDialog_AddHandler)
    $f.Add_KeyDown({ProcessNodeDialog-OnKeyDown $this $_})

    $l = New-Object System.Windows.Forms.Label
    $l.Text = "Process Name"
    $l.Dock = [System.Windows.Forms.DockStyle]::Left

    $t = New-Object System.Windows.Forms.TextBox
    $t.Dock = [System.Windows.Forms.DockStyle]::Fill

    $f.Controls.AddRange(@($t, $l))
    
    Add-Member -InputObject $f -MemberType NoteProperty -Name Node  -Value $n
    Add-Member -InputObject $f -MemberType NoteProperty -Name Input -Value $t

    [void]$f.ShowDialog($TreeView.Form)

    # Insert Process Node
    if (![string]::IsNullOrEmpty($n.Text)) {
        $i = $TreeView.ProcessNodes.Add($n)

        # Get index of previous process node
        if ($i -gt 0) {
            $i = $TreeView.ProcessNodes[$i - 1].Index + 1
        }

        # Insert node after previous process node
        [void]$TreeView.DocumentNode.Nodes.Insert($i, $n)
    }
}

$DocumentNode_Accepts = [Scriptblock]{
    param($node)
    Switch ($node.Tag.Type) {
        'Field'   {return $true }
        'Process' {return $true }
        Default   {return $false}
    }
}

$DocumentNode_ContextMenuStrip = & {
    $context = New-Object System.Windows.Forms.ContextMenuStrip
    [Void]$context.Items.Add( (New-Object System.Windows.Forms.ToolStripMenuItem("Add Process Node", $null, {Add-ProcessNode $this $_})) )

    return $context
}

$ProcessNodeDialog_AddHandler = [Scriptblock]{
    $this.Node.Text = $this.Input.Text.Trim()
}

function ProcessNodeDialog-OnKeyDown {
    param(
        [System.Windows.Forms.Form]
        $sender,

        [System.Windows.Forms.KeyEventArgs]
        $e
    )

    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $e.Handled = $true
        $sender.Close()
    }
    else {
        $e.Handled = $false
    }
}

$ProcessNode_Accepts = [Scriptblock]{
    param($node)
    Switch ($node.Tag.Type) {
        'Field'      {return $true }
        Default      {return $false}
    }
}

$FieldItem_Select  = [Scriptblock]{
    $this.Node.TreeView.SelectedNode = $this.Node
}

$FieldNode_Accepts = [Scriptblock]{
    param($node)
    Switch ($node.Tag.Type) {
        'Value'      {return $true }
        'Validation' {return $true }
        Default      {return $false}
    }
}

$FieldNode_ToggleHighlight = [Scriptblock]{
    $this.TreeView.Viewer.ClearHighlights()
    $this.Tag.Field.ToggleHighlight()
    $this.TreeView.Form.Layout.Panel2.ScrollControlIntoView($this.Tag.Field)
}

$ValueNode_Accepts = {
    param($node)
    Switch ($node.Tag.Type) {
        Default      {return $false}
    }
}

$ValidationNode_Accepts = {
    param($node)
    Switch ($node.Tag.Type) {
        Default      {return $false}
    }
}

function TreeView-OnClick {
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

function TreeView-AfterSelect ($sender, $e) {
    $node = $sender.SelectedNode
    # Toggle viewer highlighted fields
    if ($node.Tag.Type -eq 'Field') {
        $node.ToggleHighlight()
    }
}

function TreeView-ItemDrag ($sender, [System.Windows.Forms.ItemDragEventArgs]$e) {
    $sender.DoDragDrop($e.Item, [System.Windows.Forms.DragDropEffects]::Move)
}

function TreeView-DragEnter ($sender, [System.Windows.Forms.DragEventArgs]$e) {
    $e.Effect = [System.Windows.Forms.DragDropEffects]::Move
}

function TreeView-DragOver ($sender, [System.Windows.Forms.DragEventArgs]$e) {
    $e.Effect = [System.Windows.Forms.DragDropEffects]::Move
}

function TreeView-DragDrop ($sender, [System.Windows.Forms.DragEventArgs]$e) {
    # Retrieve the node at the drop location
    $point  = $sender.PointToClient( (New-Object System.Drawing.Point($e.X, $e.Y)) )
    $target = $sender.GetNodeAt($point)

    # Retrieve the node that was dragged
    $source = [System.Windows.Forms.TreeNode]$e.Data.GetData([System.Windows.Forms.TreeNode])

    # Confirm that the target and source nodes are not the same, and
    # that the target is not null (for example if you dragged outside
    # the TreeView control).
    if ($target -and !$source.Equals($target)) {
        
        # Validate that the target accepts children of this type
        if ($target.Accepts($source)) {
            # Remove source from it's current parent
            $source.Remove()

            # Add source node to it's new parent (target)
            [void]$target.Nodes.Add($source)
            $target.Expand()
        }
        elseif ($target.Parent -and $target.Parent.Accepts($source)) {
            # Remove source from it's current parent
            $source.Remove()

            # Insert source before target
            $target.Parent.Nodes.Insert($target.Index, $source)
        }
    }
}

function Start-Debugging {
    Write-Host Debugging...
    Write-Host Debugging...
    Write-Host Debugging...
    Write-Host Debugging...
    Write-Host Debugging...
}

$ReadFields_OnClick = [Scriptblock]{
    foreach ($ctrl in $this.Form.Viewer.Fields) {
        $ctrl.Text = $ctrl.Ref.Value
        $ctrl.ToolTip.SetToolTip($ctrl, $ctrl.Text)
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Width  = 300
$form.Height = 200
#$form.Add_Paint({Viewer-OnPaint $this $_})

$layout = New-Object System.Windows.Forms.SplitContainer
$layout.Dock = [System.Windows.Forms.DockStyle]::Fill
$layout.SplitterWidth = 5
$layout.Panel2.AutoScroll = $true
    [void]$form.Controls.Add($layout)

$treeview = New-Object System.Windows.Forms.TreeView
$treeview.Dock = [System.Windows.Forms.DockStyle]::Fill
$treeview.AllowDrop = $true
$treeview.Add_Click({TreeView-OnClick $this $_})
$treeview.Add_AfterSelect({TreeView-AfterSelect $this $_})
$treeview.Add_ItemDrag({TreeView-ItemDrag $this $_})
$treeview.Add_DragEnter({TreeView-DragEnter $this $_})
$treeview.Add_DragOver({TreeView-DragOver $this $_})
$treeview.Add_DragDrop({TreeView-DragDrop $this $_})
    [void]$layout.Panel1.Controls.Add($treeview)

$menu = @{}
$menu.Strip = New-Object System.Windows.Forms.MenuStrip
    $form.MainMenuStrip = $menu.Strip
    [void]$form.Controls.Add($menu.Strip)

$menu.Debug = New-Object System.Windows.Forms.ToolStripMenuItem('Debug', $null, {Start-Debugging})
[void]$menu.Strip.Items.Add( $menu.Debug )

$menu.ReadFields = New-Object System.Windows.Forms.ToolStripMenuItem('Read Fields', $null, $ReadFields_OnClick)
Add-Member -InputObject $menu.ReadFields -MemberType NoteProperty -Name Form -Value $form
[void]$menu.Strip.Items.Add( $menu.ReadFields)

Add-Member -InputObject $form -MemberType NoteProperty -Name Reader -Value $null
Add-Member -InputObject $form -MemberType NoteProperty -Name TreeView -Value $treeview
Add-Member -InputObject $form -MemberType NoteProperty -Name Viewer -Value $null
Add-Member -InputObject $form -MemberType NoteProperty -Name Layout -Value $layout
Add-Member -InputObject $treeview -MemberType NoteProperty -Name Form -Value $form
Add-Member -InputObject $treeview -MemberType NoteProperty -Name Viewer -Value $null
Add-Member -InputObject $treeview -MemberType NoteProperty -Name Reader   -Value $null
Add-Member -InputObject $treeview -MemberType NoteProperty -Name DocumentNode -Value $null
Add-Member -InputObject $treeview -MemberType NoteProperty -Name ProcessNodes -Value (New-Object System.Collections.ArrayList)

$form.Add_Load({Form-OnLoad $this $_})
[void]$form.ShowDialog()
      $form.Reader.Close()
      $form.Dispose()