###############################################################################
###############################################################################
## SECTION 01 ## PUBILC FUNCTIONS AND VARIABLES
##
## Pass-thru Export-ModuleMember calls export all functions and variables
## to the global session that were passed to this modules session from nested
## modules.
###############################################################################
###############################################################################
# FIELD FORMAT HELPERS
function Format-ShortDate {
    param(
        [Parameter(Mandatory = $false)]
            [DateTime]
            $date = [DateTime]::Today
    )
    return ("{0}{1:d2}{2:d2}" -f 
        $date.Year,
        $date.Month,
        $date.Day)
}

function Format-SignatureBlock {
    param(
        [Parameter(Mandatory = $true)]
            [String]
            $fname,
        
        [Parameter(Mandatory = $true)]
            [String]
            $mi,
        
        [Parameter(Mandatory = $true)]
            [String]
            $lname,
        
        [Parameter(Mandatory = $true)]
            [String]
            $rank,

        [Parameter(Mandatory = $true)]
        [ValidateSet("User",
                     "Supervisor",
                     "UnitSecurityManager",
                     "ProcessedBy",
                     "RevalidatedBy")]
            [String]
            $field,

        [Parameter(Mandatory = $true)]
            $document
    )

    $full = ("{0} {1}. {2}, {3}, USAF" -f
        $fname.ToUpper(),
        $mi.ToUpper(),
        $lname.ToUpper(),
        $rank)

    $truncated = ("{0} {1}. {2}" -f
        $fname.ToUpper(),
        $mi.ToUpper(),
        $lname.ToUpper())

    $font = Get-FieldFont $document $sigblocks[$field]
    $wmax = Get-FieldWidth $document $sigblocks[$field]

    if ($graphics.MeasureString($full, $font).Width -le $wmax) {
        return $full
    }
    else {
        return $truncated
    }
}

function Dialog-OpenForm {
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

    $reader   = New-Object iTextSharp.text.pdf.PdfReader -ArgumentList $fpath
    return $reader
}

function Scan-Folder ($path) {
    $dialog  = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description  = "Select Directory..."
    $dialog.SelectedPath = $erm

    if ([String]::IsNullOrEmpty($path)) {
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $path = $dialog.SelectedPath
        }
        else {
            return
        }
    }

    $dd2875s = Get-ChildItem $path -Filter *2875*.pdf -Recurse
    $report  = New-Object System.Collections.ArrayList
    
    foreach ($file in $dd2875s) {
        $reader = New-Object iTextSharp.text.pdf.PdfReader -ArgumentList $file.FullName
        $signed   = $reader.AcroFields.GetSignatureNames()

        $form = @{
            Path    = $file.FullName
            Package = (Split-Path $file.FullName -Parent).Split('\')[-1]
            Name    = $file.Name
        }

        switch -Regex ($reader.Info.Producer) {
            $AcroDistiller8 {
                foreach ($key in $signatures.GetEnumerator()) {
                    $form.Add($key.Name, $signed.Contains($key.Value))
                }
            }

            $AcroDesigner9 {
                foreach($key in $signatures.GetEnumerator()) {
                    $flag = $false
                    foreach ($item in $signed) {
                        if ($key.Value -match $item) {
                            $flag = $true
                            $signed.Remove($item)
                            break
                        }
                    }
                    $form.Add($key.Name, $flag)
                }
            }

            default {
                foreach ($key in $signatures.GetEnumerator()) {
                    $form.Add($key.Name, 'ERROR: Unknown form version')
                }
            }
        }

        [void]$report.Add([PSCustomObject]$form)
        $reader.Close()
    }

    return $report
}

###############################################################################
###############################################################################
## SECTION 02 ## PRIVATE FUNCTIONS AND VARIABLES
##
## No function or variable in this section is exported unless done so by an
## explicit call to Export-ModuleMember
###############################################################################
###############################################################################
# Assemblies
Import-Module iTextSharp

# MeasureString Graphics Object !-- Clean-Up Graphics.Form.Dispose() --!
$graphics = & {
    $f = New-Object System.Windows.Forms.Form
    $g = $f.CreateGraphics()
    
    Add-Member -InputObject $g -MemberType NoteProperty -Name Form -Value $f

    return $g
}

# Default Paths
$erm = "\\52tyfr-fs-001v.area52.afnoapps.usaf.mil\ERM\21 - (PA) Forms used to Formally Record Authorization for Access to Special Program Material\2- DD 2875\691 COS"

$AcroDistiller8 = "Acrobat Distiller 8.1.0"
$AcroDesigner9  = "Adobe LiveCycle Designer ES 9.0"

# FIELD KEYS
$signatures = @{
    User                        = "usersign"              # block 11
    Supervisor                  = "supvsign"              # block 18
    InformationOwner            = "ownersign"             # block 21
    InformationAssuranceOfficer = "iaosign"               # block 22
    UnitSecurityManager         = "sec_mgr_sign"          # block 31
    ProcessedBy                 = "procsign"              # part IV
    RevalidatedBy               = "revalsign"             # part IV
}

$request = @{
    Type                        = "xtype"                 # checkbox group
    UserID                      = "userid"                # textbox 
    Date                        = "reqdate"               # textbox
    System                      = "syst_name"             # textbox
    Location                    = "location"              # textbox
}

$requestor = @{
    Name                        = "name"                  # block 1   textbox
    Organization                = "reqorg"                # block 2   textbox
    Department                  = "reqsymb"               # block 3   textbox
    Phone                       = "reqphone"              # block 4   textbox
    Email                       = "reqemail"              # block 5   textbox
    Title                       = "reqtitle"              # block 6   textbox
    Address                     = "reqaddr"               # block 7   textbox
    Citizenship                 = "xcitizen"              # block 8   checkbox  group
    Designation                 = "xdesignation"          # block 9   textbox
    IAConfimration              = "xia"                   # block 10  checkbox
    IATrainingDate              = "trngdate"              # block 10  textbox
    SignatureBlock              = "user_name"             # block 11  textbox
    Signature                   = $signatures.User        # block 11  signature
    SignatureDate               = "userdate"              # block 12  textbox
}

$endorsement = @{
    Justification               = "justify"               # block 13  textbox   multiline
    AdditionalInformation       = "optinfo"               # block 27  textbox   multiline
    Authorized                  = "xauth"                 # block 14  checkbox
    Privileged                  = "xpriv"                 # block 14  checkbox
    Unclassified                = "xunclass"              # block 15  checkbox
    Classified                  = "xclass"                # block 15  checkbox
    Classification              = "classcat"              # block 15  textbox
    OtherConfirm                = "xotheracc"             # block 15  checkbox
    OtherDetail                 = "other_acc"             # block 15  textbox
    NeedToKnow                  = "xverif"                # block 16  checkbox
    Expiration                  = "expdate"               # block 16a textbox
    ExpirationDetail            = "acc_exp"               # block 16a textbox
}

$supervisor = @{
    SignatureBlock              = "supvname"              # block 17  textbox
    Signature                   = $signatures.Supervisor  # block 18  signature
    SignatureDate               = "supvdate"              # block 19  textbox
    Organization                = "supvorg"               # block 20  textbox
    Email                       = "supvemail"             # block 20a textbox
    Phone                       = "supvphone"             # block 20b textbox
}

$owner = @{
    Signature = $signatures.InformationOwner              # block 21  signature
    Phone                       = "ownerphone"            # block 21a textbox
    Date                        = "ownerdate"             # block 21b textbox
}

$iaofficer = @{
    Signature = $signatures.InformationAssuranceOfficer   # block 22  signature
    Organization                = "iaoorg"                # block 23  textbox
    Phone                       = "iaophone"              # block 24  textbox
    Date                        = "iaodate"               # block 25  textbox
}

$security = @{
    InvestigationType           = "typeinv"               # block 28  textbox
    InvestigationDate           = "invest_date"           # block 28a textbox
    Clearance                   = "clr_level"             # block 28b textbox
    ITDesignation               = "xit_lvl"               # block 28c checkbox  group
    Verifier                    = "verifname"             # block 29  textbox
    Phone                       = "sec_mgr_phone"         # block 30  textbox
    Signature = $signatures.UnitSecurityManager           # block 31  signature
    Date                        = "sec_mgr_date"          # block 32  textbox
}

$preparation = @{
    SystemTitle                 = "title1"                # part IV   textbox
    SystemName                  = "system"                # part IV   textbox
    SystemCode                  = "acctcode1"             # part IV   textbox
    DomainTitle                 = "title2"                # part IV   textbox
    DomainName                  = "domain"                # part IV   textbox
    DomainCode                  = "acctcode2"             # part IV   textbox
    ServerTitle                 = "title3"                # part IV   textbox
    ServerName                  = "server"                # part IV   textbox
    ServerCode                  = "acctcode3"             # part IV   textbox
    ApplicationTitle            = "title4"                # part IV   textbox
    ApplicationName             = "applic"                # part IV   textbox
    ApplicationCode             = "acctcode4"             # part IV   textbox
    DirectoryTitle              = "title5"                # part IV   textbox
    DirectoryName               = "direc"                 # part IV   textbox
    DirectoryCode               = "acctcode5"             # part IV   textbox
    FilesTitle                  = "title6"                # part IV   textbox
    FilesName                   = "files"                 # part IV   textbox
    FilesCode                   = "acctcode6"             # part IV   textbox
    DatasetsTitle               = "title7"                # part IV   textbox
    DatasetsName                = "datasets"              # part IV   textbox
    DatasetsCode                = "acctcode7"             # part IV   textbox
    DateProcessed               = "dateproc"              # part IV   textbox
    ProcessedBy                 = "procname"              # part IV   textbox
    ProcessedSign = $signatures.ProcessedBy               # part IV   signature
    ProcessedDate               = "procdate"              # part IV   textbox
    DateRevalidated             = "datereval"             # part IV   textbox
    RevalidatedBy               = "revalname"             # part IV   textbox
    RevalidatedSign = $signatures.RevalidatedBy           # part IV   signature
    RevalidatedDate             = "revaldate"             # part IV   textbox
}

$sigblocks = @{
    User                = $requestor.SignatureBlock
    Supervisor          = $supervisor.SignatureBlock
    UnitSecurityManager = $security.Verifier
    ProcessedBy         = $preparation.ProcessedBy
    RevalidatedBy       = $preparation.RevalidatedBy
}