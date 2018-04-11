Add-Type -AssemblyName System.Windows.Forms

###############################################################################
###############################################################################
## SECTION 01 ## PUBILC FUNCTIONS AND VARIABLES
##
## Pass-thru Export-ModuleMember calls export all functions and variables
## to the global session that were passed to this modules session from nested
## modules.
###############################################################################
###############################################################################

function Get-iTextAssembly () {
    $f = Get-Item (Join-Path $ModuleInvocationPath itextsharp.dll)
    return $f.FullName
}

function Convert-Font ([String]$pdfFont, [Single]$size) {
    switch -Regex ($pdfFont) {
        'Arial' {
            switch -Regex ($_) {
                Italic {
                    $font = New-Object System.Drawing.Font('Arial', $size, [System.Drawing.FontStyle]::Italic, [System.Drawing.GraphicsUnit]::Point)
                }

                Bold {
                    $font = New-Object System.Drawing.Font('Arial', $size, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point)
                }

                Default {
                    $font = New-Object System.Drawing.Font('Arial', $size, [System.Drawing.GraphicsUnit]::Point)
                }
            }
            break
        }

        'TimesNewRoman' {
            switch -Regex ($_) {
                Italic {
                    $font = New-Object System.Drawing.Font('Times New Roman', $size, [System.Drawing.FontStyle]::Italic, [System.Drawing.GraphicsUnit]::Point)
                }

                Bold {
                    $font = New-Object System.Drawing.Font('Times New Roman', $size, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point)
                }

                Default {
                    $font = New-Object System.Drawing.Font('Times New Roman', $size, [System.Drawing.GraphicsUnit]::Point)
                }
            }
            break
        }

        Default {
            switch -Regex ($_) {
                Italic {
                    $font = New-Object System.Drawing.Font('Microsoft Sans Serif', $size, [System.Drawing.FontStyle]::Italic, [System.Drawing.GraphicsUnit]::Point)
                }

                Bold {
                    $font = New-Object System.Drawing.Font('Microsoft Sans Serif', $size, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point)
                }

                Default {
                    $font = New-Object System.Drawing.Font('Microsoft Sans Serif', $size, [System.Drawing.GraphicsUnit]::Point)
                }
            }
            
            Write-Warning (New-Object System.ArgumentException("Unknown Font: $pdfFont"))
        }
    }
    return $font
}

function Get-FieldFont ($document, $field) {
    $item = $document.AcroFields.GetFieldItem($field)
    $decode = New-Object iTextSharp.text.pdf.TextField($null, $null, $null)

    if ($item) {
        $document.AcroFields.DecodeGenericDictionary($item.GetMerged(0), $decode)
        $t = $decode.Font.FullFontName
        return (Convert-Font $t[0][3] $decode.FontSize)
    }
}

function Get-FieldWidth ($document, $field) {
    $ref, $null = $document.AcroFields.GetFieldPositions($field)

    if ($ref) {
        return $ref.position.Width
    }
}

function Get-FieldHeight ($document, $field) {
    $ref, $null = $document.AcroFields.GetFieldPositions($field)

    if ($ref) {
        return $ref.position.Height
    }
}

function Get-Field ($document, $field) {
    $ref, $null = $document.AcroFields.GetFieldPositions($field)
    $obj        = $document.AcroFields.GetField($field)
    if ($obj -and $ref) {
        $output = [PSCustomObject]@{
            Name     = $field
            Location = $ref
            Ref      = $obj
        }

        $get_value = {
            return $this.Ref
        }

        $set_value = {
            param([String]$string)
            $this.Ref = $string
            
        }

        Add-Member -InputObject $output -MemberType ScriptProperty -Name Value -Value $get_value $set_value

        return $output
    }

    $obj        = $document.AcroFields.GetFieldItem($field)
    if ($obj -and $ref) {
        $output = [PSCustomObject]@{
            Name     = $field
            Location = $ref
            Ref      = $obj
        }

        $get_value = {
            return $this.Ref.Data
        }

        $set_value = {
            param([String]$string)
            $this.Ref.Ref.Reader.AcroFields.SetField($this.Name, $string)
            $this.Ref.Data = $string
        }

        Add-Member -InputObject $output -MemberType ScriptProperty -Name Value -Value $get_value $set_value
        Add-Member -InputObject $output.Ref -MemberType NoteProperty -Name Data -Value $null

        return $output
    }
}

function Get-FieldList ($document) {
    [String[]]$fields = $document.AcroFields.Fields | Select-Object -Property Key | % {$_.Key}
    return $fields
}

function Get-FieldGraphics ($document, [Int]$page) {
    $graphics = New-Object System.Collections.ArrayList

    $psize = $document.GetPageSize($page)
    $ptop  = [Float]$psize.Top

    foreach ($field in (Get-FieldList $document)) {
        $f = Get-Field $document $field

        if ($f.Location.page -ne $page) {
            continue
        }

        $ctop = [Float]$f.Location.position.Top
        $g = [PSCustomObject]@{
            Name  = $f.Name
            Field = $f
            Brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::LightBlue)
            Rect  = [System.Drawing.Rectangle]::new($f.Location.position.Left, ($ptop - $ctop),
                                                    $f.Location.position.Width, $f.Location.position.Height)
        }

        Add-Member -InputObject $g -MemberType ScriptMethod -Name Paint -Value $Paint_FieldGraphic

        [void]$graphics.Add($g)
    }
    return $graphics
}

function Get-PageTextChunks ($document, [Int]$page) {
    $extractor = New-Object iTextSharp.text.pdf.parser.ChunkLocationTextExtractionStrategy
    $e = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($document, $page, $extractor)

    $psize = $document.GetPageSize($page)
    $ptop  = [Float]$psize.Top

    $chunks = New-Object System.Collections.ArrayList
    foreach ($chunk in $extractor.Chunks) {
        $ctop  = [Float]$chunk.Rect.Top
        $g = [PSCustomObject]@{
            Font  = Convert-Font -pdfFont $chunk.Font.FullFontName[0][3] -size ($chunk.Rect.Height - 2.25 )
            Brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::Black)
            Point = New-Object System.Drawing.PointF($chunk.Rect.Left, ($ptop - $ctop))
        }
        $c = [PSCustomObject]@{
            Chunk = $chunk
            Draw  = $g
        }
        Add-Member -InputObject $c -MemberType ScriptMethod -Name Paint -Value $Paint_TextChunk

        [void]$chunks.Add($c)
    }

    return $chunks
}

function Get-PageViewer ($document, [Int]$npage) {
    $chunks = Get-PageTextChunks $document $npage
    $graphics = Get-FieldGraphics $document $npage

    $psize = $document.GetPageSize(1)

    $page = New-Object System.Windows.Forms.Control
    $page.Height = [Float]$psize.Height
    $page.Width  = [Float]$psize.Width
    $page.BackColor = [System.Drawing.Color]::White
    $page.Add_Paint({Paint-PageViewer $this $_})

    $fields = New-Object System.Collections.ArrayList
    foreach ($g in $graphics) {
        $c = New-Object System.Windows.Forms.Label
        $c.BackColor = [System.Drawing.Color]::Transparent
        $c.Name   = $g.Name
        $c.Top    = $g.Rect.Top
        $c.Left   = $g.Rect.Left
        $c.Width  = $g.Rect.Width
        $c.Height = $g.Rect.Height
        $c.Add_Click({Field-OnClick $this $_})
        Add-Member -InputObject $c -MemberType ScriptMethod -Name ToggleHighlight -Value $Field_ToggleHighlight
        Add-Member -InputObject $c -MemberType NoteProperty -Name Ref             -Value $g.Field
        [void]$page.Controls.Add($c)
        [void]$fields.Add($c)
    }

    Add-Member -InputObject $page -MemberType NoteProperty -Name Chunks          -Value $chunks
    Add-Member -InputObject $page -MemberType NoteProperty -Name Graphics        -Value $graphics
    Add-Member -InputObject $page -MemberType NoteProperty -Name Fields          -Value $fields
    Add-Member -InputObject $page -MemberType NoteProperty -Name Highlights      -Value (New-Object System.Collections.ArrayList)
    Add-Member -InputObject $page -MemberType ScriptMethod -Name ClearHighlights -Value $Page_ClearHighlights

    return $page
}

function Get-Viewer ($document) {
    $viewer = New-Object System.Windows.Forms.Panel
    $viewer.AutoScroll = $true

    Add-Member -InputObject $viewer -MemberType NoteProperty -Name Fields -Value (New-Object System.Collections.ArrayList)
    Add-Member -InputObject $viewer -MemberType ScriptMethod -Name ClearHighlights -Value $Viewer_ClearHighlights

    $top    = 0
    for ($i = 1; $i -le $document.NumberOfPages; $i++) {
        $page = Get-PageViewer $document $i
        $page.Top = $top
        $top += $page.Height + 20

        [void]$viewer.Controls.Add($page)
              $viewer.Fields.AddRange($page.Fields)
    }

    return $viewer
}

function Repair-FormFields ($document) {
# https://stackoverflow.com/questions/22909979/itextsharp-acrofields-are-empty
    [iTextSharp.text.pdf.PdfDictionary]$root = $document.Catalog
    [iTextSharp.text.pdf.PdfDictionary]$form = $root.GetAsDict( [iTextSharp.text.pdf.PdfName]::ACROFORM )
    [iTextSharp.text.pdf.PdfArray]$fields    = $form.GetAsArray( [iTextSharp.text.pdf.PdfName]::FIELDS )

    [iTextSharp.text.pdf.PdfDictionary]$page = $null
    [iTextSharp.text.pdf.PdfArray]$annots = $null
    for ($i = 1; $i -le $document.NumberOfPages; $i++) {
        $page = $document.GetPageN($i)
        $annots = $page.GetAsArray( [iTextSharp.text.pdf.PdfName]::ANNOTS )
        for ($j = 0; $j -lt $annots.Size; $j++) {
            [void]$fields.Add($annots.GetAsIndirectObject($j))
        }
    }

    return $document

    #PdfStamper stamper = new PdfStamper(reader, new FileOutputStream(dest));
    #stamper.close();
    #reader.close();
}

Export-ModuleMember *

###############################################################################
###############################################################################
## SECTION 02 ## PRIVATE FUNCTIONS
##
## No function or variable in this section is exported unless done so by an
## explicit call to Export-ModuleMember
###############################################################################
###############################################################################

$ModuleInvocationPath = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Definition)

$Csharp = @"
using System;
using System.Drawing;
using System.Collections;
using System.Collections.Generic;
using iTextSharp;
using iTextSharp.text.pdf.parser;

namespace iTextSharp.text.pdf.parser
{
    // Helper struct class to store rectangle and text
    public class TextInfo
    {
        public iTextSharp.text.Rectangle Rect;
        public String Text;
        public iTextSharp.text.pdf.DocumentFont Font;

        public TextInfo(iTextSharp.text.Rectangle rect, String text, iTextSharp.text.pdf.DocumentFont font) {
            this.Rect = rect;
            this.Text = text;
            this.Font = font;
        }
    }

    public class ChunkLocationTextExtractionStrategy : LocationTextExtractionStrategy
    {
        // Coordinates
        public List<TextInfo> Chunks = new List<TextInfo>();

        // Automatically called for each chunk of text in the PDF
        public override void RenderText(TextRenderInfo renderInfo) {
            base.RenderText(renderInfo);

            // Get the bounding box for the text chunk
            var bottomLeft = renderInfo.GetDescentLine().GetStartPoint();
            var topRight   = renderInfo.GetAscentLine().GetEndPoint();

            // Create rectangle
            var rect = new iTextSharp.text.Rectangle(
                bottomLeft[Vector.I1],
                bottomLeft[Vector.I2],
                topRight[Vector.I1],
                topRight[Vector.I2]
                );

            // Add to point collection
            this.Chunks.Add(new TextInfo(rect, renderInfo.GetText(), renderInfo.GetFont()));
        }
    }
}
"@

Add-Type -TypeDefinition $Csharp -ReferencedAssemblies @('System.Drawing', (Get-iTextAssembly))

$Paint_FieldGraphic = [Scriptblock]{
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.PaintEventArgs]$e
    )
    $e.Graphics.FillRectangle($this.Brush, $this.Rect)
}

$Paint_TextChunk = [Scriptblock]{
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.PaintEventArgs]$e
    )
    $e.Graphics.DrawString($this.Chunk.Text, $this.Draw.Font, $this.Draw.Brush, $this.Draw.Point)
}

$Field_ToggleHighlight = [Scriptblock]{
    if ($this.BackColor -eq [System.Drawing.Color]::Transparent) {
        $this.BackColor = [System.Drawing.Color]::Yellow
        [void]$this.Parent.Highlights.Add($this)
        return
    }
    $this.BackColor = [System.Drawing.Color]::Transparent
    [void]$this.Parent.Highlights.Remove($this)
}

$Page_ClearHighlights = [Scriptblock]{
    foreach ($item in $this.Highlights.ToArray()) {
        $item.ToggleHighlight()
    }
}

$Viewer_ClearHighlights = [Scriptblock]{
    foreach ($page in $this.Controls) {
        $page.ClearHighlights()
    }
}

function Field-OnClick ($sender, $e) {
    $this.ToggleHighlight()
}

function Paint-PageViewer ($sender, $e) {
    $e.Graphics.DrawRectangle([System.Drawing.Pen]::new([System.Drawing.Color]::Black, 2),
                              0, 0, $this.Width, $this.Height)

    foreach ($g in $this.Graphics) {
        $g.Paint($e)
    }

    foreach ($chunk in $this.Chunks) {
        $chunk.Paint($e)
    }
}

function New-Checkbox ($field) {

}

function New-Combobox ($field) {

}

function New-Listbox ($field) {

}

function New-RadioButton ($field) {

}

function New-Textbox ($field) {

}

function New-Page () {

}