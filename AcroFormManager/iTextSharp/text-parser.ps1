# Token Tags
$tag = @{
    BeginText = "BT"
    EndText   = "ET"
}

function Parse-Raw ($data) {
    $rbuffer = New-Object System.Text.StringBuilder
    $wbuffer = New-Object System.Text.StringBuilder
    $max = $data.Length
    $pos = 0
    while ($pos -ne $max) {
        $char = [System.Text.ASCIIEncoding]::UTF8.GetChars($data, $pos++, 1)

        while ($char -ne '[' -and $pos -lt $max) {
            [void]$rbuffer.Append($char)
            $char = [System.Text.ASCIIEncoding]::UTF8.GetChars($data, $pos++, 1)
        }
        [void]$wbuffer.Append($rbuffer.ToString())
        $rbuffer.Length = 0

        while ($char -ne ']' -and $pos -lt $max) {
            [void]$rbuffer.Append($char)
            $char = [System.Text.ASCIIEncoding]::UTF8.GetChars($data, $pos++, 1)
        }
        [void]$rbuffer.Append($char)
        [void]$wbuffer.AppendLine($rbuffer.ToString())
        $rbuffer.Length = 0
    }

    return $wbuffer.ToString()
}

#$text = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader, 1)