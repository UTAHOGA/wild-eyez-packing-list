$docx = 'C:\Users\tyler\GitHub\wild-eyez-packing-list\working-files\PRIMARY_DESIGN_SOURCE_USE_THIS.docx'

$left = 47.5
$top = 295
$width = 523.5
$height = 112

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open($docx)

$bg = $null
$bar = $null
$txt = $null

foreach ($shape in $doc.Shapes) {
  if ($shape.Name -eq 'priority_safety_point_editable_background') { $bg = $shape }
  if ($shape.Name -eq 'priority_safety_point_editable_left_border') { $bar = $shape }
  if ($shape.Name -eq 'priority_safety_point_editable_text') { $txt = $shape }
}

if ($bg -ne $null) {
  $bg.Left = $left
  $bg.Top = $top
  $bg.Width = $width
  $bg.Height = $height
  $bg.WrapFormat.Type = 3
  $bg.ZOrder(0)
}

if ($bar -ne $null) {
  $bar.Left = $left
  $bar.Top = $top
  $bar.Width = 6
  $bar.Height = $height
  $bar.WrapFormat.Type = 3
  $bar.ZOrder(0)
}

if ($txt -ne $null) {
  $txt.Left = $left + 18
  $txt.Top = $top + 12
  $txt.Width = $width - 32
  $txt.Height = $height - 20
  $txt.WrapFormat.Type = 3
  $txt.Fill.Visible = 0
  $txt.Line.Visible = 0

  $range = $txt.TextFrame.TextRange
  $range.Font.Name = 'Georgia'
  $range.Font.Size = 10
  $range.Font.Bold = $false
  $range.Font.Color = 0

  $firstBreak = $range.Text.IndexOf("`r")
  if ($firstBreak -gt 0) {
    $titleRange = $range.Duplicate
    $titleRange.End = $titleRange.Start + $firstBreak
    $titleRange.Font.Name = 'Lucida Sans'
    $titleRange.Font.Size = 12
    $titleRange.Font.Bold = $true
    $titleRange.Font.Color = 3157330
  }

  for ($i = 0; $i -lt 8; $i++) {
    $txt.ZOrder(0)
  }
}

$doc.Save()
$doc.Close($false)
$word.Quit()

Write-Output "Fixed page 6 maroon editable safety box stack and placement."
