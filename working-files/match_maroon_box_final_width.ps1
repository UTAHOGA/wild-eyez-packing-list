$docx = 'C:\Users\tyler\GitHub\wild-eyez-packing-list\working-files\PRIMARY_DESIGN_SOURCE_USE_THIS.docx'

$left = 47.5
$width = 523.5

function Fix-Box {
  param($doc, $prefix, $top, $height)
  foreach ($shape in $doc.Shapes) {
    try {
      if ($shape.Name -eq "${prefix}_editable_background") {
        $shape.Left = $left
        $shape.Top = $top
        $shape.Width = $width
        $shape.Height = $height
        $shape.ZOrder(0)
      }
      if ($shape.Name -eq "${prefix}_editable_left_border") {
        $shape.Left = $left
        $shape.Top = $top
        $shape.Width = 6
        $shape.Height = $height
        $shape.ZOrder(0)
      }
      if ($shape.Name -eq "${prefix}_editable_text") {
        $shape.Left = $left + 18
        $shape.Top = $top + 12
        $shape.Width = $width - 32
        $shape.Height = $height - 20
        $shape.ZOrder(0)
      }
    } catch {}
  }
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open($docx)

Fix-Box $doc 'priority_1_safety' 538 94
Fix-Box $doc 'priority_safety_point' 295 112

$doc.Save()
$doc.Close($false)
$word.Quit()

Write-Output "Matched editable maroon box widths and covered old flattened edges."
