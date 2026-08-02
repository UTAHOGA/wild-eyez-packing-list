$docx = 'C:\Users\tyler\GitHub\wild-eyez-packing-list\working-files\PRIMARY_DESIGN_SOURCE_USE_THIS.docx'

$left = 47.5
$top = 538
$width = 517.3
$height = 94

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open($docx)

foreach ($shape in $doc.Shapes) {
  try {
    if ($shape.Name -eq 'priority_1_safety_editable_background') {
      $shape.Left = $left
      $shape.Top = $top
      $shape.Width = $width
      $shape.Height = $height
      $shape.ZOrder(0)
    }
    if ($shape.Name -eq 'priority_1_safety_editable_left_border') {
      $shape.Left = $left
      $shape.Top = $top
      $shape.Width = 6
      $shape.Height = $height
      $shape.ZOrder(0)
    }
    if ($shape.Name -eq 'priority_1_safety_editable_text') {
      $shape.Left = $left + 18
      $shape.Top = $top + 12
      $shape.Width = $width - 32
      $shape.Height = $height - 20
      $shape.ZOrder(0)
    }
  } catch {}
}

$doc.Save()
$doc.Close($false)
$word.Quit()

Write-Output "Moved page 2 editable priority safety box into final position."
