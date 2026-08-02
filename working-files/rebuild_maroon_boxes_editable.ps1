$docx = 'C:\Users\tyler\GitHub\wild-eyez-packing-list\working-files\PRIMARY_DESIGN_SOURCE_USE_THIS.docx'

function RgbInt($r, $g, $b) {
  return [int]($r + (256 * $g) + (65536 * $b))
}

$maroon = RgbInt 126 37 48
$pink = RgbInt 245 223 227
$paper = RgbInt 248 243 234

function Get-AnchorForPage {
  param($doc, $pageNumber)
  foreach ($p in $doc.Paragraphs) {
    try {
      if ($p.Range.Information(3) -eq $pageNumber) {
        return $p.Range
      }
    } catch {}
  }
  return $doc.Paragraphs.Item(1).Range
}

function Add-EditableMaroonBox {
  param(
    $doc,
    [int]$pageNumber,
    [double]$left,
    [double]$top,
    [double]$width,
    [double]$height,
    [string]$title,
    [string]$body,
    [string]$namePrefix
  )

  $anchor = Get-AnchorForPage $doc $pageNumber

  # Cover the flattened/artwork version and become the real editable box background.
  $bg = $doc.Shapes.AddShape(1, $left, $top, $width, $height, $anchor)
  $bg.Name = "${namePrefix}_editable_background"
  $bg.RelativeHorizontalPosition = 1
  $bg.RelativeVerticalPosition = 1
  $bg.WrapFormat.Type = 3
  $bg.Fill.Visible = -1
  $bg.Fill.ForeColor.RGB = $pink
  $bg.Line.Visible = -1
  $bg.Line.ForeColor.RGB = $maroon
  $bg.Line.Weight = 0.75
  $bg.ZOrder(0)

  # Thick left border as its own editable shape.
  $bar = $doc.Shapes.AddShape(1, $left, $top, 6, $height, $anchor)
  $bar.Name = "${namePrefix}_editable_left_border"
  $bar.RelativeHorizontalPosition = 1
  $bar.RelativeVerticalPosition = 1
  $bar.WrapFormat.Type = 3
  $bar.Fill.Visible = -1
  $bar.Fill.ForeColor.RGB = $maroon
  $bar.Line.Visible = 0
  $bar.ZOrder(0)

  # Live editable text.
  $tb = $doc.Shapes.AddTextbox(1, $left + 18, $top + 12, $width - 32, $height - 20, $anchor)
  $tb.Name = "${namePrefix}_editable_text"
  $tb.RelativeHorizontalPosition = 1
  $tb.RelativeVerticalPosition = 1
  $tb.WrapFormat.Type = 3
  $tb.Fill.Visible = 0
  $tb.Line.Visible = 0
  $tb.TextFrame.MarginLeft = 0
  $tb.TextFrame.MarginRight = 0
  $tb.TextFrame.MarginTop = 0
  $tb.TextFrame.MarginBottom = 0
  $tb.TextFrame.WordWrap = $true
  $tb.TextFrame.TextRange.Text = "$title`r$body"

  $range = $tb.TextFrame.TextRange
  $range.Font.Name = 'Georgia'
  $range.Font.Size = 10
  $range.Font.Bold = $false
  $range.Font.Color = 0
  $range.ParagraphFormat.SpaceBefore = 0
  $range.ParagraphFormat.SpaceAfter = 0

  $titleRange = $range.Duplicate
  $titleRange.End = $titleRange.Start + $title.Length
  $titleRange.Font.Name = 'Lucida Sans'
  $titleRange.Font.Size = 12
  $titleRange.Font.Bold = $true
  $titleRange.Font.Color = $maroon

  $tb.ZOrder(0)
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open($docx)

Add-EditableMaroonBox `
  -doc $doc `
  -pageNumber 2 `
  -left 52 `
  -top 303 `
  -width 520 `
  -height 66 `
  -title 'PRIORITY #1 - YOUR SAFETY' `
  -body 'Conditions may become extreme in our hunting environment. Know your own personal limitations and follow all health practitioner advice that has been given to you above what your hunting guide may suggest or expect of you. Big game hunting is an extreme, adventurous sport. It is demanding both physically and mentally. Pursuing and taking of a big game animal is an intense and rigorous activity.' `
  -namePrefix 'priority_1_safety'

Add-EditableMaroonBox `
  -doc $doc `
  -pageNumber 6 `
  -left 52 `
  -top 295 `
  -width 520 `
  -height 112 `
  -title 'PRIORITY SAFETY POINT' `
  -body 'As hunting guides, we will often lead from the front or hunt from the front while locating game, evaluating terrain, and positioning the hunter for a possible shot. Because a guide may be forward of the hunter, Wild Eyez Outfitters applies a safety policy that goes above and beyond state law, rules, and regulations: no live round is to be chambered into the barrel until the target animal has been positively identified, the decision to take the shot has been made, and the shooting lane to the front is entirely clear.' `
  -namePrefix 'priority_safety_point'

$doc.Save()
$doc.Close($false)
$word.Quit()

Write-Output "Editable maroon safety boxes rebuilt in $docx"
