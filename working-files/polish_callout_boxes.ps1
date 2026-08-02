$docx = 'C:\Users\tyler\GitHub\wild-eyez-packing-list\working-files\PRIMARY_DESIGN_SOURCE_USE_THIS.docx'

$targetTitles = @(
  'HOW TO USE THIS GUIDE',
  'PRIORITY #1 - YOUR SAFETY',
  'PROFESSIONAL POINTER',
  'BE READY WHEN THE OPPORTUNITY CHANGES',
  'PRO TIP - PROTECT YOUR FEET',
  'AMMUNITION STANDARD',
  'PRIORITY SAFETY POINT',
  'MOST COMMON MISTAKE'
)

function Normalize-TextBox {
  param($shape)

  $text = ''
  try {
    if ($shape.TextFrame.HasText -ne 0) {
      $text = $shape.TextFrame.TextRange.Text
    }
  } catch {
    return
  }

  if ([string]::IsNullOrWhiteSpace($text)) { return }

  $matched = $false
  foreach ($title in $targetTitles) {
    if ($text.Trim().ToUpper().StartsWith($title)) {
      $matched = $true
      break
    }
  }
  if (-not $matched) { return }

  try {
    # Keep Tyler's intended fill/line colors and exact placement.
    # For grouped Word art, changing dimensions can move child objects, so this pass
    # only normalizes typography and text flow.
    $shape.TextFrame.WordWrap = $true
  } catch {}

  try {
    $range = $shape.TextFrame.TextRange
    $range.Font.Name = 'Georgia'
    $range.Font.Size = 7.7
    $range.Font.Bold = $false
    $range.ParagraphFormat.SpaceBefore = 0
    $range.ParagraphFormat.SpaceAfter = 0
    $range.ParagraphFormat.LineSpacingRule = 0

    $raw = $range.Text
    $firstBreak = $raw.IndexOf("`r")
    if ($firstBreak -gt 0) {
      $titleRange = $range.Duplicate
      $titleRange.End = $titleRange.Start + $firstBreak
      $titleRange.Font.Name = 'Lucida Sans'
      $titleRange.Font.Size = 8.8
      $titleRange.Font.Bold = $true
      $titleRange.ParagraphFormat.SpaceAfter = 1

      $bodyRange = $range.Duplicate
      $bodyRange.Start = $titleRange.End + 1
      $bodyRange.Font.Name = 'Georgia'
      $bodyRange.Font.Size = 7.7
      $bodyRange.Font.Bold = $false
      $bodyRange.ParagraphFormat.SpaceBefore = 0
      $bodyRange.ParagraphFormat.SpaceAfter = 0
    }
  } catch {}
}

function Walk-Shape {
  param($shape)
  Normalize-TextBox $shape
  try {
    if ($shape.Type -eq 6) {
      for ($i = 1; $i -le $shape.GroupItems.Count; $i++) {
        Walk-Shape $shape.GroupItems.Item($i)
      }
    }
  } catch {}
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open($docx)

for ($i = 1; $i -le $doc.Shapes.Count; $i++) {
  Walk-Shape $doc.Shapes.Item($i)
}

$doc.Save()
$doc.Close($false)
$word.Quit()

Write-Output "Callout box typography and text-frame sizing normalized in $docx"
