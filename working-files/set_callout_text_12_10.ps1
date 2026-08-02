$docx = 'C:\Users\tyler\GitHub\wild-eyez-packing-list\working-files\PRIMARY_DESIGN_SOURCE_USE_THIS.docx'

$targetTitles = @(
  'HOW TO USE THIS GUIDE',
  'PROFESSIONAL POINTER',
  'BE READY WHEN THE OPPORTUNITY CHANGES',
  'PRO TIP - PROTECT YOUR FEET',
  'AMMUNITION STANDARD',
  'MOST COMMON MISTAKE'
)

function Get-Text {
  param($shape)
  try {
    if ($shape.TextFrame.HasText -ne 0) {
      return $shape.TextFrame.TextRange.Text.Trim()
    }
  } catch {}
  return ''
}

function Is-TargetText {
  param($text)
  if ([string]::IsNullOrWhiteSpace($text)) { return $false }
  $upper = $text.ToUpper()
  foreach ($title in $targetTitles) {
    if ($upper.StartsWith($title)) { return $true }
  }
  return $false
}

function Set-Text-12-10 {
  param($shape)
  $text = Get-Text $shape
  if (-not (Is-TargetText $text)) { return }

  try {
    $shape.TextFrame.WordWrap = $true
    $shape.TextFrame.MarginLeft = 7
    $shape.TextFrame.MarginRight = 7
    $shape.TextFrame.MarginTop = 5
    $shape.TextFrame.MarginBottom = 5

    $range = $shape.TextFrame.TextRange
    $range.Font.Name = 'Georgia'
    $range.Font.Size = 10
    $range.Font.Bold = $false
    $range.ParagraphFormat.SpaceBefore = 0
    $range.ParagraphFormat.SpaceAfter = 0

    $firstBreak = $range.Text.IndexOf("`r")
    if ($firstBreak -gt 0) {
      $titleRange = $range.Duplicate
      $titleRange.End = $titleRange.Start + $firstBreak
      $titleRange.Font.Name = 'Lucida Sans'
      $titleRange.Font.Size = 12
      $titleRange.Font.Bold = $true

      $bodyRange = $range.Duplicate
      $bodyRange.Start = $titleRange.End + 1
      $bodyRange.Font.Name = 'Georgia'
      $bodyRange.Font.Size = 10
      $bodyRange.Font.Bold = $false
    }
  } catch {}
}

function Walk-Shape {
  param($shape)
  Set-Text-12-10 $shape
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

Write-Output "Editable callout text set to 12 pt headers and 10 pt body in $docx"
