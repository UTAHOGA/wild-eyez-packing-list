$docx = 'C:\Users\tyler\GitHub\wild-eyez-packing-list\working-files\PRIMARY_DESIGN_SOURCE_USE_THIS.docx'

# Main page grid used by the existing checklist/header elements.
$standardLeft = 47.5
$standardWidth = 517.3

$targetTitles = @(
  'HOW TO USE THIS GUIDE',
  'PROFESSIONAL POINTER',
  'BE READY WHEN THE OPPORTUNITY CHANGES',
  'PRO TIP - PROTECT YOUR FEET',
  'AMMUNITION STANDARD',
  'PRIORITY SAFETY POINT',
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

function Group-ContainsTarget {
  param($shape)
  try {
    if ($shape.Type -ne 6) { return $false }
    for ($i = 1; $i -le $shape.GroupItems.Count; $i++) {
      $child = $shape.GroupItems.Item($i)
      if (Is-TargetText (Get-Text $child)) { return $true }
    }
  } catch {}
  return $false
}

function Normalize-CalloutText {
  param($shape)
  $text = Get-Text $shape
  if (-not (Is-TargetText $text)) { return }

  try {
    $shape.TextFrame.WordWrap = $true
    $shape.TextFrame.MarginLeft = 7
    $shape.TextFrame.MarginRight = 7
    $shape.TextFrame.MarginTop = 4
    $shape.TextFrame.MarginBottom = 4
  } catch {}

  try {
    $range = $shape.TextFrame.TextRange
    $range.Font.Name = 'Georgia'
    $range.Font.Size = 7.3
    $range.Font.Bold = $false
    $range.ParagraphFormat.SpaceBefore = 0
    $range.ParagraphFormat.SpaceAfter = 0

    $firstBreak = $range.Text.IndexOf("`r")
    if ($firstBreak -gt 0) {
      $titleRange = $range.Duplicate
      $titleRange.End = $titleRange.Start + $firstBreak
      $titleRange.Font.Name = 'Lucida Sans'
      $titleRange.Font.Size = 8.8
      $titleRange.Font.Bold = $true

      $bodyRange = $range.Duplicate
      $bodyRange.Start = $titleRange.End + 1
      $bodyRange.Font.Name = 'Georgia'
      $bodyRange.Font.Size = 7.3
      $bodyRange.Font.Bold = $false
    }
  } catch {}
}

function Align-StandaloneTextBox {
  param($shape)
  $text = Get-Text $shape
  if (-not (Is-TargetText $text)) { return }

  try {
    $shape.Left = $standardLeft
    $shape.Width = $standardWidth
  } catch {}
  Normalize-CalloutText $shape
}

function Align-Group {
  param($shape)
  if (-not (Group-ContainsTarget $shape)) { return }

  try {
    $shape.LockAspectRatio = 0
    $shape.Left = $standardLeft
    $shape.Width = $standardWidth
  } catch {}

  try {
    for ($i = 1; $i -le $shape.GroupItems.Count; $i++) {
      $child = $shape.GroupItems.Item($i)
      Normalize-CalloutText $child
    }
  } catch {}
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open($docx)

for ($i = 1; $i -le $doc.Shapes.Count; $i++) {
  $shape = $doc.Shapes.Item($i)
  if ($shape.Type -eq 6) {
    Align-Group $shape
  } else {
    Align-StandaloneTextBox $shape
  }
}

$doc.Save()
$doc.Close($false)
$word.Quit()

Write-Output "Callout boxes aligned to left $standardLeft and width $standardWidth in $docx"
