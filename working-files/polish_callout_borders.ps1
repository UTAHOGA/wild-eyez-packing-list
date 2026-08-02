$docx = 'C:\Users\tyler\GitHub\wild-eyez-packing-list\working-files\PRIMARY_DESIGN_SOURCE_USE_THIS.docx'

$maroon = 3157330   # Word BGR-ish COM color close to burgundy/maroon.
$bronze = 5540532   # Warm bronze/tan border color.

function Get-Text {
  param($shape)
  try {
    if ($shape.TextFrame.HasText -ne 0) {
      return $shape.TextFrame.TextRange.Text.Trim()
    }
  } catch {}
  return ''
}

function Is-CalloutText {
  param($text)
  if ([string]::IsNullOrWhiteSpace($text)) { return $false }
  $upper = $text.ToUpper()
  return (
    $upper.StartsWith('HOW TO USE THIS GUIDE') -or
    $upper.StartsWith('PROFESSIONAL POINTER') -or
    $upper.StartsWith('BE READY WHEN THE OPPORTUNITY CHANGES') -or
    $upper.StartsWith('PRO TIP - PROTECT YOUR FEET') -or
    $upper.StartsWith('AMMUNITION STANDARD') -or
    $upper.StartsWith('PRIORITY SAFETY POINT') -or
    $upper.StartsWith('MOST COMMON MISTAKE')
  )
}

function Border-Color-ForText {
  param($text)
  $upper = $text.ToUpper()
  if ($upper.StartsWith('BE READY') -or $upper.StartsWith('PRIORITY SAFETY') -or $upper.StartsWith('MOST COMMON')) {
    return $maroon
  }
  return $bronze
}

function Apply-OuterBorder {
  param($shape, $color)
  try {
    $shape.Line.Visible = -1
    $shape.Line.Weight = 0.75
    $shape.Line.ForeColor.RGB = $color
  } catch {}
}

function Walk-Shape {
  param($shape, $inTargetGroup, $groupColor)

  $text = Get-Text $shape
  $isTextTarget = Is-CalloutText $text

  if ($isTextTarget) {
    $color = Border-Color-ForText $text
    Apply-OuterBorder $shape $color
  }

  try {
    if ($shape.Type -eq 6) {
      $containsTarget = $false
      $color = $bronze
      for ($i = 1; $i -le $shape.GroupItems.Count; $i++) {
        $childText = Get-Text $shape.GroupItems.Item($i)
        if (Is-CalloutText $childText) {
          $containsTarget = $true
          $color = Border-Color-ForText $childText
          break
        }
      }
      for ($i = 1; $i -le $shape.GroupItems.Count; $i++) {
        Walk-Shape $shape.GroupItems.Item($i) $containsTarget $color
      }
    } elseif ($inTargetGroup -and $shape.Type -eq 5) {
      # This is an editable rectangle inside a callout group. Give the outer
      # editable rectangle the same border weight without changing its fill.
      Apply-OuterBorder $shape $groupColor
    }
  } catch {}
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open($docx)

for ($i = 1; $i -le $doc.Shapes.Count; $i++) {
  Walk-Shape $doc.Shapes.Item($i) $false $bronze
}

$doc.Save()
$doc.Close($false)
$word.Quit()

Write-Output "Callout outer border weights normalized in $docx"
