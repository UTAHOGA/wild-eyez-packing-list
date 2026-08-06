$ErrorActionPreference = "Stop"

$repo = "C:\Users\tyler\GitHub\wild-eyez-packing-list"
$baseline = "C:\DOCUMENTS . DOWNLOADS  . MUSIC\DOWNLOADS\WILD EYEZ PACKING AND PREPARATION GUIDE.pptx"
$scriptPath = Join-Path $repo "ADOBE_PRESENTATION_SCRIPT_EQUIPMENT_PREP_PACKET.txt"
$out = Join-Path $repo "WILD EYEZ PACKING AND PREPARATION GUIDE_PRESENTATION_WITH_NOTES.pptx"

if (!(Test-Path -LiteralPath $baseline)) {
  throw "Missing baseline PowerPoint: $baseline"
}
if (!(Test-Path -LiteralPath $scriptPath)) {
  throw "Missing Adobe presentation script: $scriptPath"
}

$script = Get-Content -LiteralPath $scriptPath -Raw
$fullStart = $script.IndexOf("FULL PRESENTATION SCRIPT")
if ($fullStart -lt 0) {
  throw "Could not find FULL PRESENTATION SCRIPT in $scriptPath"
}

$presentationScript = $script.Substring($fullStart)
$pageMatches = [regex]::Matches($presentationScript, "(?ms)^PAGE\s+(\d+)\s+-\s+(.+?)(?=^PAGE\s+\d+\s+-|\z)")
if ($pageMatches.Count -lt 9) {
  throw "Expected at least 9 page script sections, found $($pageMatches.Count)"
}

$notesByPage = @{}
foreach ($match in $pageMatches) {
  $pageNumber = [int]$match.Groups[1].Value
  if ($pageNumber -ge 1 -and $pageNumber -le 9) {
    $notesByPage[$pageNumber] = $match.Value.Trim() + "`r`n`r`n[Sources]`r`nWild Eyez Outfitters baseline PowerPoint and Adobe presentation script."
  }
}

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$pres = $null

try {
  $pres = $ppt.Presentations.Open($baseline, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue)

  if ($pres.Slides.Count -lt 9) {
    throw "Baseline PowerPoint has only $($pres.Slides.Count) slides."
  }

  for ($i = 1; $i -le 9; $i++) {
    $slide = $pres.Slides.Item($i)
    $notes = $notesByPage[$i]
    if ([string]::IsNullOrWhiteSpace($notes)) {
      throw "Missing notes for page $i"
    }

    try {
      $notesBody = $slide.NotesPage.Shapes.Placeholders(2)
      $notesBody.TextFrame.TextRange.Text = $notes
    } catch {
      $noteBox = $slide.NotesPage.Shapes.AddTextbox(1, 60, 120, 500, 420)
      $noteBox.TextFrame.TextRange.Text = $notes
    }
  }

  if (Test-Path -LiteralPath $out) {
    Remove-Item -LiteralPath $out -Force
  }
  $pres.SaveAs($out)
}
finally {
  if ($pres -ne $null) {
    $pres.Close()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
  }
  $ppt.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}

Write-Host $out
