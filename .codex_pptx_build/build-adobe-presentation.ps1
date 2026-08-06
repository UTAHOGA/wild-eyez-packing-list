$ErrorActionPreference = "Stop"

$repo = "C:\Users\tyler\GitHub\wild-eyez-packing-list"
$scriptPath = Join-Path $repo "ADOBE_PRESENTATION_SCRIPT_EQUIPMENT_PREP_PACKET.txt"
$pagesDir = Join-Path $repo "assets\pages"
$out = Join-Path $repo "WILD EYEZ EQUIPMENT PREPARATION PACKET ADOBE PRESENTATION.pptx"

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

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$pres = $ppt.Presentations.Add([Microsoft.Office.Core.MsoTriState]::msoTrue)

# Portrait letter proportions, in points.
$pres.PageSetup.SlideWidth = 612
$pres.PageSetup.SlideHeight = 792

try {
  foreach ($match in $pageMatches) {
    $pageNumber = [int]$match.Groups[1].Value
    if ($pageNumber -lt 1 -or $pageNumber -gt 9) { continue }

    $title = "Page $pageNumber - " + ($match.Groups[2].Value.Trim())
    $notes = $match.Value.Trim()

    $imagePath = Join-Path $pagesDir ("WILD EYEZ PACKING AND PREPARATION GUIDE6_Page_{0}.jpg" -f $pageNumber)
    if (!(Test-Path -LiteralPath $imagePath)) {
      throw "Missing page image: $imagePath"
    }

    $slide = $pres.Slides.Add($pres.Slides.Count + 1, 12) # ppLayoutBlank
    $slide.Shapes.AddPicture($imagePath, 0, -1, 0, 0, $pres.PageSetup.SlideWidth, $pres.PageSetup.SlideHeight) | Out-Null

    # Add a small hidden/off-canvas label so PowerPoint/Adobe can still identify the slide.
    $label = $slide.Shapes.AddTextbox(1, -500, -500, 400, 40)
    $label.TextFrame.TextRange.Text = $title

    # Speaker notes: placeholder 2 is normally the notes body on PowerPoint notes pages.
    try {
      $notesBody = $slide.NotesPage.Shapes.Placeholders(2)
      $notesBody.TextFrame.TextRange.Text = $notes + "`r`n`r`n[Sources]`r`nWild Eyez Outfitters current packet page image and Adobe presentation script."
    } catch {
      $noteBox = $slide.NotesPage.Shapes.AddTextbox(1, 60, 120, 500, 420)
      $noteBox.TextFrame.TextRange.Text = $notes + "`r`n`r`n[Sources]`r`nWild Eyez Outfitters current packet page image and Adobe presentation script."
    }
  }

  if (Test-Path -LiteralPath $out) {
    Remove-Item -LiteralPath $out -Force
  }
  $pres.SaveAs($out)
}
finally {
  $pres.Close()
  $ppt.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}

Write-Host $out
