$ErrorActionPreference = "Stop"

$repo = "C:\Users\tyler\GitHub\wild-eyez-packing-list"
$source = Join-Path $repo "WILD EYEZ PACKING AND PREPARATION GUIDE_PRESENTATION_WITH_NOTES.pptx"
$audio = Join-Path $repo "assets\audio\AUDIO_ADOBE_UNDER_16MB_128K.mp3"
$out = Join-Path $repo "WILD EYEZ PACKING GUIDE_TIMED_AUDIO_PRESENTATION.pptx"

if (!(Test-Path -LiteralPath $source)) { throw "Missing source presentation: $source" }
if (!(Test-Path -LiteralPath $audio)) { throw "Missing audio: $audio" }

# Estimated page durations based on current script word count and 12:06 audio.
$durations = @(43.68, 153.47, 106.76, 93.43, 77.37, 38.83, 137.98, 103.72, 71.17)

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$pres = $null

try {
  $pres = $ppt.Presentations.Open($source, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue)
  if ($pres.Slides.Count -lt 9) { throw "Expected 9 slides." }

  for ($i = 1; $i -le 9; $i++) {
    $slide = $pres.Slides.Item($i)
    $transition = $slide.SlideShowTransition
    $transition.AdvanceOnClick = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $transition.AdvanceOnTime = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $transition.AdvanceTime = [double]$durations[$i - 1]
  }

  $slide1 = $pres.Slides.Item(1)
  $media = $slide1.Shapes.AddMediaObject2($audio, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue, 18, 742, 34, 34)
  $media.Name = "Adobe-safe narration audio - plays across slides"
  try {
    $media.AnimationSettings.PlaySettings.PlayOnEntry = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $media.AnimationSettings.PlaySettings.HideWhileNotPlaying = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $media.AnimationSettings.PlaySettings.LoopUntilStopped = [Microsoft.Office.Core.MsoTriState]::msoFalse
    $media.AnimationSettings.PlaySettings.StopAfterSlides = 9
  } catch {
    # Some PowerPoint installs expose media playback settings differently.
  }

  try {
    $pres.SlideShowSettings.AdvanceMode = 2 # ppSlideShowUseSlideTimings
  } catch {}

  if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Force }
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
