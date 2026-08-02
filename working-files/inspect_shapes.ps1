$docx='C:\Users\tyler\GitHub\wild-eyez-packing-list\working-files\PRIMARY_DESIGN_SOURCE_USE_THIS.docx'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open($docx)
$i=0
foreach($s in $doc.Shapes){
  $i++
  $text=''
  try { if($s.TextFrame.HasText -ne 0){ $text=$s.TextFrame.TextRange.Text.Trim() } } catch {}
  if($text.Length -gt 0){
    $snippet=$text -replace "`r|`n"," "
    if($snippet.Length -gt 160){$snippet=$snippet.Substring(0,160)}
    $font=''; $size=''; $bold='';
    try {$font=$s.TextFrame.TextRange.Font.Name; $size=$s.TextFrame.TextRange.Font.Size; $bold=$s.TextFrame.TextRange.Font.Bold} catch {}
    Write-Output ("#{0} Page={1} Left={2:N1} Top={3:N1} W={4:N1} H={5:N1} Font={6} Size={7} Bold={8} Text={9}" -f $i,$s.Anchor.Information(3),$s.Left,$s.Top,$s.Width,$s.Height,$font,$size,$bold,$snippet)
  }
}
$doc.Close($false)
$word.Quit()
