function Dump-Shape($s, $prefix){
  $text=''
  try { if($s.TextFrame.HasText -ne 0){ $text=$s.TextFrame.TextRange.Text.Trim() } } catch {}
  $children=0; try { if($s.Type -eq 6){$children=$s.GroupItems.Count} } catch {}
  $snippet=$text -replace "`r|`n"," "
  if($snippet.Length -gt 120){$snippet=$snippet.Substring(0,120)}
  $page=''; try {$page=$s.Anchor.Information(3)} catch {}
  Write-Output ("{0} Type={1} Children={2} Page={3} L={4:N1} T={5:N1} W={6:N1} H={7:N1} Text={8}" -f $prefix,$s.Type,$children,$page,$s.Left,$s.Top,$s.Width,$s.Height,$snippet)
  try {
    if($s.Type -eq 6){ for($j=1;$j -le $s.GroupItems.Count;$j++){ Dump-Shape $s.GroupItems.Item($j) ("$prefix.$j") } }
  } catch {}
}
$docx='C:\Users\tyler\GitHub\wild-eyez-packing-list\working-files\PRIMARY_DESIGN_SOURCE_USE_THIS.docx'
$word = New-Object -ComObject Word.Application; $word.Visible = $false; $doc = $word.Documents.Open($docx)
foreach($n in 58,60,63){ Dump-Shape $doc.Shapes.Item($n) "#$n" }
$doc.Close($false); $word.Quit()
