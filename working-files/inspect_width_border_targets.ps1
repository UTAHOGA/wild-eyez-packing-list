function Walk($s,$label){
  $text=''
  try { if($s.TextFrame.HasText -ne 0){$text=$s.TextFrame.TextRange.Text.Trim()} } catch {}
  $fill=''; $line=''; $weight=''; $visible=''; $type=''; $page=''
  try {$fill=$s.Fill.ForeColor.RGB} catch {}
  try {$line=$s.Line.ForeColor.RGB; $weight=$s.Line.Weight; $visible=$s.Line.Visible} catch {}
  try {$type=$s.Type; $page=$s.Anchor.Information(3)} catch {}
  $snippet=$text -replace "`r|`n"," "
  if($snippet.Length -gt 120){$snippet=$snippet.Substring(0,120)}
  $isHit = $false
  if($snippet -match 'HOW TO USE|PROFESSIONAL|BE READY|PRO TIP|AMMUNITION|PRIORITY SAFETY|PRIORITY #1|MOST COMMON'){$isHit=$true}
  try { if($s.Type -eq 6){
    for($j=1;$j -le $s.GroupItems.Count;$j++){
      $child=$s.GroupItems.Item($j); $ct=''
      try { if($child.TextFrame.HasText -ne 0){$ct=$child.TextFrame.TextRange.Text.Trim()} } catch {}
      if($ct -match 'HOW TO USE|PROFESSIONAL|BE READY|PRO TIP|AMMUNITION|PRIORITY SAFETY|PRIORITY #1|MOST COMMON'){$isHit=$true}
    }
  }} catch {}
  if($isHit -or ($label -match '^#(50|52|58|60|63)$')){
    Write-Output ("{0} Type={1} Page={2} L={3:N1} T={4:N1} W={5:N1} H={6:N1} Fill={7} Line={8} LW={9} Vis={10} Text={11}" -f $label,$type,$page,$s.Left,$s.Top,$s.Width,$s.Height,$fill,$line,$weight,$visible,$snippet)
  }
  try { if($s.Type -eq 6){ for($j=1;$j -le $s.GroupItems.Count;$j++){ Walk $s.GroupItems.Item($j) "$label.$j" } } } catch {}
}
$docx='C:\Users\tyler\GitHub\wild-eyez-packing-list\working-files\PRIMARY_DESIGN_SOURCE_USE_THIS.docx'
$word = New-Object -ComObject Word.Application; $word.Visible = $false; $doc = $word.Documents.Open($docx)
for($i=1;$i -le $doc.Shapes.Count;$i++){ Walk $doc.Shapes.Item($i) "#$i" }
$doc.Close($false); $word.Quit()
