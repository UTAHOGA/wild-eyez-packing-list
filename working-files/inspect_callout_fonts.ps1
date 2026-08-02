function DumpTextShape($s, $label){
  try { if($s.TextFrame.HasText -eq 0){return} } catch { return }
  $txt=$s.TextFrame.TextRange.Text.Trim() -replace "`r", "|"
  if($txt -notmatch 'HOW TO USE|PROFESSIONAL|BE READY|PRO TIP|AMMUNITION') { return }
  Write-Output "--- $label W=$($s.Width) H=$($s.Height)"
  $pars=$s.TextFrame.TextRange.Paragraphs()
  for($i=1;$i -le $pars.Count;$i++){
    $p=$pars.Item($i)
    $t=$p.Text.Trim() -replace "`r|`n",""
    if($t.Length -gt 130){$t=$t.Substring(0,130)}
    Write-Output ("P{0}: Font={1} Size={2} Bold={3} Color={4} Text={5}" -f $i,$p.Font.Name,$p.Font.Size,$p.Font.Bold,$p.Font.Color,$t)
  }
}
function Walk($s,$label){
  DumpTextShape $s $label
  try { if($s.Type -eq 6){ for($j=1;$j -le $s.GroupItems.Count;$j++){ Walk $s.GroupItems.Item($j) "$label.$j" } } } catch {}
}
$docx='C:\Users\tyler\GitHub\wild-eyez-packing-list\working-files\PRIMARY_DESIGN_SOURCE_USE_THIS.docx'
$word = New-Object -ComObject Word.Application; $word.Visible = $false; $doc = $word.Documents.Open($docx)
for($i=1;$i -le $doc.Shapes.Count;$i++){ Walk $doc.Shapes.Item($i) "#$i" }
$doc.Close($false); $word.Quit()
