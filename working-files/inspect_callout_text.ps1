function Walk($s,$label){
  $text=''
  try { if($s.TextFrame.HasText -ne 0){$text=$s.TextFrame.TextRange.Text} } catch {}
  if($text -match 'HOW TO USE|PROFESSIONAL|BE READY|PRO TIP|AMMUNITION'){
    Write-Output "--- $label W=$($s.Width) H=$($s.Height)"
    $escaped=($text -replace "`r", '<CR>' -replace "`n", '<LF>')
    Write-Output $escaped
    for($i=1;$i -le [Math]::Min(60,$text.Length);$i++){
      $ch=$text.Substring($i-1,1)
      $code=[int][char]$ch
      Write-Output "$i '$ch' $code"
    }
  }
  try { if($s.Type -eq 6){ for($j=1;$j -le $s.GroupItems.Count;$j++){ Walk $s.GroupItems.Item($j) "$label.$j" } } } catch {}
}
$docx='C:\Users\tyler\GitHub\wild-eyez-packing-list\working-files\PRIMARY_DESIGN_SOURCE_USE_THIS.docx'
$word = New-Object -ComObject Word.Application; $word.Visible = $false; $doc = $word.Documents.Open($docx)
for($i=1;$i -le $doc.Shapes.Count;$i++){ Walk $doc.Shapes.Item($i) "#$i" }
$doc.Close($false); $word.Quit()
