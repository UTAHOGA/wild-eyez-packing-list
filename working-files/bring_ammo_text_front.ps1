$docx='C:\Users\tyler\GitHub\wild-eyez-packing-list\working-files\PRIMARY_DESIGN_SOURCE_USE_THIS.docx'
$word=New-Object -ComObject Word.Application; $word.Visible=$false; $doc=$word.Documents.Open($docx)
for($i=1;$i -le $doc.Shapes.Count;$i++){
  $s=$doc.Shapes.Item($i)
  $txt=''
  try{if($s.TextFrame.HasText -ne 0){$txt=$s.TextFrame.TextRange.Text.Trim().ToUpper()}}catch{}
  if($txt.StartsWith('AMMUNITION STANDARD')){
    try{$s.ZOrder(0)}catch{}
    try{$s.Left=47.5; $s.Width=517.3}catch{}
  }
}
$doc.Save(); $doc.Close($false); $word.Quit(); Write-Output 'ammo text brought to front'
