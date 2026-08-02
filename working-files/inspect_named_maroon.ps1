$docx='C:\Users\tyler\GitHub\wild-eyez-packing-list\working-files\PRIMARY_DESIGN_SOURCE_USE_THIS.docx'
$word=New-Object -ComObject Word.Application; $word.Visible=$false; $doc=$word.Documents.Open($docx)
foreach($s in $doc.Shapes){
 if($s.Name -like 'priority_1_safety*' -or $s.Name -like 'priority_safety_point*'){
  $txt=''; try{if($s.TextFrame.HasText -ne 0){$txt=$s.TextFrame.TextRange.Text.Trim() -replace "`r|`n"," "}}catch{}
  if($txt.Length -gt 140){$txt=$txt.Substring(0,140)}
  $page=''; try{$page=$s.Anchor.Information(3)}catch{}
  Write-Output ("Name={0} Page={1} Type={2} L={3:N1} T={4:N1} W={5:N1} H={6:N1} Text={7}" -f $s.Name,$page,$s.Type,$s.Left,$s.Top,$s.Width,$s.Height,$txt)
 }
}
$doc.Close($false);$word.Quit()
