function Walk($s,$label){
 $txt=''; try{if($s.TextFrame.HasText -ne 0){$txt=$s.TextFrame.TextRange.Text.Trim()}}catch{}
 if($txt -match 'PRIORITY|BE READY|AMMUNITION|HOW TO USE|PROFESSIONAL|PRO TIP|SAFETY'){
   $sn=$txt -replace "`r|`n"," "; if($sn.Length -gt 180){$sn=$sn.Substring(0,180)}
   $page=''; try{$page=$s.Anchor.Information(3)}catch{}
   Write-Output ("{0} Page={1} Type={2} L={3:N1} T={4:N1} W={5:N1} H={6:N1} Text={7}" -f $label,$page,$s.Type,$s.Left,$s.Top,$s.Width,$s.Height,$sn)
 }
 try{if($s.Type -eq 6){for($j=1;$j -le $s.GroupItems.Count;$j++){Walk $s.GroupItems.Item($j) "$label.$j"}}}catch{}
}
$docx='C:\Users\tyler\GitHub\wild-eyez-packing-list\working-files\PRIMARY_DESIGN_SOURCE_USE_THIS.docx'
$word=New-Object -ComObject Word.Application; $word.Visible=$false; $doc=$word.Documents.Open($docx)
for($i=1;$i -le $doc.Shapes.Count;$i++){Walk $doc.Shapes.Item($i) "#$i"}
$doc.Close($false);$word.Quit()
