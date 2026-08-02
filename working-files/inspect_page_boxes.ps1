function ShapeText($s){
  try { if($s.TextFrame.HasText -ne 0){ return ($s.TextFrame.TextRange.Text.Trim() -replace "`r|`n"," ") } } catch {}
  return ''
}
$docx='C:\Users\tyler\GitHub\wild-eyez-packing-list\working-files\PRIMARY_DESIGN_SOURCE_USE_THIS.docx'
$word = New-Object -ComObject Word.Application; $word.Visible = $false; $doc = $word.Documents.Open($docx)
for($i=1;$i -le $doc.Shapes.Count;$i++){
  $s=$doc.Shapes.Item($i)
  $page=$null; try{$page=$s.Anchor.Information(3)}catch{}
  if($page -in 2,3,4,6,7){
    $txt=ShapeText $s
    if($txt.Length -gt 100){$txt=$txt.Substring(0,100)}
    $fill='';$line='';$lw='';$vis='';$type='';$children=0
    try{$fill=$s.Fill.ForeColor.RGB}catch{}
    try{$line=$s.Line.ForeColor.RGB;$lw=$s.Line.Weight;$vis=$s.Line.Visible}catch{}
    try{$type=$s.Type;if($type -eq 6){$children=$s.GroupItems.Count}}catch{}
    if(($s.Left -gt 25 -and $s.Left -lt 70 -and $s.Width -gt 480) -or $txt -match 'PRIORITY|PROFESSIONAL|BE READY|PRO TIP|AMMUNITION|HOW TO USE|COMMON'){
      Write-Output ("#{0} Page={1} Type={2} Child={3} L={4:N1} T={5:N1} W={6:N1} H={7:N1} Fill={8} Line={9} LW={10} Vis={11} Text={12}" -f $i,$page,$type,$children,$s.Left,$s.Top,$s.Width,$s.Height,$fill,$line,$lw,$vis,$txt)
      if($type -eq 6){
        for($j=1;$j -le $s.GroupItems.Count;$j++){
          $c=$s.GroupItems.Item($j); $ct=ShapeText $c; if($ct.Length -gt 80){$ct=$ct.Substring(0,80)}
          $cf='';$cl='';$clw='';$cv=''; try{$cf=$c.Fill.ForeColor.RGB}catch{}; try{$cl=$c.Line.ForeColor.RGB;$clw=$c.Line.Weight;$cv=$c.Line.Visible}catch{}
          Write-Output ("  .{0} Type={1} L={2:N1} T={3:N1} W={4:N1} H={5:N1} Fill={6} Line={7} LW={8} Vis={9} Text={10}" -f $j,$c.Type,$c.Left,$c.Top,$c.Width,$c.Height,$cf,$cl,$clw,$cv,$ct)
        }
      }
    }
  }
}
$doc.Close($false); $word.Quit()
