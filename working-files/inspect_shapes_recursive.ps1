function Dump-Shape($s, $prefix){
  $text=''
  try { if($s.TextFrame.HasText -ne 0){ $text=$s.TextFrame.TextRange.Text.Trim() } } catch {}
  $fill=''; $line=''; $weight=''; $type=''; $left=''; $top=''; $w=''; $h=''; $page=''
  try {$fill=$s.Fill.ForeColor.RGB} catch {}
  try {$line=$s.Line.ForeColor.RGB; $weight=$s.Line.Weight} catch {}
  try {$type=$s.Type; $left=$s.Left; $top=$s.Top; $w=$s.Width; $h=$s.Height; $page=$s.Anchor.Information(3)} catch {}
  if($text.Length -gt 0){
    $snippet=$text -replace "`r|`n"," "
    if($snippet.Length -gt 220){$snippet=$snippet.Substring(0,220)}
    $font=''; $size=''; $bold=''
    try {$font=$s.TextFrame.TextRange.Font.Name; $size=$s.TextFrame.TextRange.Font.Size; $bold=$s.TextFrame.TextRange.Font.Bold} catch {}
    Write-Output ("{0} Type={1} Page={2} L={3:N1} T={4:N1} W={5:N1} H={6:N1} Fill={7} Line={8} LW={9} Font={10} Size={11} Bold={12} Text={13}" -f $prefix,$type,$page,$left,$top,$w,$h,$fill,$line,$weight,$font,$size,$bold,$snippet)
  }
  try {
    if($s.Type -eq 6){
      for($j=1;$j -le $s.GroupItems.Count;$j++){
        Dump-Shape $s.GroupItems.Item($j) ("$prefix.$j")
      }
    }
  } catch {}
}
$docx='C:\Users\tyler\GitHub\wild-eyez-packing-list\working-files\PRIMARY_DESIGN_SOURCE_USE_THIS.docx'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open($docx)
for($i=1;$i -le $doc.Shapes.Count;$i++){ Dump-Shape $doc.Shapes.Item($i) "#$i" }
$doc.Close($false)
$word.Quit()
