$ErrorActionPreference='Stop'
$p=(Get-ChildItem -LiteralPath 'D:\KLG\Danh_Sach_CSDL' -Filter '*.xlsx' | Select-Object -First 1).FullName
Add-Type -AssemblyName System.IO.Compression.FileSystem
$z=[IO.Compression.ZipFile]::OpenRead($p)
function ReadXml($entry){$d=New-Object Xml.XmlDocument; $sr=New-Object IO.StreamReader($entry.Open()); $d.LoadXml($sr.ReadToEnd()); $sr.Dispose(); return $d}
$wb=ReadXml $z.GetEntry('xl/workbook.xml'); $ns=New-Object Xml.XmlNamespaceManager($wb.NameTable); $ns.AddNamespace('x','http://schemas.openxmlformats.org/spreadsheetml/2006/main'); $ns.AddNamespace('r','http://schemas.openxmlformats.org/officeDocument/2006/relationships'); $rels=ReadXml $z.GetEntry('xl/_rels/workbook.xml.rels'); $rns=New-Object Xml.XmlNamespaceManager($rels.NameTable); $rns.AddNamespace('p','http://schemas.openxmlformats.org/package/2006/relationships')
$ss=@(); $e=$z.GetEntry('xl/sharedStrings.xml'); if($e){$sd=ReadXml $e; foreach($si in $sd.GetElementsByTagName('si')){$ss += (($si.GetElementsByTagName('t')|%{$_.InnerText}) -join '')}}
function GetSheet($name){$sn=$wb.SelectSingleNode("//x:sheet" ,$ns)|?{$_.name -eq $name}; if(-not $sn){$avail=( $wb.SelectNodes('//x:sheet',$ns)|%{$_.name}) -join ','; throw "sheet not found: $name; available=$avail"}; $rid=$sn.GetAttribute('id','http://schemas.openxmlformats.org/officeDocument/2006/relationships'); $rel=$rels.SelectSingleNode("//p:Relationship[@Id='$rid']",$rns); return ReadXml $z.GetEntry(('xl/'+$rel.Target.TrimStart('/')))}
function GetSheetAt($idx){$sn=$wb.SelectNodes('//x:sheet',$ns)[$idx]; $rid=$sn.GetAttribute('id','http://schemas.openxmlformats.org/officeDocument/2006/relationships'); $rel=$rels.SelectSingleNode("//p:Relationship[@Id='$rid']",$rns); return ReadXml $z.GetEntry(('xl/'+$rel.Target.TrimStart('/')))}
function Val($c){$v=$c.SelectSingleNode('./x:v',$ns); if($v){if($c.GetAttribute('t') -eq 's'){return $ss[[int]$v.InnerText]} return $v.InnerText}; $is=$c.SelectSingleNode('./x:is',$ns); if($is){return (($is.GetElementsByTagName('t')|%{$_.InnerText})-join '')}; return ''}
function Rows($d){$out=@(); foreach($row in $d.SelectNodes('//x:sheetData/x:row',$ns)){ $rn=$row.GetAttribute('r'); $v=@{}; foreach($c in $row.SelectNodes('./x:c',$ns)){$v[$c.GetAttribute('r')]=Val $c}; if($v['A'+$rn] -match '^\d+$'){$out += [pscustomobject]@{row=[int]$rn;stt=$v['A'+$rn];name=$v['B'+$rn];title=$v['C'+$rn];time=$v['D'+$rn];gab=$v['F'+$rn];source=$v['G'+$rn]}}}; return $out}
function SourceRows($d){$out=@(); foreach($row in $d.SelectNodes('//x:sheetData/x:row',$ns)){ $rn=$row.GetAttribute('r'); $v=@{}; foreach($c in $row.SelectNodes('./x:c',$ns)){$v[$c.GetAttribute('r')]=Val $c}; if($v['C'+$rn]){$out += [pscustomobject]@{row=[int]$rn;stt=$v['B'+$rn];name=$v['C'+$rn];title=$v['E'+$rn];time=$v['D'+$rn];gab='';source=$v['F'+$rn]}}}; return $out}
$srcDoc=GetSheet 'Data_record_01-01-2022_01-01-20'; $dstDoc=GetSheetAt 1; $src=SourceRows $srcDoc; $dst=Rows $dstDoc
$keys=@{}; foreach($x in $dst){$keys[(($x.name+'|'+$x.title).ToUpper()).Trim()]=$true}
$miss=$src|?{$_.row -gt 1 -and -not $keys.ContainsKey((($_.name+'|'+$_.title).ToUpper()).Trim())}
"SRC=$($src.Count) DST=$($dst.Count) MISSING=$($miss.Count)"
$miss|Sort-Object row|Select-Object row,stt,name,title,time,source|Format-Table -AutoSize -Wrap
$dst|Sort-Object row|Select-Object -Last 8 row,stt,name,title,time,gab,source|Format-Table -AutoSize -Wrap
