# update-dashboard.ps1
# Laeser nyaste .xlsm-fil, genererar data.json och pushar till GitHub.
param(
    [string]$WatchFolder = "C:\Users\angelica.molin\OneDrive - Silva Sweden AB\Document Center - E-commerce Private\5. E-com Planner",
    [string]$RepoFolder  = "C:\Users\angelica.molin\ecom-dashboard"
)
$ERR = -2146826246
$logFile = Join-Path $RepoFolder "scripts\update.log"
function Log($msg){$ts=(Get-Date).ToString("yyyy-MM-dd HH:mm:ss");Add-Content -Path $logFile -Value "[$ts] $msg" -Encoding UTF8;Write-Host "[$ts] $msg"}
function IsVal($v){return ($null -ne $v -and $v -ne $ERR -and $v -ne "")}
function SafeInt($v){if(IsVal($v)){return [int][Math]::Round([double]$v)}else{return 0}}
function SafeDec1($v){if(IsVal($v)){return [Math]::Round([double]$v,1)}else{return 0.0}}
function EscJ($s){return $s.ToString().Replace('\','\\').Replace('"','\"')}

Log "=== Startar uppdatering ==="

$newest = Get-ChildItem $WatchFolder -Filter "*.xlsm" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($null -eq $newest){Log "FEL: Ingen xlsm hittades";exit 1}
Log ("Fil: " + $newest.Name)

$excelWasOpen = $false
$excel = $null
try{
    $excel=[System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    $excelWasOpen=$true
    Log "Ansluten till korrande Excel"
}catch{
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false
    Log "Startade ny Excel"
}
$excel.DisplayAlerts=$false

$wb=$null
if($excelWasOpen){foreach($openWb in $excel.Workbooks){if($openWb.Name -eq $newest.Name){$wb=$openWb;break}}}
if($null -eq $wb){$wb=$excel.Workbooks.Open($newest.FullName,0,$true);Log "Oppnade workbook"}
else{Log "Workbook var redan opppen"}

$ws=$wb.Sheets.Item("Overview")
Log "Laser Overview..."

$budgetArr=@();$salesArr=@()
for($c=7;$c -le 18;$c++){
    $b=$ws.Cells.Item(2,$c).Value2;$s=$ws.Cells.Item(3,$c).Value2
    $budgetArr+=if(IsVal($b)){SafeInt $b}else{0}
    $salesArr+=if(IsVal($s)){SafeInt $s}else{0}
}
$ytdBudget=SafeInt($ws.Cells.Item(2,19).Value2)
$ytdSales=SafeInt($ws.Cells.Item(3,19).Value2)
$rawPct=$ws.Cells.Item(4,19).Value2
$budgetPct=if(IsVal($rawPct)){[Math]::Round([double]$rawPct*100,1)}else{0.0}

$y2024=@(0)*12;$y2025=@(0)*12;$y2026=@(0)*12
$catMap=@{}
for($row=9;$row -le 2729;$row++){
    $status=$ws.Cells.Item($row,6).Value2
    if(-not(IsVal($status))-or $status -eq "Spare Parts"){continue}
    $cat=[string]$ws.Cells.Item($row,5).Value2
    if(-not $catMap.ContainsKey($cat)){$catMap[$cat]=@{y2024=0;y2025=0;y2026=0}}
    for($m=0;$m -lt 12;$m++){
        $v24=SafeInt($ws.Cells.Item($row,77+$m).Value2)
        $v25=SafeInt($ws.Cells.Item($row,90+$m).Value2)
        $v26=SafeInt($ws.Cells.Item($row,103+$m).Value2)
        $y2024[$m]+=$v24;$y2025[$m]+=$v25;$y2026[$m]+=$v26
        $catMap[$cat].y2024+=$v24;$catMap[$cat].y2025+=$v25;$catMap[$cat].y2026+=$v26
    }
}
Log "Historisk data klar"

$repItems=@();$cntH=0;$cntM=0;$cntC=0;$w0has=0;$w0empty=0
for($row=9;$row -le 2729;$row++){
    $status=$ws.Cells.Item($row,6).Value2
    if(-not(IsVal($status))-or $status -eq "Spare Parts"){continue}
    $wh0=$ws.Cells.Item($row,36).Value2;$wh15=$ws.Cells.Item($row,128).Value2
    $moh=$ws.Cells.Item($row,33).Value2;$slm=$ws.Cells.Item($row,37).Value2
    if(-not(IsVal($wh0))){continue}
    $w0v=SafeInt $wh0
    if($w0v -gt 0){$w0has++}else{$w0empty++}
    if(-not(IsVal($wh15))-or $wh15 -le 0){continue}
    $w15v=SafeInt $wh15;$mv=SafeDec1 $moh;$sv=SafeInt $slm
    $urg=$null
    if($w0v -eq 0){$urg="critical";$cntC++}
    elseif($mv -lt 1){$urg="high";$cntH++}
    elseif($mv -lt 3){$urg="medium";$cntM++}
    if($null -ne $urg){
        $iN=EscJ $ws.Cells.Item($row,2).Value2;$iA=EscJ $ws.Cells.Item($row,3).Value2
        $br=EscJ $ws.Cells.Item($row,4).Value2;$ca=EscJ $ws.Cells.Item($row,5).Value2;$st=EscJ $status
        $repItems+="{`"itemNo`":`"$iN`",`"itemName`":`"$iA`",`"brand`":`"$br`",`"category`":`"$ca`",`"status`":`"$st`",`"moh`":$mv,`"wh0Stock`":$w0v,`"wh15Stock`":$w15v,`"soldLM`":$sv,`"urgency`":`"$urg`"}"
    }
}
$totalRep=$cntC+$cntH+$cntM
Log ("Replenishment: " + $totalRep)

$catJson=@()
foreach($cat in ($catMap.GetEnumerator()|Sort-Object{$_.Value.y2026}-Descending)){
    if($cat.Value.y2026 -gt 0 -or $cat.Value.y2025 -gt 0){
        $cn=EscJ $cat.Key
        $catJson+="{`"name`":`"$cn`",`"y2024`":" + $cat.Value.y2024 + ",`"y2025`":" + $cat.Value.y2025 + ",`"y2026`":" + $cat.Value.y2026 + "}"
    }
}

$ts=(Get-Date -Format "yyyy-MM-ddTHH:mm:ss")
$fnSafe=EscJ $newest.Name
$months='["Maj 26","Jun 26","Jul 26","Aug 26","Sep 26","Okt 26","Nov 26","Dec 26","Jan 27","Feb 27","Mar 27","Apr 27"]'

$json = "{`n  `"lastUpdated`": `"$ts`",`n  `"fileName`": `"$fnSafe`",`n  `"period`": `"202605-202704`",`n  `"summary`": {`n    `"ytdGrossSales`": $ytdSales,`n    `"ytdBudget`": $ytdBudget,`n    `"budgetPct`": $budgetPct,`n    `"months`": $months,`n    `"budget`": [" + ($budgetArr -join ",") + "],`n    `"grossSales`": [" + ($salesArr -join ",") + "]`n  },`n  `"historicalQty`": {`n    `"months`": [`"Jan`",`"Feb`",`"Mar`",`"Apr`",`"Maj`",`"Jun`",`"Jul`",`"Aug`",`"Sep`",`"Okt`",`"Nov`",`"Dec`"],`n    `"y2024`": [" + ($y2024 -join ",") + "],`n    `"y2025`": [" + ($y2025 -join ",") + "],`n    `"y2026`": [" + ($y2026 -join ",") + "]`n  },`n  `"categories`": [" + ($catJson -join ",") + "],`n  `"replenishment`": [" + ($repItems -join ",") + "],`n  `"stockSummary`": {`n    `"wh0HasStock`": $w0has,`n    `"wh0Empty`": $w0empty,`n    `"needsReplenishment`": $totalRep,`n    `"critical`": $cntC,`n    `"high`": $cntH,`n    `"medium`": $cntM`n  }`n}"

$jsonPath=Join-Path $RepoFolder "data.json"
$json|Out-File -FilePath $jsonPath -Encoding utf8 -NoNewline
Log "data.json skriven"

if(-not $excelWasOpen){$wb.Close($false);$excel.Quit();[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)|Out-Null;Log "Excel staengdes"}

Set-Location $RepoFolder
try{
    git add data.json 2>&1|Out-Null
    $changes=git status --porcelain 2>&1
    if($changes){
        $msg="Auto-update: " + $newest.Name + " (" + (Get-Date -Format "yyyy-MM-dd HH:mm") + ")"
        git commit -m $msg 2>&1|Out-Null
        git push origin main 2>&1|Out-Null
        Log "Git push klar"
    }else{Log "Inga andringar"}
}catch{Log ("Git-fel: " + $_.Exception.Message)}
Log "=== Klart ==="
