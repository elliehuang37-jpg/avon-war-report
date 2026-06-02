# Avon Daily Battle Report - Auto Update Script
# Usage: Right-click -> Run with PowerShell
param([string]$XL = "")

$DataJs  = Join-Path $PSScriptRoot "war_data.js"
$HtmlOut = Join-Path $PSScriptRoot "index.html"

# Auto-detect latest real time Excel on Desktop
if (-not $XL) {
    $searchPaths = @($PSScriptRoot, "$env:USERPROFILE\Desktop")
    $latest = $null
    foreach ($path in $searchPaths) {
        $found = Get-ChildItem $path -Filter "*.xlsx" | Where-Object { $_.Name -match "real.?time" } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($found) { $latest = $found; break }
    }
    if ($latest) {
        $XL = $latest.FullName
        Write-Host "Auto-detected: $($latest.Name)" -ForegroundColor Cyan
    } else {
        Write-Host "ERROR: Cannot find real time xlsx!" -ForegroundColor Red
        Write-Host "Please put the Excel file in the same folder or on Desktop." -ForegroundColor Yellow
        Read-Host "Press Enter to close"
        exit
    }
}

function Num($v)   { try { [math]::Round([double]$v) } catch { 0 } }
function Pct($a,$b){ $bn = Num $b; if($bn -gt 0){ [math]::Round((Num $a)/$bn*100,1) } else { 0 } }
function Esc($s)   { "$s" -replace '"','\"' }

Write-Host "=== Avon Daily Report Updater ===" -ForegroundColor Magenta
Write-Host "Reading: $XL"

$app = New-Object -ComObject Excel.Application
$app.Visible = $false
$app.DisplayAlerts = $false

try {
    $wb = $app.Workbooks.Open($XL)
    $sheets = $wb.Worksheets | ForEach-Object { $_.Name }
    Write-Host "Sheets: $($sheets -join ' | ')"

    $ws = $wb.Worksheets.Item($sheets[0])
    $updateTime = $ws.Cells(2,2).Text

    $r4=4; $r5=5; $r6=6; $r7=7; $r11=11; $r12=12; $r13=13; $r14=14; $r18=18; $r19=19; $r20=20; $r21=21

    $sTA=Num $ws.Cells($r4,4).Value2; $sTB=Num $ws.Cells($r5,4).Value2; $sTC=Num $ws.Cells($r6,4).Value2
    $sPvA=Num $ws.Cells($r4,5).Value2; $sPvB=Num $ws.Cells($r5,5).Value2; $sPvC=Num $ws.Cells($r6,5).Value2
    $sDTA=Num $ws.Cells($r4,8).Value2; $sDTB=Num $ws.Cells($r5,8).Value2; $sDTC=Num $ws.Cells($r6,8).Value2
    $sDAA=Num $ws.Cells($r4,9).Value2; $sDAB=Num $ws.Cells($r5,9).Value2; $sDAC=Num $ws.Cells($r6,9).Value2
    $sCAA=Num $ws.Cells($r4,12).Value2; $sCAB=Num $ws.Cells($r5,12).Value2; $sCAC=Num $ws.Cells($r6,12).Value2
    $sGPA=Num $ws.Cells($r4,13).Value2; $sGPB=Num $ws.Cells($r5,13).Value2; $sGPC=Num $ws.Cells($r6,13).Value2
    $sPcA=Pct $sCAA $sTA; $sPcB=Pct $sCAB $sTB; $sPcC=Pct $sCAC $sTC

    $aTA=Num $ws.Cells($r11,4).Value2; $aTB=Num $ws.Cells($r12,4).Value2; $aTC=Num $ws.Cells($r13,4).Value2
    $aDTA=Num $ws.Cells($r11,8).Value2; $aDTB=Num $ws.Cells($r12,8).Value2; $aDTC=Num $ws.Cells($r13,8).Value2
    $aDAA=Num $ws.Cells($r11,9).Value2; $aDAB=Num $ws.Cells($r12,9).Value2; $aDAC=Num $ws.Cells($r13,9).Value2
    $aCAA=Num $ws.Cells($r11,12).Value2; $aCAB=Num $ws.Cells($r12,12).Value2; $aCAC=Num $ws.Cells($r13,12).Value2
    $aPcA=Pct $aCAA $aTA; $aPcB=Pct $aCAB $aTB; $aPcC=Pct $aCAC $aTC

    $rTA=Num $ws.Cells($r18,4).Value2; $rTB=Num $ws.Cells($r19,4).Value2; $rTC=Num $ws.Cells($r20,4).Value2
    $rDTA=Num $ws.Cells($r18,8).Value2; $rDTB=Num $ws.Cells($r19,8).Value2; $rDTC=Num $ws.Cells($r20,8).Value2
    $rDAA=Num $ws.Cells($r18,9).Value2; $rDAB=Num $ws.Cells($r19,9).Value2; $rDAC=Num $ws.Cells($r20,9).Value2
    $rCAA=Num $ws.Cells($r18,12).Value2; $rCAB=Num $ws.Cells($r19,12).Value2; $rCAC=Num $ws.Cells($r20,12).Value2
    $rPcA=Pct $rCAA $rTA; $rPcB=Pct $rCAB $rTB; $rPcC=Pct $rCAC $rTC

    $ttlSTgt=Num $ws.Cells($r7,4).Value2; $ttlSCum=Num $ws.Cells($r7,12).Value2; $ttlSGap=Num $ws.Cells($r7,13).Value2
    $ttlATgt=Num $ws.Cells($r14,4).Value2; $ttlACum=Num $ws.Cells($r14,12).Value2
    $ttlRTgt=Num $ws.Cells($r21,4).Value2; $ttlRCum=Num $ws.Cells($r21,12).Value2
    $ttlSPct=Pct $ttlSCum $ttlSTgt; $ttlAPct=Pct $ttlACum $ttlATgt; $ttlRPct=Pct $ttlRCum $ttlRTgt

    $dST=Num $ws.Cells($r7,8).Value2; $dSA=Num $ws.Cells($r7,9).Value2; $dSPct=Pct $dSA $dST
    $dAT=Num $ws.Cells($r14,8).Value2; $dAA=Num $ws.Cells($r14,9).Value2; $dAPct=Pct $dAA $dAT
    $dRT=Num $ws.Cells($r21,8).Value2; $dRA=Num $ws.Cells($r21,9).Value2; $dRPct=Pct $dRA $dRT

    $repsArr = @()
    foreach ($sname in $sheets) {
        if ($sname -match "^\^|raw|Target|Holiday|SAFDR|0504|0601|三合一|by") { continue }
        $wsR = $wb.Worksheets.Item($sname)
        $used = $wsR.UsedRange.Rows.Count
        for ($row = 3; $row -le $used; $row++) {
            $dv = $wsR.Cells($row,2).Text
            if ($dv -notmatch "^[ABC]$") { continue }
            $nm = Esc $wsR.Cells($row,7).Text
            if (-not $nm) { continue }
            $bdmN=Esc $wsR.Cells($row,3).Text; $famN=Esc $wsR.Cells($row,4).Text; $zn=$wsR.Cells($row,5).Text
            $tgt=try{[math]::Round([double]$wsR.Cells($row,8).Value2)}catch{0}
            $da=try{[math]::Round([double]$wsR.Cells($row,11).Value2)}catch{0}
            $ca=try{[math]::Round([double]$wsR.Cells($row,13).Value2)}catch{0}
            $gap=try{[math]::Round([double]$wsR.Cells($row,15).Value2)}catch{0}
            $pc=if($tgt -gt 0){[math]::Round($ca/$tgt*100,1)}else{0}
            $repsArr += "{`"div`":`"$dv`",`"bdm`":`"$bdmN`",`"family`":`"$famN`",`"zone`":`"$zn`",`"name`":`"$nm`",`"target`":$tgt,`"dayActual`":$da,`"cumActual`":$ca,`"gap`":$gap,`"pct`":$pc}"
        }
    }
    Write-Host "Reps: $($repsArr.Count)"

    $wsd = $wb.Worksheets.Item("^Daily%")
    $dailyArr = @()
    for ($row = 3; $row -le 34; $row++) {
        $dt=$wsd.Cells($row,2).Text; $wd=$wsd.Cells($row,3).Text; $sp=$wsd.Cells($row,4).Text
        $dA2=try{[math]::Round([double]$wsd.Cells($row,5).Value2)}catch{0}
        $dB2=try{[math]::Round([double]$wsd.Cells($row,6).Value2)}catch{0}
        $dC2=try{[math]::Round([double]$wsd.Cells($row,7).Value2)}catch{0}
        $dT2=try{[math]::Round([double]$wsd.Cells($row,8).Value2)}catch{0}
        $ap=$wsd.Cells($row,9).Text
        $at2=try{[math]::Round([double]$wsd.Cells($row,13).Value2)}catch{0}
        $rp=$wsd.Cells($row,15).Text
        $rt2=try{[math]::Round([double]$wsd.Cells($row,19).Value2)}catch{0}
        if ($dt -and $dT2 -gt 0) {
            $dailyArr += "{`"date`":`"$dt`",`"wd`":`"$wd`",`"salesPct`":`"$sp`",`"A`":$dA2,`"B`":$dB2,`"C`":$dC2,`"ttlSales`":$dT2,`"activePct`":`"$ap`",`"ttlActive`":$at2,`"recruitPct`":`"$rp`",`"ttlRecruit`":$rt2}"
        }
    }

    $nl = [System.Environment]::NewLine
    $js  = "window.AVON_DATA = {$nl"
    $js += "  updateTime: `"$updateTime`",$nl"
    $js += "  excelFile: `"$(Split-Path $XL -Leaf)`",$nl"
    $js += "  ttl: { sales: { target: $ttlSTgt, cumAct: $ttlSCum, gap: $ttlSGap, pct: $ttlSPct }, active: { target: $ttlATgt, cumAct: $ttlACum, pct: $ttlAPct }, recruit: { target: $ttlRTgt, cumAct: $ttlRCum, pct: $ttlRPct } },$nl"
    $js += "  todayKPI: { salesDayTgt: $dST, salesDayAct: $dSA, salesDayPct: $dSPct, activeDayTgt: $dAT, activeDayAct: $dAA, activeDayPct: $dAPct, recruitDayTgt: $dRT, recruitDayAct: $dRA, recruitDayPct: $dRPct },$nl"
    $js += "  bdm: { names: [`"Eva Ma`",`"Angus Tsai`",`"Lerry Chang`"], divs: [`"A`",`"B`",`"C`"],$nl"
    $js += "    sales:   { target:[$sTA,$sTB,$sTC], prevAct:[$sPvA,$sPvB,$sPvC], cumAct:[$sCAA,$sCAB,$sCAC], gap:[$sGPA,$sGPB,$sGPC], dayTgt:[$sDTA,$sDTB,$sDTC], dayAct:[$sDAA,$sDAB,$sDAC], pct:[$sPcA,$sPcB,$sPcC] },$nl"
    $js += "    active:  { target:[$aTA,$aTB,$aTC], cumAct:[$aCAA,$aCAB,$aCAC], dayTgt:[$aDTA,$aDTB,$aDTC], dayAct:[$aDAA,$aDAB,$aDAC], pct:[$aPcA,$aPcB,$aPcC] },$nl"
    $js += "    recruit: { target:[$rTA,$rTB,$rTC], cumAct:[$rCAA,$rCAB,$rCAC], dayTgt:[$rDTA,$rDTB,$rDTC], dayAct:[$rDAA,$rDAB,$rDAC], pct:[$rPcA,$rPcB,$rPcC] }$nl"
    $js += "  },$nl"
    $js += "  daily: [$($dailyArr -join ',')],$nl"
    $js += "  reps:  [$($repsArr  -join ',')]$nl"
    $js += "};$nl"

    [System.IO.File]::WriteAllText($DataJs, $js, [System.Text.Encoding]::UTF8)
    Write-Host "Saved: $DataJs" -ForegroundColor Green
    Write-Host "Time: $updateTime | Reps: $($repsArr.Count)" -ForegroundColor Green

} finally {
    try { $wb.Close($false) } catch {}
    $app.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
}

if (Test-Path $HtmlOut) {
    Start-Process $HtmlOut
    Write-Host "Report opened!" -ForegroundColor Cyan
}
