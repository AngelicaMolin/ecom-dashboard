# watcher.ps1
# Ovakar OneDrive-mappen och triggar update-dashboard.ps1 nar ny xlsm detekteras.
# Kors automatiskt som schemalagd uppgift vid inloggning.
param(
    [string]$WatchFolder = "C:\Users\angelica.molin\OneDrive - Silva Sweden AB\Document Center - E-commerce Private\5. E-com Planner",
    [string]$RepoFolder  = "C:\Users\angelica.molin\ecom-dashboard"
)
$logFile=Join-Path $RepoFolder "scripts\update.log"
function Log($msg){$ts=(Get-Date).ToString("yyyy-MM-dd HH:mm:ss");Add-Content -Path $logFile -Value "[$ts] $msg" -Encoding UTF8;Write-Host "[$ts] $msg"}

Log "Watcher startad. Overvakar: $WatchFolder"

$watcher=New-Object System.IO.FileSystemWatcher
$watcher.Path=$WatchFolder
$watcher.Filter="*.xlsm"
$watcher.NotifyFilter=[System.IO.NotifyFilters]::FileName -bor [System.IO.NotifyFilters]::LastWrite
$watcher.EnableRaisingEvents=$true

$script:lastTriggered=[datetime]::MinValue

$action={
    $name=$Event.SourceEventArgs.Name
    $type=$Event.SourceEventArgs.ChangeType
    $now=Get-Date
    if(($now-$script:lastTriggered).TotalSeconds -lt 60){return}
    $script:lastTriggered=$now
    Add-Content -Path $using:logFile -Value ("[" + $now.ToString("yyyy-MM-dd HH:mm:ss") + "] Fil detekterad: $name ($type)") -Encoding UTF8
    Start-Sleep -Seconds 30
    $updateScript=Join-Path $using:RepoFolder "scripts\update-dashboard.ps1"
    try{
        & powershell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass -File $updateScript
    }catch{
        Add-Content -Path $using:logFile -Value ("[" + (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") + "] FEL: " + $_.Exception.Message) -Encoding UTF8
    }
}

Register-ObjectEvent -InputObject $watcher -EventName Created -Action $action -SourceIdentifier "xlsm_created"|Out-Null
Register-ObjectEvent -InputObject $watcher -EventName Changed -Action $action -SourceIdentifier "xlsm_changed"|Out-Null

Log "Watcher aktiv. Vantar pa fil..."
while($true){Start-Sleep -Seconds 30}
