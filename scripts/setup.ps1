# setup.ps1
# Kors EN GANG som administratoer.
# Klonar repo och registrerar watcher som schemalagd uppgift vid inloggning.

$repoUrl    = "https://github.com/AngelicaMolin/ecom-dashboard.git"
$repoFolder = "C:\Users\angelica.molin\ecom-dashboard"
$taskName   = "EcomDashboardWatcher"

Write-Host ""
Write-Host "=== D2C Dashboard Setup ===" -ForegroundColor Cyan
Write-Host ""

if(Test-Path $repoFolder){
    Write-Host "Repomapp finns: $repoFolder" -ForegroundColor Yellow
    Set-Location $repoFolder
    git pull origin main
}else{
    Write-Host "Klonar repo..."
    git clone $repoUrl $repoFolder
}
if(-not(Test-Path $repoFolder)){Write-Host "FEL: Klon misslyckades" -ForegroundColor Red;exit 1}

Set-Location $repoFolder
$gitUser=git config user.name 2>$null
if(-not $gitUser){
    git config user.name "AngelicaMolin"
    git config user.email "angelica.molin@primus-silva.com"
    Write-Host "Git-anvandare konfigurerad" -ForegroundColor Green
}

$watcherScript="$repoFolder\scripts\watcher.ps1"
$psExe="C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
$psArgs="-NonInteractive -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$watcherScript`""

$existing=Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if($existing){Unregister-ScheduledTask -TaskName $taskName -Confirm:$false;Write-Host "Gammal uppgift borttagen" -ForegroundColor Yellow}

$action=New-ScheduledTaskAction -Execute $psExe -Argument $psArgs
$trigger=New-ScheduledTaskTrigger -AtLogOn
$settings=New-ScheduledTaskSettingsSet -ExecutionTimeLimit ([timespan]::Zero) -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)
Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -RunLevel Highest -Description "D2C Dashboard watcher" |Out-Null

Write-Host ""
Write-Host "=== Setup klar! ===" -ForegroundColor Green
Write-Host ""
Write-Host "Starta watchern nu med:"
Write-Host "  Start-ScheduledTask -TaskName '$taskName'" -ForegroundColor Cyan
Write-Host ""
Write-Host "Dashboard-URL: https://angelicamolin.github.io/ecom-dashboard" -ForegroundColor Cyan
Write-Host ""
