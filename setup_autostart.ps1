# ============================================================
# setup_autostart.ps1 - one-time watcher installation
# Creates a Startup shortcut so the folder watcher runs at
# every logon, then starts it immediately.
# Usage: powershell -ExecutionPolicy Bypass -File setup_autostart.ps1
# ============================================================

$ws = New-Object -ComObject WScript.Shell
$sc = $ws.CreateShortcut("$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\SynergyPPTWatcher.lnk")
$sc.TargetPath = "C:\Users\dpksi\AppData\Local\Programs\Python\Python313\pythonw.exe"
$sc.Arguments  = "C:\SynergyPPTs\synergy_convert.py"
$sc.WorkingDirectory = "C:\SynergyPPTs"
$sc.Save()
Write-Host "Startup shortcut created."

Start-Process "C:\Users\dpksi\AppData\Local\Programs\Python\Python313\pythonw.exe" -ArgumentList "C:\SynergyPPTs\synergy_convert.py"
Write-Host "Watcher started."
