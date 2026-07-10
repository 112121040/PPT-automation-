# ============================================================
# watch.ps1 - Folder Watcher (pure PowerShell alternative)
# Watches the Inputs folder and runs the converter script
# automatically whenever a new .ppt/.pptx file arrives.
# Usage: powershell -ExecutionPolicy Bypass -File watch.ps1
# ============================================================

$inputs  = "C:\SynergyPPTs\Inputs"
$convert = "C:\SynergyPPTs\convert.ps1"   # PowerShell converter to trigger
$log     = "C:\SynergyPPTs\watch_log.txt"

function Write-Log($msg) {
    "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $msg" | Add-Content -Path $log
}

function Wait-FileReady($path) {
    # wait until the file copy is finished (can be opened exclusively)
    for ($i = 0; $i -lt 60; $i++) {
        try {
            $fs = [IO.File]::Open($path, 'Open', 'Read', 'None')
            $fs.Close()
            return $true
        }
        catch { Start-Sleep -Seconds 2 }
    }
    return $false
}

Write-Log "Watcher started. Monitoring: $inputs"

$fsw = New-Object System.IO.FileSystemWatcher
$fsw.Path = $inputs
$fsw.Filter = "*.ppt*"
$fsw.IncludeSubdirectories = $false
$fsw.EnableRaisingEvents = $true

Register-ObjectEvent $fsw Created -SourceIdentifier PPTCreated | Out-Null
Register-ObjectEvent $fsw Renamed -SourceIdentifier PPTRenamed | Out-Null

$busy = $false
try {
    while ($true) {
        $evt = Wait-Event -Timeout 5
        if ($evt) {
            $name = $evt.SourceEventArgs.Name
            Remove-Event -EventIdentifier $evt.EventIdentifier
            Get-Event -ErrorAction SilentlyContinue | Remove-Event

            if ($name -like '~$*') { continue }
            if ($busy) { continue }

            $busy = $true
            Write-Log "New file detected: $name"

            $fullPath = Join-Path $inputs $name
            if (Test-Path $fullPath) {
                if (Wait-FileReady $fullPath) {
                    Write-Log "Running conversion..."
                    try {
                        & powershell -ExecutionPolicy Bypass -File $convert *>&1 |
                            Add-Content -Path $log
                        Write-Log "Conversion finished."
                    }
                    catch {
                        Write-Log "ERROR: $($_.Exception.Message)"
                    }
                }
                else {
                    Write-Log "File never became ready (still locked): $name"
                }
            }
            $busy = $false
        }
    }
}
finally {
    Unregister-Event -SourceIdentifier PPTCreated -ErrorAction SilentlyContinue
    Unregister-Event -SourceIdentifier PPTRenamed -ErrorAction SilentlyContinue
    $fsw.Dispose()
    Write-Log "Watcher stopped."
}
