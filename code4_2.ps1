Add-Type -AssemblyName System.Windows.Forms

# 監視するフォルダパス
$folder = 'C:\Users\yhk4n\Desktop'
$filter = '*.*'

$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $folder
$watcher.Filter = $filter
$watcher.IncludeSubdirectories = $false
$watcher.EnableRaisingEvents = $true
Write-Host "$folder の監視を開始します。"
$lastRaised = [DateTime]::MinValue

$action = {
    param($source, $eventArgs)
    
    $now = Get-Date
    
    if ($now - $lastRaised -gt [TimeSpan]::FromSeconds(1.5)) {
        $fileName = $eventArgs.Name
        $changeType = $eventArgs.ChangeType
        $message = " $fileName が追加されました。"
        [System.Windows.Forms.MessageBox]::Show($message, "新しいファイルの通知")
        
        $lastRaised = $now
    }
}

Register-ObjectEvent -InputObject $watcher -EventName Created -Action $action | Out-Null

while ($true) { Start-Sleep -Seconds 2 }
