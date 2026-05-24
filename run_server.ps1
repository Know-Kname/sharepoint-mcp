# SharePoint MCP — HTTP server wrapper with auto-restart
# Registered via Task Scheduler; runs under user credentials
# Transport: streamable-http on port 8765
# Logs to: V:\repos\sharepoint-mcp\service_stderr.log

$python = 'V:\repos\sharepoint-mcp\.venv-312\Scripts\python.exe'
$script = 'V:\repos\sharepoint-mcp\server.py'
$workDir = 'V:\repos\sharepoint-mcp'
$logFile = 'V:\repos\sharepoint-mcp\service_stderr.log'
$args = @('--transport', 'streamable-http', '--host', '127.0.0.1', '--port', '8765')

Set-Location $workDir

while ($true) {
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Add-Content $logFile "[$timestamp] Starting SharePoint MCP server..."
    
    try {
        $proc = Start-Process -FilePath $python `
            -ArgumentList (@($script) + $args) `
            -WorkingDirectory $workDir `
            -NoNewWindow `
            -Wait `
            -PassThru `
            -RedirectStandardError "$workDir\server_stderr_raw.log"
        
        $exitCode = $proc.ExitCode
        $ts2 = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Add-Content $logFile "[$ts2] Server exited with code $exitCode. Restarting in 5s..."
    } catch {
        $ts2 = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Add-Content $logFile "[$ts2] Exception: $_. Restarting in 5s..."
    }
    
    Start-Sleep -Seconds 5
}
