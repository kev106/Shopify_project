param(
    [string]$ProjectRoot = "C:\Users\kev\My Drive\Shopify_project"
)

# Run the project's playwright_runner.py using the venv python if present, otherwise system python
$python = Join-Path -Path $ProjectRoot -ChildPath ".venv\Scripts\python.exe"
if (-not (Test-Path $python)) {
    Write-Output "Virtualenv python not found at $python; falling back to 'python' on PATH."
    $python = "python"
}

$script = Join-Path -Path $ProjectRoot -ChildPath "playwright_runner.py"
$logDir = Join-Path -Path $ProjectRoot -ChildPath "downloads\debug"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
$logFile = Join-Path -Path $logDir -ChildPath "scheduled_run.log"

Write-Output "Running $script using $python; logging to $logFile"
& $python $script *>&1 | Tee-Object -FilePath $logFile -Append

if ($LASTEXITCODE -ne 0) {
    Write-Output "Runner exited with code $LASTEXITCODE"
} else {
    Write-Output "Runner completed successfully"
}
