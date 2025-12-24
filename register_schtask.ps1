param(
    [string]$ProjectRoot = "C:\Users\kev\My Drive\Shopify_project",
    [string]$TaskName = "ShopifyExportPlaywright",
    [string]$StartTime = "02:00"
)

$python = Join-Path -Path $ProjectRoot -ChildPath ".venv\Scripts\python.exe"
$script = Join-Path -Path $ProjectRoot -ChildPath "playwright_runner.py"

if (Test-Path $python) {
    $action = '"' + $python + '" "' + $script + '"'
} else {
    $action = '"python" "' + $script + '"'
}

Write-Output "Registering scheduled task '$TaskName' to run daily at $StartTime"
# Create or replace existing scheduled task
schtasks /Create /SC DAILY /TN $TaskName /TR $action /ST $StartTime /F | Out-Null

if ($LASTEXITCODE -eq 0) {
    Write-Output "Scheduled task '$TaskName' created/updated."
} else {
    Write-Output "Failed to create scheduled task; exit code $LASTEXITCODE"
}
