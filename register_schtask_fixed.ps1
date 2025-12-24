param(
    [string]$ProjectRoot = "C:\Users\kev\My Drive\Shopify_project",
    [string]$TaskName = "ShopifyExportPlaywright",
    [string]$StartTime = "02:00"
)

$python = Join-Path -Path $ProjectRoot -ChildPath ".venv\Scripts\python.exe"
$script = Join-Path -Path $ProjectRoot -ChildPath "playwright_runner.py"

if (Test-Path $python) {
    $tr = '"' + $python + '" "' + $script + '"'
} else {
    $tr = '"python" "' + $script + '"'
}

Write-Output "Registering scheduled task '$TaskName' to run daily at $StartTime"

# Use Start-Process with an argument list so PowerShell handles quoting properly
$args = @('/Create', '/SC', 'DAILY', '/TN', $TaskName, '/TR', $tr, '/ST', $StartTime, '/F')
$proc = Start-Process -FilePath schtasks -ArgumentList $args -NoNewWindow -Wait -PassThru
if ($proc.ExitCode -eq 0) {
    Write-Output "Scheduled task '$TaskName' created/updated."
} else {
    Write-Output "Failed to create scheduled task; exit code $($proc.ExitCode)"
}
