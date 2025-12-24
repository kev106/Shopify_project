param(
    [string]$ProjectRoot = "C:\Users\kev\My Drive\Shopify_project",
    [string]$TaskName = "ShopifyExportPlaywright",
    [string]$StartTime = "02:00"
)

$cmd = Join-Path -Path $ProjectRoot -ChildPath "run_playwright_runner.cmd"
$tr = '"' + $cmd + '"'
Write-Output "Registering scheduled task '$TaskName' to run weekly on Friday at $StartTime (command: $cmd)"
$args = @('/Create', '/SC', 'WEEKLY', '/D', 'FRI', '/TN', $TaskName, '/TR', $tr, '/ST', $StartTime, '/F')
$proc = Start-Process -FilePath schtasks -ArgumentList $args -NoNewWindow -Wait -PassThru
if ($proc.ExitCode -eq 0) {
    Write-Output "Scheduled task '$TaskName' created/updated (Weekly on Friday)."
} else {
    Write-Output "Failed to create scheduled task; exit code $($proc.ExitCode)"
}
