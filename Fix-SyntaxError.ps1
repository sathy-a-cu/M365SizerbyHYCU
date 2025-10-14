# Fix-SyntaxError.ps1
# This script fixes the PowerShell syntax error

Write-Host "Fixing PowerShell syntax error..." -ForegroundColor Yellow

# Read the current file
$lines = Get-Content "Get-HYCUM365SizingInfo.ps1"

# Fix the problematic line
for ($i = 0; $i -lt $lines.Count; $i++) {
    if ($lines[$i] -match 'Failed to install.*\$\(\$_\.Ex') {
        $lines[$i] = $lines[$i] -replace '\$\(\$_\.Ex[^"]*', '$($_.Exception.Message)'
        Write-Host "Fixed line $($i + 1): $($lines[$i])" -ForegroundColor Green
    }
}

# Write the fixed content back
$lines | Set-Content "Get-HYCUM365SizingInfo.ps1" -Encoding UTF8

Write-Host "Syntax error fixed! You can now run the script." -ForegroundColor Green
Write-Host "Try running: .\Get-HYCUM365SizingInfo.ps1" -ForegroundColor Cyan