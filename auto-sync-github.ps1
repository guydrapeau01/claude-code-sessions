# Auto-Sync to GitHub - Daily Backup
# This script commits and pushes all changes automatically

# Change to your CLAUDE directory
Set-Location "C:\Users\GuyDrapeauTHORASYS\OneDrive - Thorasys\Documents\CLAUDE"

# Get current date for commit message
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm"

Write-Host "==================================" -ForegroundColor Cyan
Write-Host "  Auto-Sync to GitHub" -ForegroundColor Cyan
Write-Host "==================================" -ForegroundColor Cyan
Write-Host ""

# Check if there are any changes
$status = git status --porcelain

if ($status) {
    Write-Host "✓ Changes detected" -ForegroundColor Green
    Write-Host ""
    
    # Show what changed
    Write-Host "Files changed:" -ForegroundColor Yellow
    git status --short
    Write-Host ""
    
    # Add all changes
    Write-Host "→ Adding files..." -ForegroundColor Yellow
    git add .
    
    # Commit with auto-generated message
    Write-Host "→ Creating commit..." -ForegroundColor Yellow
    git commit -m "Auto-backup: $timestamp"
    
    # Push to GitHub
    Write-Host "→ Pushing to GitHub..." -ForegroundColor Yellow
    $pushResult = git push 2>&1
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host ""
        Write-Host "✓ SUCCESS! Changes pushed to GitHub" -ForegroundColor Green
        Write-Host "   Time: $timestamp" -ForegroundColor Cyan
    } else {
        Write-Host ""
        Write-Host "✗ Push failed:" -ForegroundColor Red
        Write-Host $pushResult -ForegroundColor Red
        
        # Log error to file
        $errorLog = "git-sync-errors.log"
        "$timestamp - Push failed: $pushResult" | Out-File -Append $errorLog
        Write-Host ""
        Write-Host "Error logged to: $errorLog" -ForegroundColor Yellow
    }
    
} else {
    Write-Host "ℹ No changes to commit" -ForegroundColor Yellow
    Write-Host "   Everything is up to date!" -ForegroundColor Cyan
}

Write-Host ""
Write-Host "==================================" -ForegroundColor Cyan
Write-Host "  Sync Complete" -ForegroundColor Cyan
Write-Host "==================================" -ForegroundColor Cyan

# Exit successfully
exit 0
