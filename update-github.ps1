# One-Click GitHub Update Script
# Double-click this file to commit and push all changes to GitHub

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  GitHub Auto-Update" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# Check if we're in a git repository
if (-not (Test-Path .git)) {
    Write-Host "❌ Error: Not a git repository!" -ForegroundColor Red
    Write-Host "   Make sure you're running this from your CLAUDE folder" -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

# Check for changes
Write-Host "→ Checking for changes..." -ForegroundColor Yellow
$status = git status --porcelain

if ($status) {
    Write-Host "✅ Found changes to commit" -ForegroundColor Green
    Write-Host ""
    
    # Show what's changed
    Write-Host "Changes detected:" -ForegroundColor Cyan
    git status --short
    Write-Host ""
    
    # Ask for commit message
    $message = Read-Host "Enter commit message (or press Enter for auto-message)"
    
    if ([string]::IsNullOrWhiteSpace($message)) {
        $date = Get-Date -Format "yyyy-MM-dd HH:mm"
        $message = "Auto-update: $date"
    }
    
    # Add all changes
    Write-Host ""
    Write-Host "→ Adding files..." -ForegroundColor Yellow
    git add .
    
    # Commit
    Write-Host "→ Creating commit..." -ForegroundColor Yellow
    git commit -m "$message"
    
    # Push
    Write-Host "→ Pushing to GitHub..." -ForegroundColor Yellow
    git push
    
    Write-Host ""
    Write-Host "✅ SUCCESS! Your changes are on GitHub!" -ForegroundColor Green
    Write-Host ""
    Write-Host "View at: https://github.com/guydrapeau01/claude-code-sessions" -ForegroundColor Cyan
    
} else {
    Write-Host "ℹ️  No changes to commit" -ForegroundColor Yellow
    Write-Host "   Everything is already up to date!" -ForegroundColor Cyan
}

Write-Host ""
Read-Host "Press Enter to exit"
