param(
    [Parameter(Mandatory=$true)]
    [string]$Message
)

$RepoRoot = git rev-parse --show-toplevel 2>$null
if (-not $RepoRoot) {
    Write-Host "Not inside a git repository" -ForegroundColor Red
    exit 1
}

$LogFile = Join-Path $RepoRoot "commit_logs/decision-log.md"
$LogDir = Split-Path $LogFile

if (!(Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}

if (!(Test-Path $LogFile)) {
    New-Item -ItemType File -Path $LogFile | Out-Null
}

$Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm"

$FilesChanged = git diff --cached --name-only | Out-String
if ([string]::IsNullOrWhiteSpace($FilesChanged)) {
    Write-Host "No staged files found. Use git add first." -ForegroundColor Red
    exit 1
}

$Entry = @"

---

## $Timestamp

**Commit Message:** $Message

**Files Changed:**
$FilesChanged

**Problem:** NA
**Decision:** NA
**Next:** NA

"@

Add-Content -Path $LogFile -Value $Entry

git add $LogFile
git commit -m $Message

$ActualHash = git rev-parse HEAD

Write-Host "Committed: $Message" -ForegroundColor Green
Write-Host "Hash: $ActualHash" -ForegroundColor Cyan