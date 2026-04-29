param(
    [Parameter(Mandatory=$true)]
    [string]$Name
)

# Resolve repo root
$RepoRoot = git rev-parse --show-toplevel 2>$null
if (-not $RepoRoot) {
    Write-Host "Not inside a git repository" -ForegroundColor Red
    exit 1
}

# Define paths
$LogFile = Join-Path $RepoRoot "commit_logs/decision-log.md"
$LogDir = Split-Path $LogFile
$LOG_FILE_REL = "commit_logs/decision-log.md"

$Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm"
$FullTag = "checkpoint/v1.3/$Name"

# Ensure working directory is clean

# Unstaged tracked changes
git diff --quiet -- . ":(exclude)$LOG_FILE_REL"
if ($LASTEXITCODE -ne 0) {
    Write-Host "Unstaged tracked changes present (excluding decision log)." -ForegroundColor Red
    exit 1
}

# Staged but uncommitted
git diff --cached --quiet -- . ":(exclude)$LOG_FILE_REL"
if ($LASTEXITCODE -ne 0) {
    Write-Host "Staged but uncommitted changes present (excluding decision log)." -ForegroundColor Red
    exit 1
}

# Prevent duplicate tag
$TagExists = git tag -l $FullTag
if ($TagExists) {
    Write-Host "Tag already exists: $FullTag" -ForegroundColor Red
    exit 1
}

# Ensure commit_logs directory exists
if (!(Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}

# Ensure log file exists
if (!(Test-Path $LogFile)) {
    New-Item -ItemType File -Path $LogFile | Out-Null
}

# Create tag
git tag $FullTag

# Push ONLY this tag
git push origin $FullTag

# Append log entry
$Entry = @"

---

## $Timestamp — CHECKPOINT

**Tag:** $FullTag

**State at this point:**
NA

**Why checkpoint created:**
NA

**Rollback:**
git reset --hard $FullTag

"@

Add-Content -Path $LogFile -Value $Entry

# Commit log entry
git add $LogFile
git commit -m "checkpoint: $Name"

Write-Host "Checkpoint created: $FullTag" -ForegroundColor Green