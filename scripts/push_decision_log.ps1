param(
    [Parameter(Mandatory=$false)]
    [string]$Message = "Push",

    [Parameter(Mandatory=$false)]
    [string]$Branch,

    [Parameter(Mandatory=$false)]
    [string]$Tag
)

# -----------------------------------
# REPO VALIDATION
# -----------------------------------
$RepoRoot = git rev-parse --show-toplevel 2>$null
if (-not $RepoRoot) {
    Write-Host "Not inside a git repository" -ForegroundColor Red
    exit 1
}

# -----------------------------------
# PATH SETUP (SEPARATE LOG FILE)
# -----------------------------------
$LogFile = Join-Path $RepoRoot "commit_logs/remote_push_log.md"
$LogDir = Split-Path $LogFile

if (!(Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}

if (!(Test-Path $LogFile)) {
    New-Item -ItemType File -Path $LogFile | Out-Null
}

# -----------------------------------
# CONTEXT EXTRACTION
# -----------------------------------
$Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm"

$LocalBranch = git branch --show-current
$CommitHash = git rev-parse HEAD

if (-not $Branch) {
    $Branch = $LocalBranch
}

$RemoteBranch = "origin/$Branch"

# -----------------------------------
# EXECUTE PUSH
# -----------------------------------
Write-Host "Pushing $LocalBranch → $RemoteBranch ..." -ForegroundColor Yellow

git push origin $Branch
if ($LASTEXITCODE -ne 0) {
    Write-Host "Push failed" -ForegroundColor Red
    exit 1
}

# Optional tag push
if ($Tag) {
    git push origin $Tag
}

# -----------------------------------
# LOG ENTRY (REMOTE PUSH LOG)
# -----------------------------------
$TagInfo = if ($Tag) { $Tag } else { "NA" }

$Entry = @"

---

## $Timestamp — PUSH

**Message:** $Message

**Local Branch:**
$LocalBranch

**Remote Branch:**
$RemoteBranch

**Commit:**
$CommitHash

**Tag Pushed:**
$TagInfo

**Result:**
SUCCESS

"@

Add-Content -Path $LogFile -Value $Entry

# -----------------------------------
# COMMIT LOG ENTRY (SEPARATE FILE)
# -----------------------------------
git add $LogFile
git commit -m "push: $Message"

# -----------------------------------
# OUTPUT
# -----------------------------------
Write-Host "Push completed and logged" -ForegroundColor Green
Write-Host "Branch: $LocalBranch → $RemoteBranch" -ForegroundColor Cyan
Write-Host "Commit: $CommitHash" -ForegroundColor Cyan