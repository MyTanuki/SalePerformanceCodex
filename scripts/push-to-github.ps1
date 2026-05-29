param(
  [string]$Message = "Update sales performance dashboard",
  [string]$Branch = "main",
  [switch]$NoPush,
  [switch]$Yes
)

$ErrorActionPreference = "Stop"

$ExpectedRemote = "https://github.com/MyTanuki/SalePerformanceCodex.git"
$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path

function Write-Step {
  param([string]$Text)
  Write-Host ""
  Write-Host "==> $Text" -ForegroundColor Cyan
}

function Find-Git {
  $cmd = Get-Command git -ErrorAction SilentlyContinue
  if ($cmd) { return $cmd.Source }

  $candidates = @(
    "C:\Program Files\Git\cmd\git.exe",
    "C:\Program Files\Git\bin\git.exe",
    "C:\Program Files (x86)\Git\cmd\git.exe"
  )
  foreach ($candidate in $candidates) {
    if (Test-Path $candidate) { return $candidate }
  }
  throw "ไม่พบ git.exe กรุณาติดตั้ง Git for Windows หรือเพิ่ม git ลง PATH"
}

function Find-Node {
  $cmd = Get-Command node -ErrorAction SilentlyContinue
  if ($cmd) { return $cmd.Source }

  $candidates = @(
    "$env:USERPROFILE\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe",
    "C:\Program Files\nodejs\node.exe"
  )
  foreach ($candidate in $candidates) {
    if ($candidate -and (Test-Path $candidate)) { return $candidate }
  }
  return $null
}

function Invoke-Git {
  param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Args)
  & $script:GitPath @Args
  if ($LASTEXITCODE -ne 0) {
    throw "Git command failed: git $($Args -join ' ')"
  }
}

function Confirm-Step {
  param([string]$Question)
  if ($Yes) { return $true }
  $answer = Read-Host "$Question (y/N)"
  return $answer -match "^(y|yes)$"
}

function Clear-StaleGitLock {
  $lockPath = Join-Path $RepoRoot ".git\index.lock"
  if (-not (Test-Path $lockPath)) { return }

  $gitProcesses = Get-Process | Where-Object { $_.ProcessName -eq "git" -or $_.ProcessName -eq "git.exe" }
  if ($gitProcesses) {
    $gitProcesses | Select-Object Id, ProcessName, Path | Format-Table -AutoSize
    throw "พบ git process กำลังทำงานอยู่ จึงไม่ลบ .git/index.lock"
  }

  if (-not (Confirm-Step "พบ .git/index.lock ค้างอยู่ ต้องการลบ lock นี้หรือไม่")) {
    throw "ยกเลิก เพราะยังมี .git/index.lock"
  }

  Remove-Item -LiteralPath $lockPath -Force
}

function Convert-PorcelainLine {
  param([string]$Line)
  if ($Line.Length -lt 4) { return $null }

  $status = $Line.Substring(0, 2)
  $pathText = $Line.Substring(3)
  if ($pathText -match " -> ") {
    $pathText = ($pathText -split " -> ", 2)[1]
  }
  $pathText = $pathText.Trim('"')
  [pscustomobject]@{
    Status = $status
    Path = $pathText
  }
}

function Test-BlockedPath {
  param([string]$Path)
  $normalized = $Path -replace "\\", "/"
  $fileName = [IO.Path]::GetFileName($normalized)
  $extension = [IO.Path]::GetExtension($normalized).ToLowerInvariant()

  if ($fileName -ieq "desktop.ini") { return $true }
  if ($extension -eq ".csv") { return $true }
  if ($normalized -match "(^|/)(backup|temp|tmp|node_modules|dist|build)(/|$)") { return $true }
  if ($fileName -match "\.(bak|tmp|log)$") { return $true }
  return $false
}

function Test-AllowedPath {
  param([string]$Path)
  if (Test-BlockedPath $Path) { return $false }

  $extension = [IO.Path]::GetExtension($Path).ToLowerInvariant()
  $allowedExtensions = @(
    ".html", ".js", ".css", ".md", ".json", ".cmd", ".ps1", ".txt", ".yml", ".yaml"
  )
  return $allowedExtensions -contains $extension
}

Set-Location $RepoRoot
$script:GitPath = Find-Git

Write-Step "ตรวจสอบ repository"
$remote = (& $GitPath remote get-url origin 2>$null).Trim()
if ($LASTEXITCODE -ne 0 -or -not $remote) {
  throw "ไม่พบ remote origin"
}
if ($remote -ne $ExpectedRemote) {
  throw "remote origin ไม่ตรงกับที่กำหนด: $remote"
}
Write-Host "Remote: $remote"
Write-Host "Branch: $Branch"

Clear-StaleGitLock

Write-Step "ตรวจสอบไฟล์ที่เปลี่ยนแปลง"
$statusLines = & $GitPath status --porcelain
if (-not $statusLines) {
  Write-Host "ไม่มีไฟล์เปลี่ยนแปลงให้ commit" -ForegroundColor Green
  exit 0
}

$changes = $statusLines | ForEach-Object { Convert-PorcelainLine $_ } | Where-Object { $_ -ne $null }
$allowed = @($changes | Where-Object { Test-AllowedPath $_.Path })
$blocked = @($changes | Where-Object { -not (Test-AllowedPath $_.Path) })

if ($allowed.Count) {
  Write-Host "ไฟล์ที่จะนำเข้า commit:" -ForegroundColor Green
  $allowed | ForEach-Object { Write-Host ("  {0} {1}" -f $_.Status, $_.Path) }
}

if ($blocked.Count) {
  Write-Host ""
  Write-Host "ไฟล์ที่ถูกกันออกจาก commit:" -ForegroundColor Yellow
  $blocked | ForEach-Object { Write-Host ("  {0} {1}" -f $_.Status, $_.Path) }
}

if (-not $allowed.Count) {
  Write-Host "ไม่มีไฟล์ที่ผ่านเงื่อนไขสำหรับ commit" -ForegroundColor Yellow
  exit 0
}

if (-not (Confirm-Step "ยืนยัน stage เฉพาะไฟล์ข้างต้น")) {
  throw "ยกเลิกก่อน stage"
}

Write-Step "ตรวจ syntax"
$node = Find-Node
if ($node -and (Test-Path (Join-Path $RepoRoot "app.js"))) {
  & $node --check (Join-Path $RepoRoot "app.js")
  if ($LASTEXITCODE -ne 0) { throw "app.js syntax check failed" }
  Write-Host "app.js syntax check passed" -ForegroundColor Green
} else {
  Write-Host "ข้าม syntax check เพราะไม่พบ node.exe" -ForegroundColor Yellow
}

Write-Step "Stage files"
Invoke-Git reset
foreach ($item in $allowed) {
  Invoke-Git add -- $item.Path
}

Write-Host ""
Invoke-Git status --short
Write-Host ""
Invoke-Git diff --cached --stat

if (-not (Confirm-Step "ยืนยัน commit ด้วยข้อความ: $Message")) {
  throw "ยกเลิกก่อน commit"
}

Write-Step "Commit"
Invoke-Git commit -m $Message

if ($NoPush) {
  Write-Host "สร้าง commit แล้ว แต่ข้าม push เพราะระบุ -NoPush" -ForegroundColor Yellow
  exit 0
}

if (-not (Confirm-Step "ยืนยัน pull --rebase และ push ไป origin/$Branch")) {
  throw "ยกเลิกก่อน push"
}

Write-Step "Pull rebase"
Invoke-Git pull --rebase origin $Branch

Write-Step "Push"
Invoke-Git push origin $Branch

Write-Step "Done"
Invoke-Git status --short --branch
