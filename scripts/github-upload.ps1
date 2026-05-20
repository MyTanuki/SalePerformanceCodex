param(
  [string]$CommitMessage = "",
  [string]$Branch = ""
)

$ErrorActionPreference = "Stop"
$RepoPath = Resolve-Path (Join-Path $PSScriptRoot "..")

function Confirm-Step {
  param(
    [string]$Message,
    [string]$Default = "N"
  )
  $suffix = if ($Default -eq "Y") { "[Y/n]" } else { "[y/N]" }
  $answer = Read-Host "$Message $suffix"
  if ([string]::IsNullOrWhiteSpace($answer)) {
    $answer = $Default
  }
  return $answer.Trim().ToLowerInvariant().StartsWith("y")
}

function Run-Git {
  param([string[]]$Args)
  & git @Args
  if ($LASTEXITCODE -ne 0) {
    throw "git $($Args -join ' ') failed."
  }
}

Set-Location $RepoPath
Write-Host ""
Write-Host "GitHub Upload Workflow" -ForegroundColor Cyan
Write-Host "Repository path: $RepoPath"
Write-Host ""

$gitCommand = Get-Command git -ErrorAction SilentlyContinue
if (-not $gitCommand) {
  Write-Host "ไม่พบ git ใน PATH ของเครื่องนี้" -ForegroundColor Red
  Write-Host "ให้ติดตั้ง Git for Windows หรือเพิ่ม git.exe เข้า PATH แล้วเรียกใช้ไฟล์นี้อีกครั้ง"
  exit 1
}

if (-not (Test-Path (Join-Path $RepoPath ".git"))) {
  Write-Host "โฟลเดอร์นี้ยังไม่ใช่ Git repository" -ForegroundColor Red
  Write-Host "ถ้าต้องการใช้ workflow นี้ ให้ clone repo จาก GitHub หรือรัน git init/ตั้ง remote ก่อน"
  exit 1
}

Write-Host "Step 1: ตรวจสอบ remote" -ForegroundColor Yellow
Run-Git @("remote", "-v")
if (-not (Confirm-Step "ยืนยันว่า remote GitHub ถูกต้องหรือไม่?")) {
  Write-Host "ยกเลิกขั้นตอน upload"
  exit 0
}

$currentBranch = (& git branch --show-current).Trim()
Write-Host ""
Write-Host "Current branch: $currentBranch"
if (-not [string]::IsNullOrWhiteSpace($Branch) -and $Branch -ne $currentBranch) {
  if (Confirm-Step "ต้องการ switch/create branch '$Branch' ก่อนทำงานหรือไม่?") {
    Run-Git @("checkout", "-B", $Branch)
    $currentBranch = $Branch
  }
}

Write-Host ""
Write-Host "Step 2: ตรวจสอบไฟล์ที่เปลี่ยนแปลง" -ForegroundColor Yellow
Run-Git @("status", "--short")
if (-not (Confirm-Step "ต้องการ stage ไฟล์ทั้งหมดที่เปลี่ยนแปลงหรือไม่?")) {
  Write-Host "ยกเลิกก่อน stage ไฟล์"
  exit 0
}

Run-Git @("add", "--all")
Write-Host ""
Write-Host "Step 3: ตรวจสอบไฟล์ที่ staged แล้ว" -ForegroundColor Yellow
Run-Git @("diff", "--cached", "--stat")
if (-not (Confirm-Step "ยืนยัน commit ไฟล์ staged เหล่านี้หรือไม่?")) {
  Write-Host "ยกเลิกก่อน commit ไฟล์ staged ยังอยู่ สามารถตรวจสอบด้วย git status"
  exit 0
}

if ([string]::IsNullOrWhiteSpace($CommitMessage)) {
  $CommitMessage = Read-Host "กรอก commit message"
}
if ([string]::IsNullOrWhiteSpace($CommitMessage)) {
  Write-Host "Commit message ว่าง ยกเลิกขั้นตอน commit" -ForegroundColor Red
  exit 1
}

Run-Git @("commit", "-m", $CommitMessage)

Write-Host ""
Write-Host "Step 4: Push ไป GitHub" -ForegroundColor Yellow
Run-Git @("status", "--short", "--branch")
if (-not (Confirm-Step "ต้องการ push branch '$currentBranch' ไป GitHub หรือไม่?")) {
  Write-Host "Commit สำเร็จแล้ว แต่ยังไม่ได้ push"
  exit 0
}

Run-Git @("push", "-u", "origin", $currentBranch)
Write-Host ""
Write-Host "Upload ไป GitHub สำเร็จ" -ForegroundColor Green
