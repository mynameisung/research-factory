<#
make_pdf.ps1 — HTML 리포트를 Chrome 헤드리스로 PDF 변환한다.

사용 예:
  powershell -ExecutionPolicy Bypass -File make_pdf.ps1 -Html "C:\작업\보고서.html" -Pdf "C:\작업\보고서.pdf"
  powershell -ExecutionPolicy Bypass -File make_pdf.ps1 -Html ".\report.html" -Pdf ".\report.pdf" -Wait 10

실측으로 확인된 함정 두 가지를 이 스크립트가 처리한다.
  1) 전용 --user-data-dir 을 주지 않으면, 이미 실행 중인 Chrome 인스턴스가 있을 때
     새 프로세스가 기존 프로세스에 위임하고 즉시 종료한다. 종료코드는 0인데 PDF는
     생성되지 않는다 (오류 메시지도 없다).
  2) 출력 경로에 한글이 들어가면 PDF 생성이 실패한다. 그래서 ASCII 경로인
     $env:TEMP\wtpdf 에 만든 뒤 최종 경로로 복사한다.
     입력 HTML 경로는 file:// URI 로 변환하면 한글이 퍼센트 인코딩되어 문제없다.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Html,

    [Parameter(Mandatory = $true)]
    [string]$Pdf,

    [int]$Wait = 6
)

$ErrorActionPreference = 'Stop'

function Write-Fail {
    param([string]$Message)
    Write-Host "[실패] $Message"
}

# --- 1. 입력 HTML 확인 -------------------------------------------------------
if (-not (Test-Path -LiteralPath $Html)) {
    Write-Fail "입력 HTML 파일이 없습니다: $Html"
    exit 1
}
$htmlFull = (Resolve-Path -LiteralPath $Html).ProviderPath

# 한글 경로는 [System.Uri] 가 자동으로 퍼센트 인코딩해 준다.
$htmlUri = ([System.Uri]$htmlFull).AbsoluteUri

# --- 2. Chrome 실행 파일 탐색 ------------------------------------------------
$candidates = @(
    "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
    "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
    "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
)

$chrome = $null
foreach ($c in $candidates) {
    if ([string]::IsNullOrWhiteSpace($c)) { continue }
    if (Test-Path -LiteralPath $c) { $chrome = $c; break }
}

if ($null -eq $chrome) {
    Write-Fail "Chrome 실행 파일을 찾지 못했습니다. 확인한 경로:"
    foreach ($c in $candidates) { Write-Host "         - $c" }
    Write-Host "         Chrome 을 설치하거나, 설치 경로를 이 스크립트의 후보 목록에 추가하세요."
    exit 1
}

# --- 3. 최종 출력 경로의 상위 디렉터리 준비 ----------------------------------
# Pdf 는 아직 존재하지 않을 수 있으므로 Resolve-Path 대신 직접 절대경로화한다.
if ([System.IO.Path]::IsPathRooted($Pdf)) {
    $pdfFull = [System.IO.Path]::GetFullPath($Pdf)
}
else {
    $pdfFull = [System.IO.Path]::GetFullPath((Join-Path (Get-Location).ProviderPath $Pdf))
}

$pdfDir = [System.IO.Path]::GetDirectoryName($pdfFull)
if (-not (Test-Path -LiteralPath $pdfDir)) {
    New-Item -ItemType Directory -Path $pdfDir -Force | Out-Null
}

# --- 4. ASCII 임시 작업 디렉터리 ---------------------------------------------
# 출력 경로에 한글이 있으면 Chrome 이 PDF 를 쓰지 못한다 (실측). 임시 경로 경유.
$workDir = Join-Path $env:TEMP ("wtpdf\" + [System.Guid]::NewGuid().ToString("N").Substring(0, 8))
$profileDir = Join-Path $workDir "profile"
$tmpPdf = Join-Path $workDir "out.pdf"

New-Item -ItemType Directory -Path $workDir -Force | Out-Null
New-Item -ItemType Directory -Path $profileDir -Force | Out-Null

$exitCode = -1
$procErr = $null

try {
    # --- 5. Chrome 헤드리스 실행 ---------------------------------------------
    $chromeArgs = @(
        "--headless=new"
        "--disable-gpu"
        "--no-sandbox"
        "--no-first-run"
        "--no-default-browser-check"
        "--run-all-compositor-stages-before-draw"
        "--virtual-time-budget=8000"
        "--no-pdf-header-footer"
        "--user-data-dir=$profileDir"      # 반드시 필요. 없으면 조용히 실패한다.
        "--print-to-pdf=$tmpPdf"
        "$htmlUri"
    )

    Write-Host "[변환] $chrome"
    Write-Host "[입력] $htmlFull"
    Write-Host "[출력] $pdfFull"

    $proc = Start-Process -FilePath $chrome -ArgumentList $chromeArgs `
        -Wait -NoNewWindow -PassThru
    $exitCode = $proc.ExitCode

    # 렌더링 직후 파일 플러시가 늦는 경우가 있어 짧게 대기하며 확인한다.
    $deadline = (Get-Date).AddSeconds($Wait)
    while (-not (Test-Path -LiteralPath $tmpPdf)) {
        if ((Get-Date) -gt $deadline) { break }
        Start-Sleep -Milliseconds 300
    }
}
catch {
    $procErr = $_.Exception.Message
}

# --- 6. 결과 판정 -------------------------------------------------------------
if ($null -ne $procErr) {
    Write-Fail "Chrome 실행 중 오류: $procErr"
    Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
    exit 1
}

if ($exitCode -ne 0) {
    Write-Fail "Chrome 이 종료코드 $exitCode 로 끝났습니다."
    Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
    exit 1
}

if (-not (Test-Path -LiteralPath $tmpPdf)) {
    Write-Fail "Chrome 이 정상 종료했지만 PDF 가 생성되지 않았습니다 (임시 경로: $tmpPdf)."
    Write-Host "         확인할 것: HTML 이 외부 리소스를 기다리며 멈추지 않는지,"
    Write-Host "         -Wait 값을 늘려야 하는지, 다른 Chrome 프로세스가 프로필을 점유하지 않는지."
    Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
    exit 1
}

$tmpInfo = Get-Item -LiteralPath $tmpPdf
if ($tmpInfo.Length -eq 0) {
    Write-Fail "생성된 PDF 크기가 0 바이트입니다. HTML 내용 또는 -Wait 값을 확인하세요."
    Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
    exit 1
}

# --- 7. 최종 경로로 복사 (한글 경로도 Copy-Item 은 정상 동작) -----------------
$copyErr = $null
try {
    Copy-Item -LiteralPath $tmpPdf -Destination $pdfFull -Force
}
catch {
    $copyErr = $_.Exception.Message
}

Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue

if ($null -ne $copyErr) {
    Write-Fail "최종 경로로 복사하지 못했습니다: $copyErr"
    exit 1
}

if (-not (Test-Path -LiteralPath $pdfFull)) {
    Write-Fail "최종 경로에 파일이 없습니다: $pdfFull"
    exit 1
}

$finalInfo = Get-Item -LiteralPath $pdfFull
$kb = [math]::Round($finalInfo.Length / 1KB, 1)
Write-Host "[성공] $pdfFull ($kb KB)"
exit 0
