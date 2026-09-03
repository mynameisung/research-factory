# =========================================================================
# 설치.ps1 — 통계표전처리.xlam 을 엑셀 추가 기능으로 설치한다.
#
# 왜 이 스크립트가 필요한가
#   .xlam 을 아무 폴더에나 두고 열면 엑셀이 「매크로를 사용할 수 없도록
#   설정했습니다」라는 보안 알림을 띄운다. 신뢰할 수 없는 위치에 있기 때문이다.
#   추가 기능 전용 폴더(%APPDATA%\Microsoft\AddIns)는 엑셀이 처음부터 믿는
#   자리라, 그리로 옮겨 등록하면 알림 없이 곧바로 쓸 수 있다.
#
# 하는 일
#   1) 인터넷·메일로 받은 파일에 붙는 차단 표시를 뗀다
#   2) 추가 기능 폴더로 복사한다
#   3) 엑셀에 추가 기능으로 등록하고 체크를 켠다
#
#   powershell -ExecutionPolicy Bypass -File 설치.ps1
# =========================================================================
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

# 소스 찾기 — 스크립트 옆이나 out 폴더
$cands = @(
    (Join-Path $root '통계표전처리.xlam'),
    (Join-Path $root 'out\통계표전처리.xlam')
)
$src = $cands | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $src) { throw "통계표전처리.xlam 을 찾지 못했습니다. 이 스크립트와 같은 폴더에 두고 다시 실행해 주세요." }

$dstDir = Join-Path $env:APPDATA 'Microsoft\AddIns'
if (-not (Test-Path $dstDir)) { New-Item -ItemType Directory -Path $dstDir -Force | Out-Null }
$dst = Join-Path $dstDir '통계표전처리.xlam'

# 1) 차단 표시 제거
Unblock-File -Path $src -ErrorAction SilentlyContinue

# 2) 복사 — 이미 쓰고 있으면 엑셀을 닫아 달라고 알린다
try {
    Copy-Item $src $dst -Force
} catch {
    Write-Host ''
    Write-Host '[중단] 추가 기능 파일을 바꾸지 못했습니다.' -ForegroundColor Yellow
    Write-Host '엑셀이 이미 이 추가 기능을 쓰고 있을 수 있습니다.'
    Write-Host '엑셀을 모두 닫은 뒤 다시 실행해 주세요.'
    exit 1
}
Unblock-File -Path $dst -ErrorAction SilentlyContinue
Write-Host "[1/2] 추가 기능 폴더로 복사했습니다." -ForegroundColor Green
Write-Host "      $dst"

# 3) 엑셀에 등록
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    if ([double]$xl.Version -lt 14) {
        Write-Host ''
        Write-Host "[주의] 엑셀 $($xl.Version) 에서는 리본 탭이 뜨지 않습니다." -ForegroundColor Yellow
        Write-Host '이 추가 기능의 리본은 엑셀 2010(버전 14) 이상에서 동작합니다.'
    }

    $ai = $null
    foreach ($a in $xl.AddIns) { if ($a.Name -eq '통계표전처리.xlam') { $ai = $a } }
    if ($null -eq $ai) { $ai = $xl.AddIns.Add($dst, $true) }
    $ai.Installed = $true
    Write-Host "[2/2] 엑셀에 추가 기능으로 등록했습니다." -ForegroundColor Green
} finally {
    if ($xl.Workbooks.Count -eq 0) { $xl.Quit() }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}

Write-Host ''
Write-Host '설치를 마쳤습니다. 엑셀을 열면 [통계표 전처리] 탭이 보입니다.' -ForegroundColor Cyan
Write-Host '보이지 않으면 파일 → 옵션 → 추가 기능 → 관리: Excel 추가 기능 → 이동 에서'
Write-Host '[통계표 전처리] 에 체크해 주세요.'
