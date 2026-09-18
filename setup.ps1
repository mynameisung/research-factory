# report-factory setup (Windows) — 의존성 설치 + 한글 환경 + 스모크 테스트
# 사용:  powershell -File setup.ps1
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "report-factory setup" -ForegroundColor Cyan

# 1. Python 확인
try { $pv = (python --version) 2>&1 } catch { Write-Host "python 이 없습니다. https://python.org 에서 3.11+ 를 설치하세요." -ForegroundColor Red; exit 1 }
Write-Host "  $pv"

# 2. 의존성
Write-Host "  pip install -r requirements.txt ..."
python -m pip install -q -r (Join-Path $root "requirements.txt")

# 3. PYTHONUTF8 — cp949 콘솔에서 한글이 깨지지 않게 사용자 환경변수로 고정
if ($env:PYTHONUTF8 -ne "1") {
    setx PYTHONUTF8 1 | Out-Null
    $env:PYTHONUTF8 = "1"
    Write-Host "  PYTHONUTF8=1 을 사용자 환경변수로 설정했습니다 (새 터미널부터 자동 적용)"
}

# 4. 스모크 테스트 — 검수기와 빌더가 기동하는가
python (Join-Path $root "report-factory\engine\tools\verify.py") --help | Out-Null
python (Join-Path $root "report-factory\engine\tools\build_guide.py") --help | Out-Null
python (Join-Path $root "report-factory\engine\tools\gen_preamble.py") --check
$qst = Join-Path $root "report-factory\engine\tools\questionnaire_selftest.py"
python $qst | Out-Null   # 파이프로 받으면 cp949 콘솔에서 한글이 깨진다 — 결과는 종료 코드로 본다
if ($LASTEXITCODE -ne 0) {
    python $qst   # 파이프 없이 다시 돌려 무엇이 실패했는지 보인다
    Write-Host "  설문지 빌더 자체 시험 실패. 위 결과를 관리자에게 보내세요" -ForegroundColor Red
    exit 1
}
Write-Host "  엔진 기동 확인 (verify · build_guide · 설문지 빌더)"

# 5. 안내
Write-Host ""
Write-Host "설치 완료. 이 폴더를 Claude Code 작업 폴더로 여세요." -ForegroundColor Green
Write-Host "  스킬: /설문지작성  /웹테스트  /제작가이드  /보고서작성  /회고  /업그레이드  /개선공유"
Write-Host "  조사 자료는 report-factory\projects\<조사명>\ 에 두세요 (git 이 추적하지 않습니다)."
