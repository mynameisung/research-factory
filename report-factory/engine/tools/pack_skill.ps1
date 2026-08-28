# pack_skill.ps1 — 제작가이드 스킬을 다른 PC 로 옮길 압축본으로 묶는다 (2026-08-26 신설)
#
#   powershell -File engine\tools\pack_skill.ps1
#   powershell -File engine\tools\pack_skill.ps1 -Project 2026_에스원_건물현황실사
#   powershell -File engine\tools\pack_skill.ps1 -OutDir D:\백업
#
# 무엇을 담나
#   .claude\skills\ · .claude\agents\ · .claude\settings.json
#   report-factory\engine\{tools,knowledge,learning,templates}
#   report-factory\{README,INSTRUCTIONS,HANDOVER*}.md · handover\
#   (-Project 를 주면) report-factory\projects\<그 조사>
#
# 무엇을 안 담나 — 용량 때문이다. 없어도 스킬은 돈다.
#   engine\refs\        157 MB  예시 제작가이드·제작본과 참고자료 코퍼스. 규칙의 근거이자
#                               벤치마크 재료이지 제작가이드를 만드는 데 쓰이지는 않는다
#   projects\ 나머지     400 MB  다른 조사 데이터
#   __pycache__                  파이썬 캐시
#
# 푸는 쪽은 압축 안의 `embrain` 폴더를 통째로 작업 폴더로 삼는다 —
# 스킬과 도구가 서로를 상대 경로로 찾으므로 구조를 바꾸면 안 된다.
# 자세한 것은 압축 안의 `README_다른PC에서_시작하기.md` 에 있다.

param(
  [string]$Project = '',
  [string]$OutDir  = "$env:USERPROFILE\Downloads"
)
$ErrorActionPreference = 'Stop'

# report-factory\engine\tools\pack_skill.ps1 → embrain
$root  = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..')).Path
$stage = Join-Path $env:TEMP ('packskill_' + [Guid]::NewGuid().ToString('N').Substring(0,8))
$core  = Join-Path $stage 'embrain'

if (-not (Test-Path (Join-Path $root '.claude\skills'))) {
  throw "작업 루트를 찾지 못했습니다: $root (.claude\skills 가 없습니다)"
}

New-Item -ItemType Directory -Force -Path "$core\.claude", "$core\report-factory\engine" | Out-Null

Copy-Item (Join-Path $root '.claude\skills') "$core\.claude\skills" -Recurse
Copy-Item (Join-Path $root '.claude\agents') "$core\.claude\agents" -Recurse
Copy-Item (Join-Path $root '.claude\settings.json') "$core\.claude\settings.json" -ErrorAction SilentlyContinue

foreach ($d in 'tools','knowledge','learning','templates') {
  Copy-Item (Join-Path $root "report-factory\engine\$d") "$core\report-factory\engine\$d" -Recurse
}
Get-ChildItem "$core\report-factory\engine" -Recurse -Directory -Filter '__pycache__' |
  ForEach-Object { Remove-Item -Recurse -Force $_.FullName }

foreach ($f in 'README.md','INSTRUCTIONS.md') {
  $p = Join-Path $root "report-factory\$f"
  if (Test-Path $p) { Copy-Item $p "$core\report-factory\$f" }
}
Get-ChildItem (Join-Path $root 'report-factory') -Filter 'HANDOVER*.md' -File -ErrorAction SilentlyContinue |
  ForEach-Object { Copy-Item $_.FullName "$core\report-factory\$($_.Name)" }
Copy-Item (Join-Path $root 'report-factory\handover') "$core\report-factory\handover" -Recurse -ErrorAction SilentlyContinue

# 설치 안내 — 원본이 있으면 그대로, 없으면 최소 안내를 새로 쓴다
$readmeSrc = Join-Path $root 'report-factory\engine\templates\README_다른PC에서_시작하기.md'
if (Test-Path $readmeSrc) {
  Copy-Item $readmeSrc "$core\README_다른PC에서_시작하기.md"
} else {
  Set-Content -Path "$core\README_다른PC에서_시작하기.md" -Encoding utf8 -Value @'
# 제작가이드 스킬 — 다른 PC 에서 시작하기

이 `embrain` 폴더를 통째로 작업 폴더로 삼고 그 폴더에서 Claude Code 를 연다.
스킬과 도구가 상대 경로로 서로를 찾으므로 구조를 바꾸지 않는다.

    pip install openpyxl python-pptx lxml Pillow PyYAML xlrd

모든 파이썬 실행에 `PYTHONUTF8=1` 을 붙인다. 안 붙이면 cp949 콘솔에서 한글이 깨진다.

먼저 읽을 것: report-factory/HANDOVER*.md → engine/learning/RULES.md → engine/knowledge/
'@
}

$projName = ''
if ($Project) {
  $src = Join-Path $root "report-factory\projects\$Project"
  if (-not (Test-Path $src)) { throw "조사 폴더가 없습니다: $src" }
  New-Item -ItemType Directory -Force -Path "$core\report-factory\projects" | Out-Null
  Copy-Item $src "$core\report-factory\projects\$Project" -Recurse
  $projName = '+' + ($Project -replace '^\d{4}_','' -replace '_.*$','') + '작업'
} else {
  New-Item -ItemType Directory -Force -Path "$core\report-factory\projects" | Out-Null
  Set-Content -Path "$core\report-factory\projects\_여기에_조사폴더를_넣는다.txt" -Encoding utf8 -Value @'
조사 폴더는 용량 때문에 이 압축본에서 뺐다.
진행 중인 작업을 이어가려면 원래 PC 의 report-factory\projects\<조사폴더> 를 이 자리에 복사한다.
새 조사를 시작하는 것이라면 이 파일은 지워도 된다.
'@
}

$stamp = Get-Date -Format 'yyMMdd'
$dest  = Join-Path $OutDir "제작가이드_스킬$projName`_$stamp.zip"
if (Test-Path $dest) { Remove-Item -Force $dest }
Compress-Archive -Path $core -DestinationPath $dest -CompressionLevel Optimal
Remove-Item -Recurse -Force $stage

'{0}  {1:N1} MB' -f $dest, ((Get-Item $dest).Length / 1MB)
