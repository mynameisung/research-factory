# =========================================================================
# build.ps1 — src\*.bas 와 customUI\customUI14.xml 로 배포용 추가 기능을
#             out\통계표전처리.xlam 으로 굽는다.
#
# 이 스크립트를 돌리는 사람(만드는 쪽)만 아래 설정이 필요하다.
#   엑셀 → 파일 → 옵션 → 보안 센터 → 보안 센터 설정 → 매크로 설정
#        → 「VBA 프로젝트 개체 모델에 안전하게 액세스」 체크
# 완성된 .xlam 을 받아 쓰는 동료에게는 이 설정이 필요 없다.
#
#   powershell -ExecutionPolicy Bypass -File build.ps1
# =========================================================================
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$src  = Join-Path $root 'src'
$ui   = Join-Path $root 'customUI\customUI14.xml'
$out  = Join-Path $root 'out'
$xlam = Join-Path $out '통계표전처리.xlam'

if (-not (Test-Path $out)) { New-Item -ItemType Directory -Path $out | Out-Null }

# --- 0. VBA 프로젝트 접근 권한 확인 --------------------------------------
$ok = $false
foreach ($v in '16.0','15.0') {
    $k = "HKCU:\Software\Microsoft\Office\$v\Excel\Security"
    if (Test-Path $k) {
        $p = Get-ItemProperty $k -Name AccessVBOM -ErrorAction SilentlyContinue
        if ($p -and $p.AccessVBOM -eq 1) { $ok = $true }
    }
}
if (-not $ok) {
    Write-Host ''
    Write-Host '[중단] 엑셀이 VBA 프로젝트를 열어 주지 않는 상태입니다.' -ForegroundColor Yellow
    Write-Host '엑셀에서 다음을 한 번 켠 뒤 이 스크립트를 다시 실행해 주세요.'
    Write-Host '  파일 → 옵션 → 보안 센터 → 보안 센터 설정 → 매크로 설정'
    Write-Host '  → [VBA 프로젝트 개체 모델에 안전하게 액세스] 체크 → 확인 → 엑셀 종료'
    Write-Host ''
    exit 1
}

# --- 1. 한글 주석이 깨지지 않도록 소스를 CP949 로 옮겨 둔다 ---------------
$tmp = Join-Path $env:TEMP ('xlam_build_' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp | Out-Null
$cp949 = [Text.Encoding]::GetEncoding(949)
$bas = @()
# CP949 에 없는 글자는 저장할 때 '?' 로 바뀌어 안내문이 깨진다. 뜻이 같은
# 글자로 미리 바꿔 둔다 (엠대시·빼기표·홑따옴표 따위).
$fold = @{ [char]0x2014='-'; [char]0x2013='-'; [char]0x2212='-';
           [char]0x2018="'"; [char]0x2019="'"; [char]0x201C='"'; [char]0x201D='"' }
foreach ($f in Get-ChildItem $src -Filter *.bas | Sort-Object Name) {
    $dst = Join-Path $tmp $f.Name
    $text = [IO.File]::ReadAllText($f.FullName, [Text.Encoding]::UTF8)
    foreach ($k in $fold.Keys) { $text = $text.Replace($k, $fold[$k]) }
    # 그래도 남은 글자가 있으면 굽기 전에 알려 준다.
    $lost = ($cp949.GetString($cp949.GetBytes($text)).ToCharArray() |
             Where-Object { $_ -eq '?' }).Count - ($text.ToCharArray() | Where-Object { $_ -eq '?' }).Count
    if ($lost -gt 0) { Write-Host "  [주의] $($f.Name) 에 CP949 로 옮길 수 없는 글자 $lost 자" -ForegroundColor Yellow }
    [IO.File]::WriteAllText($dst, $text, $cp949)
    $bas += $dst
}
if ($bas.Count -eq 0) { throw "src 폴더에 .bas 파일이 없습니다: $src" }

# --- 1-b. 굽기 전 문법 점검 ---------------------------------------------
# VBA 는 한 문장에 줄 이음(_) 을 25개까지만, 한 줄에 1023자까지만 허용한다.
# 이 한도를 넘으면 엑셀은 이유를 말하지 않고 모듈 넣기를 거절한다.
$bad = @()
foreach ($f in Get-ChildItem $src -Include *.bas,*.code.txt -Recurse | Sort-Object Name) {
    $lines = [IO.File]::ReadAllLines($f.FullName, [Text.Encoding]::UTF8)
    $run = 0; $startLine = 0
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i].Length -gt 1000) {
            $bad += "$($f.Name):$($i+1)  한 줄이 $($lines[$i].Length)자 (1023자 한도)"
        }
        if ($lines[$i] -match ' _\s*$') {
            if ($run -eq 0) { $startLine = $i + 1 }
            $run++
        } else {
            if ($run -gt 24) { $bad += "$($f.Name):$startLine  줄 이음 $run 개 (25개 한도)" }
            $run = 0
        }
    }
    if ($run -gt 24) { $bad += "$($f.Name):$startLine  줄 이음 $run 개 (25개 한도)" }
}
if ($bad.Count -gt 0) {
    Write-Host ''
    Write-Host '[중단] VBA 한도를 넘는 자리가 있습니다.' -ForegroundColor Yellow
    $bad | ForEach-Object { Write-Host "  $_" }
    Write-Host '문장을 나누어 이어 붙이면 해결됩니다.'
    exit 1
}
# --- 1-c. 없는 이름을 부르는 자리 점검 -----------------------------------
# VBA 는 필요할 때에야 컴파일한다. 그래서 있지도 않은 함수를 부르는 줄이 있어도
# 빌드는 멀쩡히 통과하고, 사람이 그 단추를 누르는 순간에야 오류가 드러난다.
# 더 나쁜 것은 자동화로 띄운 엑셀에서는 그 오류 창이 보이지 않은 채 응답만
# 멎는다는 점이다(modInsight.IsInsightColor 사건). 모듈 사이의 한정 호출을
# 정의와 대조해 그런 줄을 빌드에서 미리 잡는다.
$defs = @{}
$srcFiles = Get-ChildItem $src -Include *.bas,*.code.txt -Recurse | Sort-Object Name
foreach ($f in $srcFiles) {
    $text = [IO.File]::ReadAllText($f.FullName, [Text.Encoding]::UTF8)
    $mn = [regex]::Match($text, 'Attribute VB_Name = "([^"]+)"')
    if (-not $mn.Success) { continue }
    $mod = $mn.Groups[1].Value
    if (-not $defs.ContainsKey($mod)) {
        $defs[$mod] = New-Object 'System.Collections.Generic.HashSet[string]'
    }
    foreach ($m in [regex]::Matches($text, '(?m)^(?<vis>Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(?:Sub|Function|Property\s+\w+)\s+(?<nm>\w+)')) {
        if ($m.Groups['vis'].Value.Trim() -eq 'Private') { continue }
        [void]$defs[$mod].Add($m.Groups['nm'].Value.ToLower())
    }
    foreach ($m in [regex]::Matches($text, '(?m)^Public\s+(?:Const\s+)?(?<nm>\w+)')) {
        [void]$defs[$mod].Add($m.Groups['nm'].Value.ToLower())
    }
    foreach ($m in [regex]::Matches($text, '(?m)^(?:Public|Private)\s+Type\s+(?<nm>\w+)')) {
        [void]$defs[$mod].Add($m.Groups['nm'].Value.ToLower())
    }
}
$missing = @()
foreach ($f in $srcFiles) {
    $lines = [IO.File]::ReadAllLines($f.FullName, [Text.Encoding]::UTF8)
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        if ($line.TrimStart().StartsWith("'")) { continue }
        $code = $line.Split("'")[0]
        foreach ($m in [regex]::Matches($code, '\b(?<mod>mod[A-Z]\w*)\.(?<nm>\w+)')) {
            $mod = $m.Groups['mod'].Value
            $nm  = $m.Groups['nm'].Value
            if (-not $defs.ContainsKey($mod)) {
                $missing += "$($f.Name):$($i+1)  $mod.$nm  - 그런 모듈이 없습니다"
            } elseif (-not $defs[$mod].Contains($nm.ToLower())) {
                $missing += "$($f.Name):$($i+1)  $mod.$nm  - 그 모듈에 그런 이름이 없거나 Private 입니다"
            }
        }
    }
}
if ($missing.Count -gt 0) {
    Write-Host ''
    Write-Host '[중단] 없는 이름을 부르는 자리가 있습니다.' -ForegroundColor Yellow
    $missing | ForEach-Object { Write-Host "  $_" }
    Write-Host '이름을 고치거나, 부르려는 함수를 Public 으로 만들어 주십시오.'
    exit 1
}

Write-Host "[0/2] 문법 점검 통과 ($($bas.Count) 개 모듈)" -ForegroundColor Green

# --- 2. 엑셀을 띄워 모듈을 넣고 .xlam 으로 저장한다 -----------------------
if (Test-Path $xlam) { Remove-Item $xlam -Force }
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Add()
    foreach ($b in $bas) {
        try { $wb.VBProject.VBComponents.Import($b) | Out-Null }
        catch { throw ("모듈을 넣지 못했습니다: " + (Split-Path $b -Leaf) + " — " + $_.Exception.Message) }
    }
    # 설정 팝업은 폼이라 파일로 가져올 수 없다. 여기서 직접 만들고 코드를 붙인다.
    # 배치를 고칠 곳은 이 블록, 동작을 고칠 곳은 src/frmSettings.code.txt 다.
    $frm = $wb.VBProject.VBComponents.Add(3)      # 3 = UserForm
    $frm.Name = 'frmSettings'
    $frm.Properties.Item('Caption').Value = '통계표 전처리 - 설정'
    $frm.Properties.Item('Width').Value   = 444
    $frm.Properties.Item('Height').Value  = 330

    function Put($kind, $name, $left, $top, $width, $height, $caption) {
        $c = $frm.Designer.Controls.Add("Forms.$kind.1", $name, $true)
        $c.Left = $left; $c.Top = $top; $c.Width = $width; $c.Height = $height
        if ($null -ne $caption) { try { $c.Caption = $caption } catch {} }
        return $c
    }

    Put Label   'lblKeep'   12  10 400 14 '보존 배너군   (쉼표로 구분 · 비워 두면 Base 기준만 적용합니다)' | Out-Null
    Put TextBox 'txtKeep'   12  26 408 18 $null | Out-Null

    Put Label   'lblBase'   12  54  70 14 'Base 하한' | Out-Null
    Put TextBox 'txtBase'   86  52  50 18 $null | Out-Null
    Put Label   'lblBaseHint' 144 54 276 26 '배너 항목의 절반 이상이 이 모수를 넘어야 그 배너군을 남깁니다.' | Out-Null

    Put Label    'lblMode'  12  88  70 14 '이상치 기준' | Out-Null
    $cbo = Put ComboBox 'cboMode'  86  86 140 18 $null
    $cbo.Style = 2                                  # 목록에서만 고르게 한다
    Put Label   'lblVal'   240  88  50 14 '기준값' | Out-Null
    Put TextBox 'txtVal'   288  86  60 18 $null | Out-Null
    Put Label   'lblHint'   12 110 408 28 '' | Out-Null

    Put CheckBox 'chkBlank'  12 146 408 16 '빈 칸을 0 으로 본다   (끄면 데이터가 없는 칸으로 보고 세지 않습니다)' | Out-Null
    Put CheckBox 'chkDisp'   12 166 408 16 '표시값으로 판정한다   (저장값은 반올림되지 않은 수입니다)' | Out-Null
    Put CheckBox 'chkBackup' 12 186 408 16 '삭제 전 백업 시트를 만든다   (매크로는 Ctrl+Z 로 되돌아가지 않습니다)' | Out-Null

    Put Label 'lblNote' 12 212 408 28 '판정은 표 한 장이 아니라 통계표 전체를 놓고 배너군 이름마다 한 번만 내립니다. 어떤 표에는 그 배너군의 항목이 하나만 실려 오기 때문입니다.' | Out-Null

    Put CommandButton 'btnCatalog'  12 250 120 26 '배너군 목록 보기' | Out-Null
    Put CommandButton 'btnReset'   140 250  70 26 '초기화'          | Out-Null
    Put CommandButton 'btnSave'    258 250  80 26 '저장'            | Out-Null
    Put CommandButton 'btnCancel'  342 250  78 26 '취소'            | Out-Null

    $frmCode = [IO.File]::ReadAllText((Join-Path $src 'frmSettings.code.txt'), [Text.Encoding]::UTF8)
    $frm.CodeModule.AddFromString($frmCode)

    # --- 처음 열었을 때 한 번 뜨는 안내 --------------------------------
    $wel = $wb.VBProject.VBComponents.Add(3)
    $wel.Name = 'frmWelcome'
    $wel.Properties.Item('Caption').Value = '통계표 전처리'
    $wel.Properties.Item('Width').Value   = 432
    $wel.Properties.Item('Height').Value  = 350

    function PutW($kind, $name, $left, $top, $width, $height, $caption) {
        $c = $wel.Designer.Controls.Add("Forms.$kind.1", $name, $true)
        $c.Left = $left; $c.Top = $top; $c.Width = $width; $c.Height = $height
        if ($null -ne $caption) { try { $c.Caption = $caption } catch {} }
        return $c
    }

    # 줄마다 라벨을 따로 둔다. 라벨 하나에 다 담으면 굵게 강조할 자리를 고를 수
    # 없기 때문이다. 글꼴과 크기는 폼이 열릴 때 VBA 가 입힌다.
    PutW Label 'lblTitle' 18 16 384 22 '통계표를 제작가이드에 옮기기 전에 인사이트를 추출합니다' | Out-Null

    $steps = @(
        @('1.  [가져오기] 로 통계표 파일을 고릅니다.',                    '표 시트를 작업 사본으로 떠서 열기 때문에 원본 파일은 바뀌지 않습니다.'),
        @('2.  [설정] 에서 보고서에 쓸 [보존 배너군] 을 적습니다.',       '이름을 모르면 [검토] → [배너군 목록] 을 먼저 눌러 보세요.'),
        @('3.  [한 번에 실행] 을 누릅니다.',                              '배너 선정 → 데이터막대 → 인사이트를 잇달아 합니다.'),
        @('4.  회색으로 칠해진 배너를 확인하고 [제외 배너 삭제] 를 누릅니다.', '매크로가 한 일은 Ctrl+Z 로 되돌아가지 않으니 눈으로 보고 누르십시오.'),
        @('5.  [- 로 채우기] 로 빈 칸을 메운 뒤 [내보내기] 로 저장합니다.',   '통계표는 값이 0 인 칸을 비워서 보내므로 맨 마지막에 메웁니다.')
    )
    $y = 52
    $i = 0
    foreach ($st in $steps) {
        PutW Label ("lblH$i") 18 $y 384 17 $st[0] | Out-Null
        PutW Label ("lblD$i") 34 ($y + 18) 372 16 $st[1] | Out-Null
        $y += 40
        $i++
    }

    PutW CheckBox 'chkHide' 18 262 200 18 '다시 보지 않기' | Out-Null
    PutW Label 'lblBy' 18 288 240 16 'powered by seo-yeon' | Out-Null
    PutW CommandButton 'btnStart' 300 258 100 30 '시작하기' | Out-Null
    $welCode = [IO.File]::ReadAllText((Join-Path $src 'frmWelcome.code.txt'), [Text.Encoding]::UTF8)
    $wel.CodeModule.AddFromString($welCode)

    # 폼에 컨트롤이 실제로 들어갔는지 굽기 전에 센다. 하나라도 비면 멈춘다.
    $nSet = @($frm.Designer.Controls).Count
    $nWel = @($wel.Designer.Controls).Count
    Write-Host ("  폼 확인 : frmSettings $nSet 개 · frmWelcome $nWel 개 컨트롤")
    if ($nSet -lt 10 -or $nWel -lt 13) { throw "폼 컨트롤이 모자랍니다 (설정 $nSet · 안내 $nWel)" }

    $wb.VBProject.Name = 'TablePrep'
    $wb.Title    = '통계표 전처리'
    $wb.Comments = '통계표 → 제작가이드 전처리 (배너 선정 · 데이터막대 · 인사이트 칠하기)'
    $wb.IsAddin  = $true
    $wb.SaveAs($xlam, 55)          # 55 = xlOpenXMLAddIn
    $wb.Close($false)
} finally {
    # 사람이 따로 열어 둔 문서가 남아 있으면 엑셀을 끄지 않는다.
    if ($xl.Workbooks.Count -eq 0) { $xl.Quit() }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
}
Write-Host "[1/2] 모듈 $($bas.Count) 개를 넣어 저장했습니다." -ForegroundColor Green

# --- 3. 리본 XML 을 패키지에 끼워 넣는다 ---------------------------------
# .xlam 은 OPC(zip) 이라 customUI 파트와 그 관계 항목을 직접 넣어 준다.
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression

$work = Join-Path $env:TEMP ('xlam_ui_' + [Guid]::NewGuid().ToString('N'))
[IO.Compression.ZipFile]::ExtractToDirectory($xlam, $work)

$uiDir = Join-Path $work 'customUI'
New-Item -ItemType Directory -Path $uiDir -Force | Out-Null
Copy-Item $ui (Join-Path $uiDir 'customUI14.xml') -Force

$relsPath = Join-Path $work '_rels\.rels'
[xml]$rels = Get-Content $relsPath -Raw -Encoding UTF8
$ns = 'http://schemas.openxmlformats.org/package/2006/relationships'
$has = $false
foreach ($r in $rels.Relationships.Relationship) {
    if ($r.Target -eq 'customUI/customUI14.xml') { $has = $true }
}
if (-not $has) {
    $node = $rels.CreateElement('Relationship', $ns)
    $node.SetAttribute('Id', 'rIdCustomUI14')
    $node.SetAttribute('Type', 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility')
    $node.SetAttribute('Target', 'customUI/customUI14.xml')
    [void]$rels.Relationships.AppendChild($node)
    $rels.Save($relsPath)
}

# [Content_Types].xml 에 xml 확장자가 없으면 넣어 준다.
$ctPath = Join-Path $work '[Content_Types].xml'
[xml]$ct = Get-Content -LiteralPath $ctPath -Raw -Encoding UTF8
$hasXml = $false
foreach ($d in $ct.Types.Default) { if ($d.Extension -eq 'xml') { $hasXml = $true } }
if (-not $hasXml) {
    $ctNs = 'http://schemas.openxmlformats.org/package/2006/content-types'
    $d = $ct.CreateElement('Default', $ctNs)
    $d.SetAttribute('Extension', 'xml')
    $d.SetAttribute('ContentType', 'application/xml')
    [void]$ct.Types.PrependChild($d)
    $ct.Save($ctPath)
}

Remove-Item $xlam -Force
[IO.Compression.ZipFile]::CreateFromDirectory($work, $xlam, [IO.Compression.CompressionLevel]::Optimal, $false)
Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "[2/2] 리본 XML 을 끼워 넣었습니다." -ForegroundColor Green

# --- 4. 구운 파일을 다시 열어 확인한다 -----------------------------------
# 폼은 저장하는 과정에서 컨트롤을 잃는 수가 있다. 만드는 자리에서 세어 본 것으로는
# 모자라고, 실제로 배포될 파일을 열어 확인해야 한다.
$xl2 = New-Object -ComObject Excel.Application
$xl2.Visible = $false
$xl2.DisplayAlerts = $false
$verifyErr = @()
# 원본 경로를 그대로 열지 않고 임시 사본을 연다. 같은 파일이 다른 엑셀에 물려
# 있으면 이름이 겹쳐 열리지 않는데, 검증이 거기서 막힐 이유는 없다.
$probe = Join-Path $env:TEMP ('xlam_verify_' + [Guid]::NewGuid().ToString('N') + '.xlam')
Copy-Item $xlam $probe -Force
try {
    $wb2 = $xl2.Workbooks.Open($probe)
    # 폼 안은 밖에서 들여다봐서는 제대로 보이지 않는다. 애드인 쪽 함수를 불러
    # 폼을 메모리에 올린 뒤 컨트롤을 세게 한다.
    $nm = Split-Path $probe -Leaf
    $chk = $xl2.Run("'$nm'!FormCheck")
    Write-Host "  확인 : $chk"
    if ($chk -notmatch 'frmSettings=(\d+)' -or [int]$matches[1] -lt 10) { $verifyErr += "설정 폼의 컨트롤이 모자랍니다 ($chk)" }
    if ($chk -notmatch 'frmWelcome=(\d+)'  -or [int]$matches[1] -lt 13) { $verifyErr += "안내 폼의 컨트롤이 모자랍니다 ($chk)" }
    
    # VBA 는 필요할 때만 컴파일한다. 모듈을 하나씩 건드려 전부 컴파일시킨다.
    $mods = $xl2.Run("'$nm'!CompileAll")
    Write-Host "  확인 : 컴파일 통과 - $mods"
    if (($mods -split ' ').Count -lt 12) { $verifyErr += "컴파일된 모듈이 모자랍니다 ($mods)" }
    $wb2.Close($false)
} catch {
    $verifyErr += $_.Exception.Message
} finally {
    if ($xl2.Workbooks.Count -eq 0) { $xl2.Quit() }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl2)
    Remove-Item $probe -Force -ErrorAction SilentlyContinue
}
if ($verifyErr.Count -gt 0) {
    Write-Host ''
    Write-Host '[중단] 구운 파일에 문제가 있습니다.' -ForegroundColor Yellow
    $verifyErr | ForEach-Object { Write-Host "  $_" }
    exit 1
}
Write-Host ''
Write-Host "완성 : $xlam" -ForegroundColor Cyan
Write-Host '동료에게는 이 파일 하나와 README 의 설치 안내만 전달하면 됩니다.'
