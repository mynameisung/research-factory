Attribute VB_Name = "modRibbon"
'==========================================================================
' modRibbon — 리본 탭 [통계표 전처리] 의 콜백
'
' 리본 구성은 customUI/customUI14.xml 에 있고, 이 모듈은 그 XML 이 부르는
' 함수들을 담는다. 함수 이름을 바꾸면 XML 도 함께 고쳐야 한다.
'
' 콜백 안에서 오류가 나면 엑셀이 리본을 통째로 꺼 버리는 수가 있다. 그래서
' 모든 콜백을 On Error 로 감싸고, 사람에게는 무엇이 잘못됐는지 알린다.
'==========================================================================
Option Explicit

Private gRibbon As IRibbonUI

Public Sub OnRibbonLoad(ribbon As IRibbonUI)
    Set gRibbon = ribbon
End Sub

Private Sub Oops(ByVal what As String)
    modUtil.EndWork          ' 화면 갱신과 자동 계산을 되돌린다
    modConfig.SetSilent False
    MsgBox what & " 를 실행하다가 멈췄습니다." & vbCrLf & vbCrLf & _
           Err.Number & " : " & Err.Description & vbCrLf & vbCrLf & _
           "통계표 시트가 맞는지, 시트가 보호되어 있지 않은지 확인해 주세요.", _
           vbExclamation, "통계표 전처리"
End Sub

'-- 파일 ------------------------------------------------------------------
Public Sub cbOpenTable(control As IRibbonControl)
    On Error GoTo E
    modWelcome.ShowWelcome
    modOpen.OpenTableFile
    Exit Sub
E: Oops "통계표 가져오기"
End Sub

Public Sub cbExportXlsx(control As IRibbonControl)
    On Error GoTo E
    modWelcome.ShowWelcome
    modOpen.ExportXlsx
    Exit Sub
E: Oops "xlsx 로 내보내기"
End Sub

Public Sub cbRunAll(control As IRibbonControl)
    On Error GoTo E
    modWelcome.ShowWelcome
    modRun.RunAll
    Exit Sub
E: Oops "한 번에 실행"
End Sub

'-- 1. 배너 선정 ----------------------------------------------------------
Public Sub cbSelectBanners(control As IRibbonControl)
    On Error GoTo E
    modWelcome.ShowWelcome
    modBanner.SelectBanners
    Exit Sub
E: Oops "배너 선정"
End Sub

Public Sub cbDeleteShaded(control As IRibbonControl)
    On Error GoTo E
    modWelcome.ShowWelcome
    modBanner.DeleteShadedBanners
    Exit Sub
E: Oops "음영 배너 일괄삭제"
End Sub

Public Sub cbClearShade(control As IRibbonControl)
    On Error GoTo E
    modWelcome.ShowWelcome
    modBanner.ClearAllShade
    Exit Sub
E: Oops "음영 지우기"
End Sub

Public Sub cbBannerCatalog(control As IRibbonControl)
    On Error GoTo E
    modWelcome.ShowWelcome
    modReport.BannerCatalog
    Exit Sub
E: Oops "배너군 목록"
End Sub

'-- 2. 데이터막대 ---------------------------------------------------------
Public Sub cbBarsTotal(control As IRibbonControl)
    On Error GoTo E
    modWelcome.ShowWelcome
    modBars.BarsTotal
    Exit Sub
E: Oops "데이터막대(전체 열)"
End Sub

Public Sub cbBarsBanner(control As IRibbonControl)
    On Error GoTo E
    modWelcome.ShowWelcome
    modBars.BarsBanner
    Exit Sub
E: Oops "데이터막대(배너 열)"
End Sub

Public Sub cbBarsBoth(control As IRibbonControl)
    On Error GoTo E
    modWelcome.ShowWelcome
    modBars.BarsBoth
    Exit Sub
E: Oops "데이터막대"
End Sub

Public Sub cbClearBars(control As IRibbonControl)
    On Error GoTo E
    modWelcome.ShowWelcome
    modBars.ClearBars
    Exit Sub
E: Oops "막대 지우기"
End Sub

'-- 3. 인사이트 -----------------------------------------------------------
Public Sub cbPaintInsight(control As IRibbonControl)
    On Error GoTo E
    modWelcome.ShowWelcome
    modInsight.PaintInsight
    Exit Sub
E: Oops "형광 칠하기"
End Sub

Public Sub cbClearInsight(control As IRibbonControl)
    On Error GoTo E
    modWelcome.ShowWelcome
    modInsight.ClearInsight
    Exit Sub
E: Oops "표시 지우기"
End Sub

'-- 4. 빈 칸 -------------------------------------------------------------
Public Sub cbFillDashes(control As IRibbonControl)
    On Error GoTo E
    modWelcome.ShowWelcome
    modDash.FillDashes
    Exit Sub
E: Oops "빈 칸 - 로 채우기"
End Sub

Public Sub cbClearDashes(control As IRibbonControl)
    On Error GoTo E
    modWelcome.ShowWelcome
    modDash.ClearDashes
    Exit Sub
E: Oops "빈 칸 되돌리기"
End Sub

'-- 검토 ------------------------------------------------------------------
Public Sub cbSettings(control As IRibbonControl)
    On Error GoTo E
    modWelcome.ShowWelcome
    frmSettings.Show
    Exit Sub
E: Oops "설정"
End Sub

'-- 도움말 ----------------------------------------------------------------
Public Sub cbWelcome(control As IRibbonControl)
    On Error GoTo E
    modWelcome.ShowWelcome
    modWelcome.ShowWelcomeAgain
    Exit Sub
E: Oops "처음 안내"
End Sub

' VBA 는 한 문장에 줄 이음(_) 을 25개까지만 허용한다. 안내문이 길어서
' 문장을 나누어 이어 붙인다.
Public Sub cbHelp(control As IRibbonControl)
    Dim s As String, LN As String, BAR As String, THIN As String
    LN = vbCrLf
    BAR = String$(52, "=")
    THIN = String$(52, "-")

    s = "통계표 전처리" & LN & BAR & LN
    s = s & "DP 원본 통계표의 Table 시트에서 실행합니다." & LN & LN

    s = s & "쓰는 순서" & LN & THIN & LN
    s = s & "  1) [가져오기] 로 파일을 고릅니다. 표 시트를 작업 사본으로" & LN
    s = s & "     떠서 열기 때문에 원본 파일은 바뀌지 않습니다." & LN
    s = s & "  2) [배너군 목록] 을 눌러 이 통계표에 어떤 배너군이 있는지 봅니다." & LN
    s = s & "     맨 아래 한 줄을 복사해 [보존 배너군] 칸에 붙여 넣고, 보고서에" & LN
    s = s & "     쓰지 않을 이름만 지웁니다." & LN
    s = s & "  3) [한 번에 실행] 을 누르면 아래 네 가지를 잇달아 합니다." & LN
    s = s & "  4) 회색으로 칠해진 배너를 눈으로 확인한 뒤" & LN
    s = s & "     [제외 배너 삭제] 를 누릅니다." & LN
    s = s & "  5) [내보내기] 로 저장합니다." & LN & LN

    s = s & "네 가지가 하는 일" & LN & THIN & LN
    s = s & "1. 배너 선정" & LN
    s = s & "   보존 목록에 들면서, 배너 항목의 절반 이상이 Base 하한을 넘는" & LN
    s = s & "   배너군만 남기고 나머지는 회색으로 칠합니다. 어느 배너를 쓸지는" & LN
    s = s & "   연구팀이 정하고(보존 목록), 읽을 수 있는지는 모수가 정합니다." & LN
    s = s & "   판정은 표 한 장이 아니라 통계표 전체를 놓고 배너군 이름마다" & LN
    s = s & "   한 번만 내립니다. 어떤 표에는 그 배너군의 항목이 하나만 실려" & LN
    s = s & "   오기 때문에, 표 단위로 정하면 같은 배너가 남았다 없어졌다 합니다." & LN & LN
    s = s & "2. 데이터막대" & LN
    s = s & "   전체 열은 파랑, 배너 열은 회색입니다. 길이 기준을 0~100 으로" & LN
    s = s & "   고정해 전체의 분포와 배너의 분포를 견줄 수 있게 했습니다." & LN
    s = s & "   값이 작은 표(5점 척도 평균 등)는 그 열의 범위에 맞춥니다." & LN & LN
    s = s & "3. 인사이트" & LN
    s = s & "   전체 값에서 크게 벗어난 배너 값에 색을 칠합니다." & LN
    s = s & "   기준은 표준편차 배수와 절대값(%p) 가운데 고를 수 있습니다." & LN
    s = s & "   Base 하한에 못 미치는 열과 회색 배너 열은 칠하지 않습니다." & LN & LN
    s = s & "4. 빈 칸을 - 로 채우기" & LN
    s = s & "   통계표는 값이 0 인 칸을 비워서 보냅니다. 0 인지 묻지 않은 칸인지" & LN
    s = s & "   알 수 없으므로 - 로 채웁니다. - 는 글자라서 그 칸에는 막대가" & LN
    s = s & "   그려지지 않고 인사이트에서도 빠지므로 맨 마지막에 누릅니다." & LN & LN

    s = s & "알아 두실 것" & LN & THIN & LN
    s = s & "  · 매크로가 한 일은 Ctrl+Z 로 되돌아가지 않습니다." & LN
    s = s & "    설정의 [삭제 전 백업] 을 켜 두시기 바랍니다." & LN
    s = s & "  · 통계표는 값이 0 인 칸을 비워서 보냅니다. [빈 칸을 0 으로 본다]" & LN
    s = s & "    를 끄면 가장 큰 차이를 놓칠 수 있습니다." & LN
    s = s & "  · 판정은 화면에 보이는 표시값으로 합니다. 저장값은 반올림되지" & LN
    s = s & "    않은 수라 기준 경계에서 어긋나기 때문입니다." & LN & LN
    s = s & THIN & LN
    s = s & "powered by seo-yeon"

    MsgBox s, vbInformation, "통계표 전처리 사용법"
End Sub
'==========================================================================
Private Sub Refresh(ByVal id As String)
    On Error Resume Next
    If Not gRibbon Is Nothing Then gRibbon.InvalidateControl id
    On Error GoTo 0
End Sub

Private Sub RefreshAll()
    On Error Resume Next
    If Not gRibbon Is Nothing Then gRibbon.Invalidate
    On Error GoTo 0
End Sub

' 빌드가 모든 모듈을 실제로 컴파일하도록 부르는 자리다. VBA 는 필요할 때만
' 컴파일하기 때문에, 한 번도 불리지 않은 모듈의 오류는 사람이 그 단추를 누르는
' 순간에야 드러난다. 이 함수가 그 구멍을 막는다.
Public Function Ping() As String
    Ping = "modRibbon"
End Function
