Attribute VB_Name = "modWelcome"
'==========================================================================
' modWelcome — 처음 열었을 때 뜨는 안내
'
' 추가 기능을 열 때마다 창이 뜨면 성가시므로, [다시 보지 않기] 를 켠 뒤로는
' 뜨지 않는다. 기억은 레지스트리에 남으므로 사람마다 따로 잡힌다.
'
' 자동 검증처럼 사람이 없는 자리에서 열릴 때는 창이 뜨면 안 되므로,
' 조용히 실행 중이면 건너뛴다.
'==========================================================================
Option Explicit

' 이 세션에서 이미 띄웠는지 기억한다.
Private gShownThisSession As Boolean

'--------------------------------------------------------------------------
' 리본 단추를 처음 누를 때 한 번 띄운다.
'
' 예전에는 추가 기능을 여는 순간(Workbook_Open) 띄웠는데, 그러면 사람이 없는
' 자리에서 엑셀을 여는 경우 - 다른 프로그램이 엑셀을 뒤에서 돌리거나 자동 검증을
' 하는 경우 - 창이 뜬 채로 통째로 멈춰 버린다. 단추를 눌렀다는 것은 사람이 앞에
' 있다는 뜻이므로, 그 자리로 옮겼다.
'--------------------------------------------------------------------------
Public Sub ShowWelcome()
    On Error Resume Next
    If gShownThisSession Then Exit Sub
    If modConfig.IsSilent() Then Exit Sub
    gShownThisSession = True

    ' 리본은 엑셀 2010(버전 14) 부터 쓰는 방식으로 짜여 있다.
    If Val(Application.Version) < 14 Then Exit Sub
    If modConfig.WelcomeSeen() Then Exit Sub
    frmWelcome.Show
End Sub
Public Sub ShowWelcomeAgain()
    On Error Resume Next
    modConfig.SetWelcomeSeen False
    gShownThisSession = True
    frmWelcome.Show
End Sub

'--------------------------------------------------------------------------
' 빌드가 구운 파일을 확인할 때 쓴다. 폼을 화면에 띄우지 않고 메모리에만 올려
' 컨트롤 수를 센다. 저장된 파일을 밖에서 들여다보는 방법으로는 폼 안이 제대로
' 보이지 않기 때문에, 실제로 올려 보는 것이 유일하게 믿을 만한 확인이다.
'--------------------------------------------------------------------------
'--------------------------------------------------------------------------
' 모든 모듈을 한 번씩 건드려 컴파일을 강제한다. 빌드가 이 함수를 불러
' 모듈 하나라도 문법이 틀려 있으면 그 자리에서 멈춘다.
'--------------------------------------------------------------------------
Public Function CompileAll() As String
    Dim s As String
    s = s & " " & modBanner.Ping()
    s = s & " " & modBars.Ping()
    s = s & " " & modConfig.Ping()
    s = s & " " & modDash.Ping()
    s = s & " " & modInsight.Ping()
    s = s & " " & modOpen.Ping()
    s = s & " " & modReport.Ping()
    s = s & " " & modRibbon.Ping()
    s = s & " " & modRun.Ping()
    s = s & " " & modTable.Ping()
    s = s & " " & modUtil.Ping()
    s = s & " " & modWelcome.Ping()
    CompileAll = Trim$(s)
End Function

Public Function FormCheck() As String
    Dim s As String
    On Error Resume Next
    Load frmSettings
    s = "frmSettings=" & frmSettings.Controls.Count
    Unload frmSettings
    Load frmWelcome
    s = s & " frmWelcome=" & frmWelcome.Controls.Count
    ' 제목 줄이 실제로 채워졌는지도 함께 본다.
    s = s & " 제목=" & Len(frmWelcome.lblTitle.Caption) & "자"
    Unload frmWelcome
    On Error GoTo 0
    FormCheck = s
End Function

' 빌드가 모든 모듈을 실제로 컴파일하도록 부르는 자리다. VBA 는 필요할 때만
' 컴파일하기 때문에, 한 번도 불리지 않은 모듈의 오류는 사람이 그 단추를 누르는
' 순간에야 드러난다. 이 함수가 그 구멍을 막는다.
Public Function Ping() As String
    Ping = "modWelcome"
End Function
