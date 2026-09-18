Attribute VB_Name = "modConfig"
'==========================================================================
' modConfig — 리본에서 바꾼 설정을 담아 두고 다음에 열 때도 그대로 쓴다.
'
' 값은 레지스트리(HKCU\Software\VB and VBA Program Settings)에 저장되므로
' 통계표 파일을 바꿔 열어도 설정이 남아 있다. 프로젝트가 바뀌어 배너 이름이
' 달라지면 리본의 [보존 배너군] 칸만 다시 적으면 된다.
'==========================================================================
Option Explicit

Private Const APP_NAME As String = "통계표전처리"
Private Const SEC As String = "설정"

'--------------------------------------------------------------------------
' 조용히 실행하기 — 검증 스크립트가 대화상자에 막히지 않도록 둔 통로다.
' 사람이 리본에서 누를 때는 늘 False 이므로 평소 동작은 달라지지 않는다.
'--------------------------------------------------------------------------
Public gSilent As Boolean
Public gLastMessage As String

Public Sub Say(ByVal msg As String, Optional ByVal style As Long = vbInformation, _
               Optional ByVal title As String = "통계표 전처리")
    gLastMessage = msg
    If gSilent Then Exit Sub
    MsgBox msg, style, title
End Sub

' 확인을 묻는다. 조용히 실행할 때는 진행한 것으로 본다.
Public Function Ask(ByVal msg As String, Optional ByVal title As String = "통계표 전처리") As Boolean
    If gSilent Then
        Ask = True
    Else
        Ask = (MsgBox(msg, vbYesNo + vbExclamation, title) = vbYes)
    End If
End Function

Public Sub SetSilent(ByVal onOff As Boolean)
    gSilent = onOff
End Sub

Public Function IsSilent() As Boolean
    IsSilent = gSilent
End Function

' 조용히 실행했을 때 마지막으로 나왔을 안내문을 되돌려준다.
Public Function LastMessage() As String
    LastMessage = gLastMessage
End Function

'--------------------------------------------------------------------------
' 보고서에 쓸 배너군 이름 목록. 쉼표로 구분한다. 비우면 Base 관문만 적용된다.
'--------------------------------------------------------------------------
Public Function KeepGroups() As String
    KeepGroups = GetSetting(APP_NAME, SEC, "KeepGroups", "")
End Function

Public Sub SetKeepGroups(ByVal s As String)
    SaveSetting APP_NAME, SEC, "KeepGroups", s
End Sub

'--------------------------------------------------------------------------
' 배너 항목이 넘어야 하는 Base 하한. 기능 1·3 이 함께 쓴다.
'--------------------------------------------------------------------------
Public Function MinBase() As Long
    Dim v As String
    v = GetSetting(APP_NAME, SEC, "MinBase", "30")
    If IsNumeric(v) Then MinBase = CLng(Val(v))
    If MinBase <= 0 Then MinBase = 30
End Function

Public Sub SetMinBase(ByVal s As String)
    If IsNumeric(s) Then SaveSetting APP_NAME, SEC, "MinBase", CStr(CLng(Val(s)))
End Sub

'--------------------------------------------------------------------------
' 인사이트 판정 방식 — "STDEV"(표준편차) 또는 "ABS"(절대값 %p)
'--------------------------------------------------------------------------
Public Function InsightMode() As String
    InsightMode = GetSetting(APP_NAME, SEC, "InsightMode", "STDEV")
    If InsightMode <> "ABS" Then InsightMode = "STDEV"
End Function

Public Sub SetInsightMode(ByVal s As String)
    SaveSetting APP_NAME, SEC, "InsightMode", s
End Sub

'--------------------------------------------------------------------------
' 인사이트 기준값 — 표준편차 방식이면 σ 의 배수, 절대값 방식이면 %p
'--------------------------------------------------------------------------
Public Function InsightValue() As Double
    Dim v As String
    v = GetSetting(APP_NAME, SEC, "InsightValue", "")
    If Len(v) = 0 Or Not IsNumeric(v) Then
        InsightValue = IIf(InsightMode() = "ABS", 10#, 2#)
    Else
        InsightValue = CDbl(Val(v))
    End If
End Function

Public Sub SetInsightValue(ByVal s As String)
    If IsNumeric(s) Then SaveSetting APP_NAME, SEC, "InsightValue", CStr(CDbl(Val(s)))
End Sub

'==========================================================================
' 판정을 좌우하는 세 가지 규약. 셋 다 통계표의 성질에서 나온 것이라
' 기본값을 켜 두었고, 사정이 다른 표를 만나면 리본에서 끌 수 있다.
'==========================================================================

'--------------------------------------------------------------------------
' 빈 칸을 0 으로 본다 (기본 끔)
'
' 통계표에서 빈 칸은 두 가지를 뜻할 수 있다. 값이 0 이어서 비운 것일 수도 있고,
' 그 집단에는 그 보기를 물어보지 않아 데이터 자체가 없는 것일 수도 있다.
' 둘을 구분할 방법이 표 안에는 없으므로, 기본은 **세지 않는 쪽**으로 둔다.
' 데이터가 없는 칸을 0 으로 세면 없는 차이를 크게 만들어 내기 때문이다.
'
' 값이 0 인 것이 분명한 표(단일선택 보기 분포 등)에서는 이 항목을 켜면
' 「전체 20% 인데 이 집단은 0%」 같은 차이도 함께 잡힌다.
'--------------------------------------------------------------------------
Public Function BlankIsZero() As Boolean
    BlankIsZero = (GetSetting(APP_NAME, SEC, "BlankIsZero", "0") = "1")
End Function

Public Sub SetBlankIsZero(ByVal onOff As Boolean)
    SaveSetting APP_NAME, SEC, "BlankIsZero", IIf(onOff, "1", "0")
End Sub

'--------------------------------------------------------------------------
' 표시값으로 판정한다 (기본 켬)
'
' 셀에 저장된 값은 반올림되지 않은 부동소수다. 표시는 8.4 인데 저장값은
' 8.391608... 인 식이다. 보고서에 싣는 수치는 늘 서식이 적용된 값이므로,
' 인사이트 판정도 표시값으로 해야 사람이 화면에서 보는 것과 어긋나지 않는다.
'--------------------------------------------------------------------------
Public Function UseDisplayValue() As Boolean
    UseDisplayValue = (GetSetting(APP_NAME, SEC, "UseDisplayValue", "1") = "1")
End Function

Public Sub SetUseDisplayValue(ByVal onOff As Boolean)
    SaveSetting APP_NAME, SEC, "UseDisplayValue", IIf(onOff, "1", "0")
End Sub

'--------------------------------------------------------------------------
' 지우기 전에 백업 시트를 만든다 (기본 켬)
'
' 매크로가 한 일은 Ctrl+Z 로 되돌릴 수 없다. 배너를 지우기 직전의 시트를
' 한 장 떠 두면 잘못 지웠을 때 되돌릴 자리가 남는다.
'--------------------------------------------------------------------------
Public Function BackupBeforeDelete() As Boolean
    BackupBeforeDelete = (GetSetting(APP_NAME, SEC, "BackupBeforeDelete", "1") = "1")
End Function

Public Sub SetBackupBeforeDelete(ByVal onOff As Boolean)
    SaveSetting APP_NAME, SEC, "BackupBeforeDelete", IIf(onOff, "1", "0")
End Sub

'--------------------------------------------------------------------------
' 처음 열었을 때 뜨는 안내를 봤는지 기억한다.
'--------------------------------------------------------------------------
Public Function WelcomeSeen() As Boolean
    WelcomeSeen = (GetSetting(APP_NAME, SEC, "WelcomeSeen", "0") = "1")
End Function

Public Sub SetWelcomeSeen(ByVal onOff As Boolean)
    SaveSetting APP_NAME, SEC, "WelcomeSeen", IIf(onOff, "1", "0")
End Sub

'--------------------------------------------------------------------------
' 쉼표로 구분한 목록에 이름이 들어 있는지 본다. 앞뒤 공백은 무시한다.
'--------------------------------------------------------------------------
' 배너군 이름은 사람이 손으로 옮겨 적는 값이라 띄어쓰기와 구분자가 흔들린다.
' '관리 운영 방식' 과 '관리운영방식' 을 같은 것으로 보고, 쉼표·중점·세로줄을
' 모두 구분자로 받아 준다.
Public Function InList(ByVal name As String, ByVal listCsv As String) As Boolean
    Dim parts() As String, i As Long, target As String
    target = Squash(name)
    If Len(target) = 0 Then Exit Function
    listCsv = Replace(Replace(Replace(listCsv, "·", ","), "|", ","), ";", ",")
    parts = Split(listCsv, ",")
    For i = LBound(parts) To UBound(parts)
        If Len(Squash(parts(i))) > 0 Then
            If Squash(parts(i)) = target Then
                InList = True
                Exit Function
            End If
        End If
    Next i
End Function

' 견줄 때만 쓰는 납작한 형태 — 공백을 모두 지우고 대소문자를 맞춘다.
Public Function Squash(ByVal s As String) As String
    s = Replace(Replace(Replace(s, " ", ""), vbTab, ""), Chr$(160), "")
    Squash = LCase$(Trim$(s))
End Function

' 빌드가 모든 모듈을 실제로 컴파일하도록 부르는 자리다. VBA 는 필요할 때만
' 컴파일하기 때문에, 한 번도 불리지 않은 모듈의 오류는 사람이 그 단추를 누르는
' 순간에야 드러난다. 이 함수가 그 구멍을 막는다.
Public Function Ping() As String
    Ping = "modConfig"
End Function
