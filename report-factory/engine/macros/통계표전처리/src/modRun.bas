Attribute VB_Name = "modRun"
'==========================================================================
' modRun — 한 번에 실행, 그리고 실행 전 점검
'
' 실무에서는 1 → 2 → 3 → 4 를 늘 같은 순서로 누른다. 순서를 바꿔 누르면
' 미사용 배너에도 막대가 남거나, 지운 뒤에 다시 칠하는 헛일이 생긴다. 그래서
' 한 번에 도는 단추를 따로 두었다. 이 순서 안에서는 - 채우기가 늘 맨 마지막에
' 오므로(데이터막대·인사이트를 끝낸 뒤) 함께 묶어도 안전하다.
'
' 지우는 일만은 여기서 하지 않는다. 되돌릴 수 없는 일은 사람이 결과를 보고
' 스스로 눌러야 한다.
'==========================================================================
Option Explicit

Public Sub RunAll()
    Dim ws As Worksheet
    If Not modUtil.RequireSheet(ws) Then Exit Sub

    Dim tabs() As TTable, n As Long
    tabs = modTable.ScanTables(ws, n)
    If n = 0 Then
        modConfig.Say "이 시트에서 통계표를 찾지 못했습니다." & vbCrLf & vbCrLf & _
                      "[가져오기] 로 파일을 골라 열면 표가 있는 시트가 바로 열립니다.", vbExclamation
        Exit Sub
    End If

    Dim keep As String
    keep = modConfig.KeepGroups()
    If Not modConfig.Ask( _
        "이 시트에서 1 → 2 → 3 → 4 를 잇달아 실행합니다." & vbCrLf & vbCrLf & _
        "  1. 배너 선정 · 음영" & vbCrLf & _
        "  2. 데이터막대 (전체 + 배너)" & vbCrLf & _
        "  3. 인사이트 (인사이트 표시)" & vbCrLf & _
        "  4. 빈 칸을 - 로 채우기" & vbCrLf & vbCrLf & _
        "표 " & n & "장 · 보존 배너군 : " & IIf(Len(Trim$(keep)) = 0, "(지정 안 함)", keep) & vbCrLf & _
        "인사이트 기준 : " & modInsight.RuleText() & vbCrLf & vbCrLf & _
        "지우는 일(제외 배너 삭제)은 하지 않습니다. 결과를 보고 직접" & vbCrLf & _
        "눌러 주세요." & vbCrLf & vbCrLf & "계속하시겠습니까?", "한 번에 실행") Then Exit Sub

    Dim quiet As Boolean
    quiet = modConfig.IsSilent()
    modConfig.SetSilent True          ' 중간 안내는 접어 두고 끝에 한 번만 알린다
    On Error GoTo Done

    modBanner.SelectBanners
    Dim m1 As String: m1 = modConfig.LastMessage()
    modBars.BarsBoth
    Dim m2 As String: m2 = modConfig.LastMessage()
    modInsight.PaintInsight
    Dim m3 As String: m3 = modConfig.LastMessage()
    modDash.FillDashes
    Dim m4 As String: m4 = modConfig.LastMessage()

Done:
    modConfig.SetSilent quiet
    modConfig.Say "네 단계를 마쳤습니다." & vbCrLf & _
                  String$(46, "-") & vbCrLf & _
                  "1. " & FirstLine(m1) & vbCrLf & _
                  "2. " & FirstLine(m2) & vbCrLf & _
                  "3. " & FirstLine(m3) & vbCrLf & _
                  "4. " & FirstLine(m4) & vbCrLf & _
                  String$(46, "-") & vbCrLf & vbCrLf & _
                  "다음으로 할 일" & vbCrLf & _
                  "  · 회색 배너를 확인한 뒤 [제외 배너 삭제]" & vbCrLf & _
                  "  · 끝나면 [내보내기]", vbInformation
End Sub

'==========================================================================
Private Function FirstLine(ByVal s As String) As String
    Dim p As Long
    If Len(s) = 0 Then
        FirstLine = "(실행하지 못했습니다)"
        Exit Function
    End If
    p = InStr(s, vbCr)
    If p = 0 Then p = InStr(s, vbLf)
    If p = 0 Then FirstLine = s Else FirstLine = Left$(s, p - 1)
End Function

' 빌드가 모든 모듈을 실제로 컴파일하도록 부르는 자리다. VBA 는 필요할 때만
' 컴파일하기 때문에, 한 번도 불리지 않은 모듈의 오류는 사람이 그 단추를 누르는
' 순간에야 드러난다. 이 함수가 그 구멍을 막는다.
Public Function Ping() As String
    Ping = "modRun"
End Function
