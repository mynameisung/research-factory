Attribute VB_Name = "modUtil"
'==========================================================================
' modUtil — 여러 기능이 함께 쓰는 잔손질과 안전장치
'
' 실무에서 매크로가 사고를 내는 자리는 대체로 정해져 있다. 시트가 보호되어
' 있거나, 셀이 병합되어 있거나, 표가 아닌 시트에서 눌렀거나, 오래 걸리는데
' 아무 표시가 없어 사람이 엑셀을 강제로 닫는 경우다. 이 모듈이 그 자리를 막는다.
'==========================================================================
Option Explicit

Private savedCalc As Long
Private savedEvents As Boolean
Private savedScreen As Boolean
Private busy As Boolean

'--------------------------------------------------------------------------
' 오래 걸리는 작업을 감싼다. 화면 갱신과 자동 계산을 멈춰 속도를 올리고,
' 중간에 오류가 나도 엑셀을 원래 상태로 되돌린다.
'--------------------------------------------------------------------------
Public Sub BeginWork(ByVal caption As String)
    If busy Then Exit Sub
    busy = True
    savedCalc = Application.Calculation
    savedEvents = Application.EnableEvents
    savedScreen = Application.ScreenUpdating
    On Error Resume Next
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.StatusBar = caption
    On Error GoTo 0
End Sub

Public Sub EndWork()
    On Error Resume Next
    Application.StatusBar = False
    Application.ScreenUpdating = savedScreen
    Application.EnableEvents = savedEvents
    Application.Calculation = savedCalc
    On Error GoTo 0
    busy = False
End Sub

Public Sub Progress(ByVal caption As String, ByVal done As Long, ByVal total As Long)
    If total <= 0 Then Exit Sub
    ' 표를 한 장 처리할 때마다 고쳐 쓰면 그것 자체가 느려진다. 스무 장에 한 번만 알린다.
    If done Mod 20 <> 0 And done <> total Then Exit Sub
    On Error Resume Next
    Application.StatusBar = caption & "  " & done & " / " & total & _
                            "  (" & Format$(done / total, "0%") & ")"
    On Error GoTo 0
End Sub

'--------------------------------------------------------------------------
' 서식을 바꿀 수 있는 시트인지 미리 본다. 보호된 시트에서 매크로를 돌리면
' 중간에 멎어 무엇이 처리되고 무엇이 안 됐는지 알 수 없게 된다.
'--------------------------------------------------------------------------
Public Function SheetWritable(ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function
    If ws.ProtectContents Then
        modConfig.Say "이 시트는 보호되어 있어 서식을 바꿀 수 없습니다." & vbCrLf & vbCrLf & _
                      "엑셀 리본의 [검토] 탭 → [시트 보호 해제] 를 누른 뒤" & vbCrLf & _
                      "다시 실행해 주세요.", vbExclamation
        Exit Function
    End If
    If ActiveWorkbook Is Nothing Then Exit Function
    SheetWritable = True
End Function

'--------------------------------------------------------------------------
' 일할 시트를 고른다.
'
' [배너군 목록] 을 만들면 그 시트가 앞으로 나온다. 그 상태에서
' 다음 단추를 누르면 「통계표를 찾지 못했습니다」가 뜨는데, 사람 눈에는 방금까지
' 잘 되던 것이 갑자기 안 되는 것으로 보인다. 그래서 지금 시트에 표가 없으면
' 통합 문서에서 표가 있는 시트를 찾아 그리로 옮겨 간다.
'--------------------------------------------------------------------------
Public Function TargetSheet() As Worksheet
    Dim ws As Worksheet, n As Long, dummy() As TTable

    Set ws = ActiveSheet
    If Not ws Is Nothing Then
        dummy = modTable.ScanTables(ws, n)
        If n > 0 Then
            Set TargetSheet = ws
            Exit Function
        End If
    End If

    If ActiveWorkbook Is Nothing Then Exit Function
    For Each ws In ActiveWorkbook.Worksheets
        If Not ws.Visible = xlSheetVeryHidden Then
            dummy = modTable.ScanTables(ws, n)
            If n > 0 Then
                ws.Activate
                Set TargetSheet = ws
                Exit Function
            End If
        End If
    Next ws
End Function

'--------------------------------------------------------------------------
' 모든 기능의 첫 줄. 표가 있는 시트를 찾아 주고, 없으면 안내하고 False 를 돌려준다.
' 보호된 시트도 여기서 걸러 낸다. 기능마다 같은 검사를 따로 적다가 하나가 빠져
' 「대상이 없는 채로 진행」하는 사고가 났다 — 이 함수 하나만 부르면 된다.
'--------------------------------------------------------------------------
Public Function RequireSheet(ByRef ws As Worksheet) As Boolean
    Set ws = TargetSheet()
    If ws Is Nothing Then
        modConfig.Say "이 통합 문서에서 통계표를 찾지 못했습니다." & vbCrLf & vbCrLf & _
                      "A열 근처에 'Base for %' 가 있고 그 위에 '전체' 라벨이 있는" & vbCrLf & _
                      "원본 통계표의 Table 시트에서 실행해 주세요." & vbCrLf & _
                      "[가져오기] 로 파일을 고르면 표가 있는 시트가 바로 열립니다.", vbExclamation
        Exit Function
    End If
    RequireSheet = SheetWritable(ws)
End Function

'--------------------------------------------------------------------------
' 병합된 셀이 걸쳐 있으면 열을 지울 때 엑셀이 멎는다. 지우기 전에 푼다.
'--------------------------------------------------------------------------
Public Function UnmergeIfNeeded(rng As Range) As Long
    Dim n As Long
    On Error Resume Next
    If IsNull(rng.MergeCells) Then
        n = 1
        rng.UnMerge
    ElseIf rng.MergeCells = True Then
        n = 1
        rng.UnMerge
    End If
    On Error GoTo 0
    UnmergeIfNeeded = n
End Function

'--------------------------------------------------------------------------
' 표시값 — 통계표 셀에 저장된 값은 반올림되지 않은 부동소수다.
' 표시는 8.4 인데 저장값은 8.391608... 인 식이다. 보고서에 쓰는 수치는 늘
' 서식이 적용된 값이므로, 인사이트 판정도 표시값으로 해야 사람이 보는 것과 맞다.
'--------------------------------------------------------------------------
Public Function DisplayVal(cel As Range, ByVal raw As Double) As Double
    Dim digits As Long
    digits = DigitsOf(cel)
    DisplayVal = Round(raw, digits)
End Function

Private Function DigitsOf(cel As Range) As Long
    Dim f As String, p As Long
    On Error Resume Next
    f = CStr(cel.NumberFormat)
    On Error GoTo 0
    p = InStr(f, ".")
    If p = 0 Then
        DigitsOf = 1                      ' 통계표 % 의 기본은 소수 1자리다
        Exit Function
    End If
    Dim i As Long, n As Long
    For i = p + 1 To Len(f)
        If Mid$(f, i, 1) = "0" Or Mid$(f, i, 1) = "#" Then n = n + 1 Else Exit For
    Next i
    If n = 0 Then n = 1
    DigitsOf = n
End Function

'--------------------------------------------------------------------------
' 결과를 담을 시트를 만든다. 같은 이름이 있으면 지우고 새로 만든다.
'--------------------------------------------------------------------------
Public Function FreshSheet(wb As Workbook, ByVal name As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Set ws = wb.Worksheets(name)
    If Not ws Is Nothing Then ws.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.name = name
    Set FreshSheet = ws
End Function

Public Sub HeaderRow(ws As Worksheet, ByVal r As Long, titles As Variant)
    Dim i As Long
    For i = LBound(titles) To UBound(titles)
        ws.Cells(r, i - LBound(titles) + 1).Value = titles(i)
    Next i
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, UBound(titles) - LBound(titles) + 1))
        .Font.Bold = True
        .Interior.Color = RGB(238, 238, 238)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
End Sub

' 빌드가 모든 모듈을 실제로 컴파일하도록 부르는 자리다. VBA 는 필요할 때만
' 컴파일하기 때문에, 한 번도 불리지 않은 모듈의 오류는 사람이 그 단추를 누르는
' 순간에야 드러난다. 이 함수가 그 구멍을 막는다.
Public Function Ping() As String
    Ping = "modUtil"
End Function
