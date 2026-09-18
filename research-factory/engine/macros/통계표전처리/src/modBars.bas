Attribute VB_Name = "modBars"
'==========================================================================
' modBars — 기능 2. 데이터막대
'
' 연구원이 표 한 장의 분포를 한눈에 읽고, 전체의 분포와 배너들의 분포가 어떻게
' 다른지 곧바로 알아보게 하려는 것이다. 그래서 막대의 길이 기준을 열마다 따로
' 잡지 않고 0~100 으로 고정한다. 열마다 최댓값에 맞춰 늘이면 막대는 예뻐지지만
' 열끼리 견줄 수 없게 되어 목적을 잃는다.
'
' 다만 5점 척도 평균이나 100점 환산처럼 값의 폭이 다른 표가 섞여 있으므로,
' 그 열의 최댓값이 10 이하이면 그 열의 최솟값·최댓값에 맞춘다.
'
' 칠하는 범위는 Base 행 바로 아래부터 계 행 직전까지이고, 라벨이 없는 행과
' 【TOP2】·[5점 평균] 같은 요약 행은 뺀다.
'==========================================================================
Option Explicit

Public Const BAR_TOTAL As Long = 12611584      ' RGB(0, 112, 192) 파랑 - 전체 열
Public Const BAR_BANNER As Long = 10921638     ' RGB(166, 166, 166) 회색 - 배너 열

'--------------------------------------------------------------------------
Public Sub BarsTotal()
    ApplyBars True, False
End Sub

Public Sub BarsBanner()
    ApplyBars False, True
End Sub

Public Sub BarsBoth()
    ApplyBars True, True
End Sub

'--------------------------------------------------------------------------
Private Sub ApplyBars(doTotal As Boolean, doBanner As Boolean)
    Dim ws As Worksheet
    If Not modUtil.RequireSheet(ws) Then Exit Sub

    Dim tabs() As TTable, n As Long
    tabs = modTable.ScanTables(ws, n)
    If n = 0 Then
        modConfig.Say "이 시트에서 통계표를 찾지 못했습니다.", vbExclamation
        Exit Sub
    End If

    modUtil.BeginWork "데이터막대를 넣는 중입니다"
    Dim i As Long, c As Long, cols As Long, skipped As Long
    For i = 1 To n
        ' TOPn 이 붙은 표는 보고서가 보기별 분포를 쓰지 않고 TOPn 한 줄만 쓴다.
        ' 읽지 않을 분포에 막대를 그려 두면 눈만 그리로 끌린다.
        If Not modTable.BarsAllowed(tabs(i)) Then
            skipped = skipped + 1
        ElseIf tabs(i).DataCount > 0 Then
            If doTotal Then
                If BarOneColumn(ws, tabs(i), tabs(i).TotalCol, BAR_TOTAL) Then cols = cols + 1
            End If
            If doBanner Then
                For c = tabs(i).FirstCol To tabs(i).LastCol
                    ' 미사용으로 골라 둔(회색) 배너 열에는 막대를 넣지 않는다.
                    If ws.Cells(tabs(i).GroupRow, c).Interior.Color <> modBanner.SHADE_COLOR Then
                        If BarOneColumn(ws, tabs(i), c, BAR_BANNER) Then cols = cols + 1
                    End If
                Next c
            End If
        End If
        modUtil.Progress "데이터막대를 넣는 중입니다", i, n
    Next i
    modUtil.EndWork

    Dim msg As String
    msg = "표 " & n & "장에 데이터막대를 넣었습니다." & vbCrLf & vbCrLf & _
          "처리한 열 : " & cols & "개" & vbCrLf & _
          "길이 기준 : 0~100 고정 (값이 작은 표는 그 열의 범위에 맞춤)"
    If skipped > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "TOPn 이 있는 표 " & skipped & "장은 건너뛰었습니다." & vbCrLf & _
              "그 표는 보기별 분포가 아니라 TOPn 한 줄만 쓰기 때문입니다."
    End If
    modConfig.Say msg, vbInformation
End Sub

'--------------------------------------------------------------------------
' 열 하나에 데이터막대를 건다. 데이터 행이 떨어져 있어도 한 규칙으로 묶는다.
'--------------------------------------------------------------------------
' 값은 modTable 이 떠 놓은 배열에서 읽고, 범위는 주소 문자열로 한 번에 만든다.
' Union 을 행마다 부르면 표 이백 장짜리에서 그것만으로 몇 초가 사라진다.
Private Function BarOneColumn(ws As Worksheet, t As TTable, col As Long, barColor As Long) As Boolean
    Dim rng As Range, k As Long, v As Double, maxV As Double, hasNum As Boolean
    Dim addr As String, colL As String, runStart As Long, prevRow As Long

    colL = Split(ws.Cells(1, col).Address(True, False), "$")(0)
    For k = 1 To t.DataCount
        If modTable.Num(t.DataRow(k), col, v) Then
            hasNum = True
            If v > maxV Then maxV = v
        End If
        ' 이어진 행끼리 묶어 A5:A12 꼴로 적는다.
        If runStart = 0 Then
            runStart = t.DataRow(k)
        ElseIf t.DataRow(k) <> prevRow + 1 Then
            addr = addr & IIf(Len(addr) > 0, ",", "") & colL & runStart & ":" & colL & prevRow
            runStart = t.DataRow(k)
        End If
        prevRow = t.DataRow(k)
    Next k
    If runStart > 0 Then addr = addr & IIf(Len(addr) > 0, ",", "") & colL & runStart & ":" & colL & prevRow
    If Len(addr) = 0 Or Not hasNum Then Exit Function

    On Error Resume Next
    Set rng = ws.Range(addr)
    On Error GoTo 0
    If rng Is Nothing Then Exit Function

    ' 이 열에 이미 걸려 있던 데이터막대는 지우고 새로 건다.
    ClearBarsIn rng

    On Error GoTo Fail
    Dim db As Databar
    Set db = rng.FormatConditions.AddDatabar
    db.BarFillType = xlDataBarFillGradient
    db.BarColor.Color = barColor
    db.ShowValue = True

    If maxV <= 10 Then
        ' 5점 척도 평균처럼 값이 작은 표 - 그 열의 범위에 맞춘다.
        db.MinPoint.Modify newtype:=xlConditionValueLowestValue
        db.MaxPoint.Modify newtype:=xlConditionValueHighestValue
    Else
        ' 백분율 표 - 열끼리 견줄 수 있도록 0~100 으로 고정한다.
        db.MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=0
        db.MaxPoint.Modify newtype:=xlConditionValueNumber, newvalue:=100
    End If
    BarOneColumn = True
    Exit Function
Fail:
    ' 조건부서식 개수 상한 같은 이유로 실패해도 나머지 열은 계속 처리한다.
    BarOneColumn = False
End Function

'--------------------------------------------------------------------------
Public Sub ClearBars()
    Dim ws As Worksheet
    If Not modUtil.RequireSheet(ws) Then Exit Sub
    modUtil.BeginWork "데이터막대를 지우는 중입니다"
    ClearBarsIn ws.UsedRange
    modUtil.EndWork
    modConfig.Say "이 시트의 데이터막대를 모두 지웠습니다.", vbInformation
End Sub

' 범위에 걸린 데이터막대만 지운다. 다른 조건부서식은 건드리지 않는다.
'
' 규칙을 하나씩 지우면 지울 때마다 컬렉션이 다시 짜여, 규칙이 수천 개인 시트에서는
' 개수의 제곱에 가깝게 느려진다(실측 160초). 그래서 먼저 훑어보고 규칙이 모두
' 데이터막대이면 통째로 한 번에 지운다. 이 매크로가 넣는 조건부서식은 데이터막대뿐이라
' 대개 이 길로 간다. 다른 서식이 섞여 있을 때만 하나씩 지운다.
Public Sub ClearBarsIn(rng As Range)
    Dim i As Long, cnt As Long, onlyBars As Boolean
    On Error Resume Next
    cnt = rng.FormatConditions.Count
    If cnt = 0 Then Exit Sub

    onlyBars = True
    For i = 1 To cnt
        If TypeName(rng.FormatConditions(i)) <> "Databar" Then
            onlyBars = False
            Exit For
        End If
    Next i

    If onlyBars Then
        rng.FormatConditions.Delete
    Else
        For i = cnt To 1 Step -1
            If TypeName(rng.FormatConditions(i)) = "Databar" Then rng.FormatConditions(i).Delete
        Next i
    End If
    On Error GoTo 0
End Sub

' 빌드가 모든 모듈을 실제로 컴파일하도록 부르는 자리다. VBA 는 필요할 때만
' 컴파일하기 때문에, 한 번도 불리지 않은 모듈의 오류는 사람이 그 단추를 누르는
' 순간에야 드러난다. 이 함수가 그 구멍을 막는다.
Public Function Ping() As String
    Ping = "modBars"
End Function
