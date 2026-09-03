Attribute VB_Name = "modTable"
'==========================================================================
' modTable — DP 원본 통계표(Table 시트)의 물리 구조를 읽는다.
'
' 통계표 한 장의 레이아웃은 다음과 같이 고정되어 있다.
'
'   행 +0   제목행       A열: 'A7. 중요 관리 업무_1순위'
'   행 +1   배너 그룹행  B열 또는 C열에 '전체', 그 오른쪽에 배너군 이름
'   행 +2   배너 세부행  배너군에 속한 항목 이름
'   행 +3   Base 행      A열 'Base for %', 각 열에 '(500)' 형태의 문자열
'   행 +4~  데이터 행
'   행 +n   계 행        A열 '      계'
'   (빈 행 1줄로 다음 표와 구분)
'
' 표1(응답자 특성)만 라벨 열이 A·B 두 칸이고 전체가 C열이지만, 이 모듈은
' 열 번호를 미리 정해 두지 않고 그룹행에서 '전체' 라벨을 찾아 기준 열을
' 잡기 때문에 표1도 예외 분기 없이 그대로 처리된다.
'==========================================================================
Option Explicit

Public Type TTable
    TitleRow As Long          ' 제목행
    GroupRow As Long          ' 배너 그룹행
    HeadRow As Long           ' 배너 세부행 (없으면 0)
    BaseRow As Long           ' Base for % 행
    LabelCol As Long          ' 라벨이 시작하는 열 — 표마다 A열이 아닐 수 있다
    EndRow As Long            ' 표의 마지막 행 (요약행 포함)
    TotalCol As Long          ' '전체' 열
    FirstCol As Long          ' 배너 첫 열 = TotalCol + 1
    LastCol As Long           ' 표의 마지막 열
    Title As String
    GroupCount As Long
    GroupName() As String     ' 1..GroupCount
    GroupFirst() As Long
    GroupLast() As Long
    DataRow() As Long         ' 1..DataCount — 보기 행 (계·요약 행은 뺀다)
    DataCount As Long
    SumRow() As Long          ' 1..SumCount — 【TOP2】·[5점 평균] 같은 요약 행
    SumCount As Long
    TopRow() As Long          ' 1..TopCount — 그 가운데 TOPn 행만 (TOP2·TOP3…)
    TopCount As Long
End Type

Private Const BASE_LABEL As String = "Base for %"

'==========================================================================
' 값 읽기 — 시트를 통째로 한 번에 떠 온다.
'
' 엑셀 셀을 VBA 에서 한 칸씩 읽으면 그때마다 프로그램 경계를 넘나든다.
' 표 276장짜리 통계표라면 그 왕복이 수십만 번이 되어, 하는 일에 견주어
' 터무니없이 느려진다. 그래서 시트의 값을 배열 하나로 한 번에 떠 놓고
' 그 뒤로는 메모리에서만 읽는다. 서식을 바꾸는 일만 엑셀에 되돌려 준다.
'
' 값을 고친 뒤에는 반드시 다시 떠야 한다. ScanTables 가 늘 먼저 뜨므로
' 바깥에서 따로 챙길 일은 없다.
'==========================================================================
Private gVals As Variant
Private gRows As Long
Private gCols As Long

Private Sub LoadSheet(ws As Worksheet)
    Dim ur As Range, lastR As Long, lastC As Long
    Dim arr1() As Variant
    gRows = 0: gCols = 0
    gVals = Empty
    Set ur = ws.UsedRange
    lastR = ur.Row + ur.Rows.Count - 1
    lastC = ur.Column + ur.Columns.Count - 1
    If lastR < 1 Or lastC < 1 Then Exit Sub

    ' 좌표를 그대로 쓰려고 늘 A1 부터 뜬다.
    If lastR = 1 And lastC = 1 Then
        ' 칸이 하나뿐이면 엑셀이 배열이 아니라 값을 그대로 준다.
        ReDim arr1(1 To 1, 1 To 1)
        arr1(1, 1) = ws.Cells(1, 1).Value2
        gVals = arr1
    Else
        gVals = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).Value2
    End If
    gRows = lastR
    gCols = lastC
End Sub

' 떠 온 배열에서 글자로 읽는다. 범위를 벗어나면 빈 글자다.
Public Function Txt(ByVal r As Long, ByVal c As Long) As String
    If r < 1 Or c < 1 Or r > gRows Or c > gCols Then Exit Function
    Dim v As Variant
    v = gVals(r, c)
    If IsEmpty(v) Then Exit Function
    If IsError(v) Then Exit Function
    Txt = Trim$(CStr(v))
End Function

' 떠 온 배열에서 숫자로 읽는다. 숫자가 아니면 False.
Public Function Num(ByVal r As Long, ByVal c As Long, ByRef outVal As Double) As Boolean
    If r < 1 Or c < 1 Or r > gRows Or c > gCols Then Exit Function
    Dim v As Variant
    v = gVals(r, c)
    If IsEmpty(v) Then Exit Function
    If IsError(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    If Len(Trim$(CStr(v))) = 0 Then Exit Function
    outVal = CDbl(v)
    Num = True
End Function

' 그 칸이 비어 있는가 (떠 온 배열 기준).
Public Function IsBlank(ByVal r As Long, ByVal c As Long) As Boolean
    IsBlank = (Len(Txt(r, c)) = 0)
End Function

Public Function LoadedRows() As Long
    LoadedRows = gRows
End Function

Public Function LoadedCols() As Long
    LoadedCols = gCols
End Function

'--------------------------------------------------------------------------
' 시트에서 표를 모두 찾아 배열로 돌려준다. 표가 없으면 Count 가 0 이다.
'--------------------------------------------------------------------------
Public Function ScanTables(ws As Worksheet, ByRef outCount As Long) As TTable()
    Dim res() As TTable
    Dim n As Long, r As Long, lastRow As Long
    ReDim res(1 To 400)
    n = 0
    ' 시트를 통째로 한 번에 떠 온다. 이 뒤로는 셀을 하나도 직접 읽지 않는다.
    LoadSheet ws
    lastRow = gRows
    If lastRow < 1 Then
        outCount = 0
        ScanTables = res
        Exit Function
    End If

    ' 'Base for %' 가 늘 A열에 있지는 않다. SKEE 통계표처럼 표1만 A열이고
    ' 나머지 표는 B열에서 시작하는 경우가 있어 앞쪽 몇 열을 함께 본다.
    Dim c0 As Long
    For r = 1 To lastRow
        For c0 = 1 To 4
            If Txt(r, c0) = BASE_LABEL Then
                Dim t As TTable
                If BuildTable(ws, r, c0, t) Then
                    n = n + 1
                    If n > UBound(res) Then ReDim Preserve res(1 To UBound(res) * 2)
                    res(n) = t
                End If
                Exit For
            End If
        Next c0
    Next r

    ' 표와 표 사이에 빈 행이 없는 통계표가 오면 앞 표의 끝이 다음 표를 삼킨다.
    ' 다음 표의 제목행 앞에서 반드시 끊어 그런 일을 막는다.
    Dim i As Long
    For i = 1 To n - 1
        If res(i).EndRow >= res(i + 1).TitleRow Then
            res(i).EndRow = res(i + 1).TitleRow - 1
            TrimDataRows res(i)
        End If
    Next i

    outCount = n
    If n > 0 Then ReDim Preserve res(1 To n)
    ScanTables = res
End Function

' 표의 끝이 앞당겨졌으면 행 목록도 그 안으로 줄인다.
Private Sub TrimDataRows(ByRef t As TTable)
    Dim k As Long, keep As Long
    For k = 1 To t.DataCount
        If t.DataRow(k) <= t.EndRow Then keep = keep + 1
    Next k
    t.DataCount = keep
    If keep = 0 Then Erase t.DataRow

    keep = 0
    For k = 1 To t.SumCount
        If t.SumRow(k) <= t.EndRow Then keep = keep + 1
    Next k
    t.SumCount = keep
    If keep = 0 Then Erase t.SumRow

    keep = 0
    For k = 1 To t.TopCount
        If t.TopRow(k) <= t.EndRow Then keep = keep + 1
    Next k
    t.TopCount = keep
    If keep = 0 Then Erase t.TopRow
End Sub

'==========================================================================
' TOPn 이 있는 표는 다르게 다룬다.
'
' 5점 척도에는 【TOP2】, 7점 척도에는 【TOP3】 처럼 상위 응답을 묶은 줄이 붙는다.
' 보고서는 그 표에서 보기별 분포를 쓰지 않고 TOPn 한 줄만 쓴다. 그래서
'
'   · 데이터막대는 아예 넣지 않는다 (읽지 않을 분포에 눈이 가게 만들 뿐이다)
'   · 인사이트는 TOPn 줄에만 칠한다
'
' 보기 행을 지우지는 않는다. 근거는 표에 그대로 남겨 둔다.
'--------------------------------------------------------------------------
Public Function HasTopN(t As TTable) As Boolean
    HasTopN = (t.TopCount > 0)
End Function

' 데이터막대를 걸어도 되는 표인가.
Public Function BarsAllowed(t As TTable) As Boolean
    BarsAllowed = (t.TopCount = 0)
End Function

'--------------------------------------------------------------------------
' 인사이트가 볼 행. TOPn 이 있으면 그 줄만, 없으면 보기 행과 요약 행 전부다.
' k 는 1 부터 ScanRowCount 까지다.
'--------------------------------------------------------------------------
Public Function ScanRowCount(t As TTable) As Long
    If t.TopCount > 0 Then
        ScanRowCount = t.TopCount
    Else
        ScanRowCount = t.DataCount + t.SumCount
    End If
End Function

Public Function ScanRowAt(t As TTable, ByVal k As Long) As Long
    If t.TopCount > 0 Then
        If k <= t.TopCount Then ScanRowAt = t.TopRow(k)
        Exit Function
    End If
    If k <= t.DataCount Then
        ScanRowAt = t.DataRow(k)
    ElseIf k - t.DataCount <= t.SumCount Then
        ScanRowAt = t.SumRow(k - t.DataCount)
    End If
End Function

' 【TOP2】·[TOP3] 처럼 상위 n 개를 묶은 줄인가.
Public Function IsTopNLabel(ByVal lab As String) As Boolean
    Dim s As String, i As Long, ch As String
    s = lab
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = "[" Or ch = "]" Or ch = "【" Or ch = "】" Or ch = " " Then
            s = Left$(s, i - 1) & Mid$(s, i + 1)
            i = i - 1
        End If
        If i >= Len(s) Then Exit For
    Next i
    s = UCase$(Trim$(s))
    If Len(s) < 4 Then Exit Function
    If Left$(s, 3) <> "TOP" Then Exit Function
    IsTopNLabel = IsNumeric(Mid$(s, 4, 1))
End Function

'--------------------------------------------------------------------------
' Base 행 하나를 기준으로 표 한 장의 구조를 채운다.
'--------------------------------------------------------------------------
Private Function BuildTable(ws As Worksheet, baseRow As Long, baseCol As Long, ByRef t As TTable) As Boolean
    Dim gRow As Long, c As Long, probe As Long

    ' 1) 그룹행 — Base 행 위쪽에서 '전체' 라벨이 있는 행을 찾는다.
    gRow = 0
    For probe = baseRow - 1 To baseRow - 3 Step -1
        If probe < 1 Then Exit For
        For c = 1 To 60
            If Txt(probe, c) = "전체" Then
                gRow = probe
                t.TotalCol = c
                Exit For
            End If
        Next c
        If gRow > 0 Then Exit For
    Next probe
    If gRow = 0 Then Exit Function          ' 구조가 다르면 건드리지 않는다

    t.BaseRow = baseRow
    t.LabelCol = baseCol
    t.GroupRow = gRow
    If gRow <= baseRow - 2 Then t.HeadRow = gRow + 1 Else t.HeadRow = 0
    t.TitleRow = gRow - 1
    t.FirstCol = t.TotalCol + 1
    ' 제목도 라벨 열에서 읽는다 (A열이 아닐 수 있다).
    Dim tc As Long
    For tc = baseCol To t.TotalCol - 1
        If Len(Txt(t.TitleRow, tc)) > 0 Then
            t.Title = Txt(t.TitleRow, tc)
            Exit For
        End If
    Next tc

    ' 2) 마지막 열 — 그룹행·세부행·Base 행 가운데 가장 오른쪽
    t.LastCol = LastUsedCol(ws, gRow)
    If t.HeadRow > 0 Then t.LastCol = MaxL(t.LastCol, LastUsedCol(ws, t.HeadRow))
    t.LastCol = MaxL(t.LastCol, LastUsedCol(ws, baseRow))

    ' 3) 표의 끝 — 모든 열이 빈 행을 만나면 거기서 끊는다.
    Dim r As Long, empty_ As Boolean
    r = baseRow + 1
    Do
        empty_ = True
        For c = 1 To t.LastCol
            If Len(Txt(r, c)) > 0 Then
                empty_ = False
                Exit For
            End If
        Next c
        If empty_ Then Exit Do
        ' 다음 표의 Base 행을 만나도 끊는다 (빈 행 없이 이어 붙은 통계표 대비).
        Dim probeC As Long
        For probeC = 1 To 4
            If Txt(r, probeC) = BASE_LABEL Then Exit Do
        Next probeC
        r = r + 1
        If r > baseRow + 2000 Then Exit Do
    Loop
    t.EndRow = r - 1
    If t.EndRow <= baseRow Then Exit Function

    ' 4) 배너군 경계 — 그룹행에 이름이 적힌 열이 각 군의 시작이다.
    Dim gn() As String, gf() As Long, gl() As Long, gc As Long
    ReDim gn(1 To 40): ReDim gf(1 To 40): ReDim gl(1 To 40)
    gc = 0
    For c = t.FirstCol To t.LastCol
        If Len(Txt(gRow, c)) > 0 Then
            If gc > 0 Then gl(gc) = c - 1
            gc = gc + 1
            If gc > UBound(gn) Then
                ReDim Preserve gn(1 To gc + 20)
                ReDim Preserve gf(1 To gc + 20)
                ReDim Preserve gl(1 To gc + 20)
            End If
            gn(gc) = Txt(gRow, c)
            gf(gc) = c
        End If
    Next c
    If gc > 0 Then gl(gc) = t.LastCol
    t.GroupCount = gc
    If gc > 0 Then
        ReDim Preserve gn(1 To gc): ReDim Preserve gf(1 To gc): ReDim Preserve gl(1 To gc)
        t.GroupName = gn: t.GroupFirst = gf: t.GroupLast = gl
    End If

    ' 5) 데이터 행 — Base 바로 아래부터 '계' 행 직전까지, 라벨이 있는 행만 담는다.
    '    라벨 열은 A열부터 '전체' 열 왼쪽까지다(표1은 A·B 두 칸).
    '    보기 행과 요약 행을 갈라 담는다. 데이터막대는 보기 행에만 걸어야
    '    막대 길이의 기준이 흔들리지 않고(평균 행은 값의 폭이 다르다),
    '    인사이트은 【TOP2】 같은 요약 행까지 함께 보아야 한다.
    '    연구원이 표에서 가장 많이 읽는 줄이 바로 그 요약 행이기 때문이다.
    Dim dr() As Long, dc As Long, lab As String, k As Long, hasLabel As Boolean
    Dim sr() As Long, sc As Long, inBody As Boolean
    Dim tp() As Long, tpc As Long
    ReDim dr(1 To (t.EndRow - baseRow) + 1)
    ReDim sr(1 To (t.EndRow - baseRow) + 1)
    ReDim tp(1 To (t.EndRow - baseRow) + 1)
    dc = 0: sc = 0: tpc = 0: inBody = True
    For r = baseRow + 1 To t.EndRow
        lab = ""
        hasLabel = False
        For k = t.LabelCol To t.TotalCol - 1
            If Len(Txt(r, k)) > 0 Then
                lab = Txt(r, k)
                hasLabel = True
            End If
        Next k
        If hasLabel Then
            If lab = "계" Then
                inBody = False                               ' 계 행은 늘 100 이라 아무 데도 담지 않는다
            ElseIf Left$(lab, 1) = "[" Or Left$(lab, 1) = "【" Then
                inBody = False
                sc = sc + 1
                sr(sc) = r
                If IsTopNLabel(lab) Then
                    tpc = tpc + 1
                    tp(tpc) = r
                End If
            ElseIf inBody Then
                dc = dc + 1
                dr(dc) = r
            End If
        End If
    Next r
    t.DataCount = dc
    If dc > 0 Then
        ReDim Preserve dr(1 To dc)
        t.DataRow = dr
    End If
    t.SumCount = sc
    If sc > 0 Then
        ReDim Preserve sr(1 To sc)
        t.SumRow = sr
    End If
    t.TopCount = tpc
    If tpc > 0 Then
        ReDim Preserve tp(1 To tpc)
        t.TopRow = tp
    End If

    BuildTable = (t.GroupCount > 0 Or t.DataCount > 0)
End Function

'--------------------------------------------------------------------------
' '(500)' 형태의 Base 문자열을 숫자로 바꾼다. 읽을 수 없으면 0.
'--------------------------------------------------------------------------
Public Function BaseOf(ws As Worksheet, t As TTable, col As Long) As Long
    Dim s As String
    s = Txt(t.BaseRow, col)
    s = Replace(Replace(Replace(Replace(s, "(", ""), ")", ""), ",", ""), " ", "")
    If Len(s) = 0 Then Exit Function
    If Not IsNumeric(s) Then Exit Function
    BaseOf = CLng(Val(s))
End Function

Public Function LastUsedCol(ws As Worksheet, r As Long) As Long
    Dim c As Long
    For c = gCols To 1 Step -1
        If Len(Txt(r, c)) > 0 Then
            LastUsedCol = c
            Exit Function
        End If
    Next c
End Function

Public Function MaxL(a As Long, b As Long) As Long
    If a > b Then MaxL = a Else MaxL = b
End Function

' 빌드가 모든 모듈을 실제로 컴파일하도록 부르는 자리다. VBA 는 필요할 때만
' 컴파일하기 때문에, 한 번도 불리지 않은 모듈의 오류는 사람이 그 단추를 누르는
' 순간에야 드러난다. 이 함수가 그 구멍을 막는다.
Public Function Ping() As String
    Ping = "modTable"
End Function
