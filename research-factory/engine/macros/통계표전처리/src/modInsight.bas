Attribute VB_Name = "modInsight"
'==========================================================================
' modInsight — 기능 3. 인사이트 칠하기
'
' 전체 값을 기준으로 삼아, 그 기준에서 크게 벗어난 배너 값에 형광 배경색을 칠한다.
' 값어치는 무난한 사실이 아니라 전체와 다르게 움직이는 집단에 있기 때문이다.
'
' 인사이트 기준은 리본에서 두 방식 가운데 골라 값을 직접 바꿀 수 있다.
'
'   표준편차   그 행에 놓인 배너 값들의 표준편차 시그마를 구해
'              |배너값 - 전체값| > k x 시그마 이면 칠한다. 행마다 흩어진 정도가
'              달라도 한 기준으로 표 전체를 훑을 수 있다.
'   절대값     |배너값 - 전체값| > 입력한 %p 이면 칠한다. 기준이 눈에 보여
'              결과를 남에게 설명하기 쉽다.
'
' 세 가지를 함께 지킨다.
'   1) Base 가 하한(기본 30) 에 못 미치는 열은 보지 않는다. 모수가 모자란 칸의
'      % 는 튀는 값이 아니라 표본 잡음이라, 칠해 두면 오해를 부른다.
'   2) 미사용으로 골라 둔(회색) 배너 열도 보지 않는다. 어차피 보고서에 쓰지 않는다.
'   3) 판정은 화면에 보이는 표시값으로 한다. 저장값은 반올림되지 않은 부동소수라
'      경계에서 사람이 보는 것과 어긋난다.
'
' 배경색은 조건부서식이 아니라 셀 서식이므로 데이터막대 위에 겹쳐서 칠해진다.
'==========================================================================
Option Explicit

' 방향을 색으로 가른다. 한 가지 색으로 칠하면 「전체와 다르다」까지만 알 수 있고,
' 높아서 다른지 낮아서 다른지는 다시 숫자를 읽어야 안다.
Public Const HL_HIGH As Long = 13551615        ' RGB(255, 199, 206) 분홍 - 전체보다 높음
Public Const HL_LOW  As Long = 15652797        ' RGB(189, 215, 238) 하늘 - 전체보다 낮음
Public Const HL_OLD  As Long = 65535           ' 예전판이 칠하던 노랑 (지울 때만 쓴다)

' 모듈 수준 선언은 반드시 여기(첫 프로시저 앞)에 둔다. 프로시저 사이에 끼우면
' VBA 가 모듈 수준으로 보지 않는다.
' 어디까지 갔는지 남겨 두는 자리. 오류가 났을 때만 쓴다.
Private gAtTable As Long
Private gAtTitle As String
Private gAtRow As Long
Private gAtCol As Long

'--------------------------------------------------------------------------
Public Sub PaintInsight()
    On Error GoTo Fail
    Dim ws As Worksheet
    If Not modUtil.RequireSheet(ws) Then Exit Sub

    Dim thr As Double
    thr = modConfig.InsightValue()
    If thr <= 0 Then
        modConfig.Say "인사이트 기준값이 0 이하입니다. 리본의 [기준값] 칸에 양수를 넣어 주세요.", vbExclamation
        Exit Sub
    End If

    Dim nTab As Long, nHit As Long
    modUtil.BeginWork "인사이트를 찾는 중입니다"
    ScanInsights ws, True, nTab, nHit
    modUtil.EndWork

    If nTab = 0 Then
        modConfig.Say "이 시트에서 통계표를 찾지 못했습니다.", vbExclamation
        Exit Sub
    End If

    ' 한 칸도 안 나왔다면 정말 벗어난 값이 없어서인지, 볼 수 있는 배너 열
    ' 자체가 없어서인지를 갈라 알려 준다. [보존 배너군] 은 레지스트리에 남아
    ' 다른 통계표를 열어도 그대로 따라오므로, 이름이 안 맞으면 배너가 통째로
    ' 빠져 아무 것도 칠리지 않는 일이 흔하다 - 그때는 원인이 TOPn 이 아니라
    ' 배너 선정이라고 짚어 주어야 한다.
    Dim hint As String
    If nHit = 0 Then
        Dim tabs() As TTable, tn As Long
        tabs = modTable.ScanTables(ws, tn)
        hint = ZeroHitHint(ws, tabs, tn, modConfig.MinBase())
    End If

    modConfig.Say "표 " & nTab & "장에서 " & nHit & "칸을 칠했습니다." & vbCrLf & vbCrLf & _
                  "분홍 = 전체보다 높음 · 하늘 = 전체보다 낮음" & vbCrLf & vbCrLf & _
                  "기준 : " & RuleText() & vbCrLf & _
                  "제외 : Base " & modConfig.MinBase() & " 미만 열, 미사용(회색) 배너 열" & vbCrLf & _
                  "       TOPn 이 있는 표는 TOPn 줄에만 칠합니다" & hint, vbInformation
    Exit Sub

Fail:
    ' 오류를 삼키지 않고 어디서 멎었는지 짚어 준다. 표가 이백 장 넘는 통계표에서
    ' 「어딘가에서 오류」라는 말만으로는 사람이 손쓸 자리가 없다.
    modUtil.EndWork
    modConfig.Say "인사이트를 찾다가 멈췄습니다." & vbCrLf & vbCrLf & _
                  "오류 " & Err.Number & " : " & Err.Description & vbCrLf & _
                  "자리 : " & Position(), vbCritical
End Sub


Public Function Position() As String
    Position = "표 " & gAtTable & " (" & gAtTitle & ")  행 " & gAtRow & " 열 " & gAtCol
End Function

'--------------------------------------------------------------------------
Public Sub ClearInsight()
    On Error GoTo Fail
    Dim ws As Worksheet
    If Not modUtil.RequireSheet(ws) Then Exit Sub

    ' 배너군 범위 단위로 한 번에 지운다. 칸마다 색을 물어 가며 지우면 표가
    ' 이백 장 넘는 통계표에서는 견딜 수 없이 느리다. 회색으로 칠해 둔 배너군은
    ' 애초에 인사이트를 칠하지 않으므로 건너뛴다 - 배너 선정 결과가 그대로 남는다.
    Dim tabs() As TTable, n As Long, i As Long, g As Long, cleared As Long
    modUtil.BeginWork "인사이트 색을 지우는 중입니다"
    tabs = modTable.ScanTables(ws, n)
    For i = 1 To n
        If tabs(i).EndRow > tabs(i).BaseRow Then
            For g = 1 To tabs(i).GroupCount
                If Not modBanner.IsShaded(ws, tabs(i), g) Then
                    ws.Range(ws.Cells(tabs(i).BaseRow + 1, tabs(i).GroupFirst(g)), _
                             ws.Cells(tabs(i).EndRow, tabs(i).GroupLast(g))).Interior.ColorIndex = xlColorIndexNone
                    cleared = cleared + 1
                End If
            Next g
        End If
        modUtil.Progress "인사이트 색을 지우는 중입니다", i, n
    Next i
    modUtil.EndWork

    modConfig.Say "인사이트 색을 지웠습니다. (배너군 " & cleared & "개 범위)", vbInformation
    Exit Sub
Fail:
    modUtil.EndWork
    modConfig.Say "인사이트 색을 지우다가 멈췄습니다." & vbCrLf & vbCrLf & _
                  "오류 " & Err.Number & " : " & Err.Description, vbCritical
End Sub

'==========================================================================
' 칠하기가 판정에 쓰는 한 곳. doPaint = True 이면 셀에 형광을 칠한다.
'==========================================================================
Public Sub ScanInsights(ws As Worksheet, ByVal doPaint As Boolean, _
                        ByRef outTables As Long, ByRef outHits As Long)
    Dim tabs() As TTable, n As Long
    tabs = modTable.ScanTables(ws, n)
    outTables = n
    If n = 0 Then Exit Sub

    Dim mode As String, thr As Double, minBase As Long, blankZero As Boolean, useDisp As Boolean
    mode = modConfig.InsightMode()
    thr = modConfig.InsightValue()
    minBase = modConfig.MinBase()
    blankZero = modConfig.BlankIsZero()
    useDisp = modConfig.UseDisplayValue()



    Dim i As Long, k As Long, c As Long, r As Long
    Dim total As Double, v As Double, sd As Double, cut As Double, diff As Double
    Dim hits As Long

    Dim rowZero As Boolean, digits As Long
    Dim colOK() As Boolean, uc As Long
    ' 색을 칠할 셀은 하나씩 쓰면 느리다. 같은 색끼리 모아 두었다가 한꺼번에 칠한다.
    Dim hiRng As Range, loRng As Range, hiN As Long, loN As Long
    Const CHUNK As Long = 200

    For i = 1 To n
        ' 배너 열이 쓸 만한지는 표마다 한 번만 판정한다. 행마다 다시 물으면
        ' 셀 서식을 그 횟수만큼 물어보게 되어 그것이 병목이 된다.
        If tabs(i).LastCol >= tabs(i).FirstCol Then
            ReDim colOK(tabs(i).FirstCol To tabs(i).LastCol)
            For uc = tabs(i).FirstCol To tabs(i).LastCol
                colOK(uc) = ColumnUsable(ws, tabs(i), uc, minBase)
            Next uc
        Else
            ReDim colOK(0 To 0)          ' 배너가 없는 표에서도 배열은 잡아 둔다
        End If
        ' 보기 행에 이어 【TOP2】 같은 요약 행까지 함께 본다. 연구원이 표에서
        ' 가장 많이 읽는 줄이 그 요약 행이라, 여기서 빠지면 쓸모가 반으로 준다.
        For k = 1 To modTable.ScanRowCount(tabs(i))
            r = modTable.ScanRowAt(tabs(i), k)

            gAtTable = i
            gAtTitle = tabs(i).Title
            gAtRow = r

            ' 평균 점수 행에서는 빈 칸을 0 으로 보지 않는다. 5점 척도에서 0 점은
            ' 있을 수 없는 값이라, 0 으로 두면 없는 차이가 크게 만들어진다.
            Dim lab As String
            lab = RowLabel(ws, tabs(i), r)
            rowZero = blankZero And Not IsScoreRow(lab)
            digits = IIf(IsScoreRow(lab), 2, 1)

            ' 전체 값이 없으면 견줄 기준이 없다.
            If ValueAt(r, tabs(i).TotalCol, rowZero, useDisp, digits, total) Then

                If mode = "STDEV" Then
                    sd = RowStdev(tabs(i), r, colOK, rowZero, useDisp, digits)
                    cut = thr * sd
                Else
                    cut = thr
                End If

                If cut > 0 Then
                    For c = tabs(i).FirstCol To tabs(i).LastCol
                        gAtCol = c
                        If colOK(c) Then
                            If ValueAt(r, c, rowZero, useDisp, digits, v) Then
                                diff = v - total
                                If Abs(diff) > cut Then
                                    hits = hits + 1
                                    If doPaint Then
                                        If diff > 0 Then
                                            AddTo hiRng, ws.Cells(r, c)
                                            hiN = hiN + 1
                                            If hiN >= CHUNK Then
                                                hiRng.Interior.Color = HL_HIGH
                                                Set hiRng = Nothing
                                                hiN = 0
                                            End If
                                        Else
                                            AddTo loRng, ws.Cells(r, c)
                                            loN = loN + 1
                                            If loN >= CHUNK Then
                                                loRng.Interior.Color = HL_LOW
                                                Set loRng = Nothing
                                                loN = 0
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next c
                End If
            End If
        Next k
        modUtil.Progress "인사이트를 찾는 중입니다", i, n
    Next i

    ' 아직 칠하지 않고 모아 둔 것을 마저 칠한다.
    If Not hiRng Is Nothing Then hiRng.Interior.Color = HL_HIGH
    If Not loRng Is Nothing Then loRng.Interior.Color = HL_LOW

    outHits = hits
End Sub

' 칠할 셀을 모아 둔다. 한 덩어리가 너무 커지면 엑셀이 오히려 느려지므로
' 부르는 쪽에서 일정 개수마다 끊어 칠한다.
Private Sub AddTo(ByRef acc As Range, cel As Range)
    If acc Is Nothing Then
        Set acc = cel
    Else
        Set acc = Union(acc, cel)
    End If
End Sub

'--------------------------------------------------------------------------
' 이 열을 볼 것인가 — 미사용(회색) 배너가 아니고 Base 하한을 넘어야 한다.
'--------------------------------------------------------------------------
Private Function ColumnUsable(ws As Worksheet, t As TTable, col As Long, minBase As Long) As Boolean
    If ws.Cells(t.GroupRow, col).Interior.Color = modBanner.SHADE_COLOR Then Exit Function
    ColumnUsable = (modTable.BaseOf(ws, t, col) >= minBase)
End Function

'--------------------------------------------------------------------------
' 한 칸도 안 칠해졌을 때 왜 그런지 짚어 준다.
'
' 볼 수 있는 배너 열이 하나라도 있는데 아무 것도 안 걸렸다면 그것은 정말
' 전체와 다르게 움직이는 값이 없었다는 뜻이라 덧붙일 말이 없다. 문제는 볼 수
' 있는 열 자체가 없는 경우다 - 그 통계표에 손을 대지 않았어도, [보존 배너군]
' 이 예전에 다른 통계표를 볼 때 적어 둔 이름 그대로 레지스트리에 남아 있으면
' 이 통계표의 배너군 이름과 맞지 않아 전부 회색(미사용)으로 걸러진다. 그러면
' 인사이트뿐 아니라 데이터막대도 함께 사라지는데, 사람 눈에는 방금 본 【TOP2】
' 줄만 안 칠해진 것처럼 보이기 쉽다.
'--------------------------------------------------------------------------
Private Function ZeroHitHint(ws As Worksheet, tabs() As TTable, ByVal n As Long, ByVal minBase As Long) As String
    Dim i As Long, c As Long
    Dim total As Long, shaded As Long, usable As Long

    For i = 1 To n
        For c = tabs(i).FirstCol To tabs(i).LastCol
            total = total + 1
            If ws.Cells(tabs(i).GroupRow, c).Interior.Color = modBanner.SHADE_COLOR Then
                shaded = shaded + 1
            ElseIf modTable.BaseOf(ws, tabs(i), c) >= minBase Then
                usable = usable + 1
            End If
        Next c
    Next i

    If total = 0 Then Exit Function      ' 배너 열이 없는 통계표도 있다 - 정상
    If usable > 0 Then Exit Function     ' 볼 수 있는 열이 있는데 안 걸렸다면 정말 벗어난 값이 없는 것이다

    If shaded >= total Then
        ZeroHitHint = vbCrLf & vbCrLf & _
            "[짚이는 원인] 배너 열이 모두 회색(미사용)입니다." & vbCrLf & _
            "[보존 배너군] 은 레지스트리에 남아 다른 통계표를 열어도 그대로" & vbCrLf & _
            "따라옵니다. 이 통계표의 배너군 이름과 맞는지 [검토] → [배너군 목록]" & vbCrLf & _
            "에서 확인하고, [설정] 에서 [보존 배너군] 을 고치거나 비워 보세요."
    Else
        ZeroHitHint = vbCrLf & vbCrLf & _
            "[짚이는 원인] 배너 열의 모수가 모두 Base 하한(" & minBase & ") 미만입니다." & vbCrLf & _
            "[설정] 에서 Base 하한을 낮추거나, 이 통계표가 맞는 파일인지 확인해 주세요."
    End If
End Function

'--------------------------------------------------------------------------
' 셀 하나의 값을 규약대로 읽는다.
'--------------------------------------------------------------------------
' 값은 셀에서 직접 읽지 않고 modTable 이 통째로 떠 놓은 배열에서 가져온다.
' 자릿수도 셀 서식을 묻지 않고 통계표 규약으로 정한다 - % 는 소수 한 자리,
' 평균·환산·NPS 는 두 자리다. 서식을 칸마다 물으면 그것만으로 다시 느려진다.
Private Function ValueAt(ByVal r As Long, ByVal c As Long, _
                         ByVal blankZero As Boolean, ByVal useDisp As Boolean, _
                         ByVal digits As Long, ByRef outVal As Double) As Boolean
    Dim raw As Double
    If Not modTable.Num(r, c, raw) Then
        ' 통계표는 값이 0 인 칸을 빈 칸으로 보낸다.
        If blankZero And modTable.IsBlank(r, c) Then
            outVal = 0
            ValueAt = True
        End If
        Exit Function
    End If

    If useDisp Then
        outVal = Round(raw, digits)
    Else
        outVal = raw
    End If
    ValueAt = True
End Function

'--------------------------------------------------------------------------
' 그 행의 배너 값들이 흩어진 정도. 볼 수 있는 열만 센다.
'--------------------------------------------------------------------------
Private Function RowStdev(t As TTable, r As Long, colOK() As Boolean, _
                          ByVal blankZero As Boolean, ByVal useDisp As Boolean, _
                          ByVal digits As Long) As Double
    Dim c As Long, v As Double, n As Long, s As Double, ss As Double
    For c = t.FirstCol To t.LastCol
        If colOK(c) Then
            If ValueAt(r, c, blankZero, useDisp, digits, v) Then
                n = n + 1
                s = s + v
                ss = ss + v * v
            End If
        End If
    Next c
    If n < 3 Then Exit Function          ' 두 칸으로 잰 표준편차는 믿을 것이 못 된다
    Dim var_ As Double
    var_ = (ss - (s * s) / n) / (n - 1)
    If var_ > 0 Then RowStdev = Sqr(var_)
End Function

Public Function RowLabel(ws As Worksheet, t As TTable, ByVal r As Long) As String
    Dim k As Long, s As String
    For k = t.LabelCol To t.TotalCol - 1
        If Len(modTable.Txt(r, k)) > 0 Then s = modTable.Txt(r, k)
    Next k
    RowLabel = s
End Function

' 평균·점수 행인가. [5점 평균]·[100점 환산]·[NPS] 처럼 % 가 아닌 줄이다.
Private Function IsScoreRow(ByVal lab As String) As Boolean
    Dim s As String
    s = modConfig.Squash(lab)
    ' '점' 하나만으로 가리면 '만족스러운 점' 같은 보기까지 걸리므로 쓰지 않는다.
    IsScoreRow = (InStr(s, "평균") > 0) Or (InStr(s, "환산") > 0) Or (InStr(s, "nps") > 0)
End Function

Public Function RuleText() As String
    If modConfig.InsightMode() = "ABS" Then
        RuleText = "절대값 " & modConfig.InsightValue() & "%p 초과"
    Else
        RuleText = "표준편차 " & modConfig.InsightValue() & "배 초과"
    End If
End Function

'--------------------------------------------------------------------------
' 이 색이 인사이트로 칠한 색인가.
'
' 다른 기능이 「1~3 을 이미 돌린 시트인가」를 가늠할 때 부른다. 예전판이
' 칠하던 노랑도 함께 보아, 옛 파일을 다시 열어도 손댄 시트로 알아본다.
'--------------------------------------------------------------------------
Public Function IsInsightColor(ByVal clr As Long) As Boolean
    IsInsightColor = (clr = HL_HIGH) Or (clr = HL_LOW) Or (clr = HL_OLD)
End Function

' 빌드가 모든 모듈을 실제로 컴파일하도록 부르는 자리다. VBA 는 필요할 때만
' 컴파일하기 때문에, 한 번도 불리지 않은 모듈의 오류는 사람이 그 단추를 누르는
' 순간에야 드러난다. 이 함수가 그 구멍을 막는다.
Public Function Ping() As String
    Ping = "modInsight"
End Function
