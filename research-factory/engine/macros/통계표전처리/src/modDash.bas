Attribute VB_Name = "modDash"
'==========================================================================
' modDash — 기능 4. 빈 칸을 - 로 채우기
'
' 통계표는 값이 0 인 칸을 0.0 이 아니라 아예 빈 칸으로 보낸다. 그대로 두면
' 값이 0 인 것인지 물어보지 않은 칸인지 읽는 사람이 알 수 없으므로, 제작가이드로
' 옮기기 전에 - 로 채워 둔다.
'
' ★ 이 작업은 반드시 맨 마지막에 한다.
'   - 는 숫자가 아니라 글자다. 먼저 채워 버리면 데이터막대가 그려지지 않고,
'   인사이트에서도 그 칸이 통째로 빠진다. 그래서 리본에서도 맨 오른쪽에 두었고,
'   [한 번에 실행] 에도 넣지 않았다. 사람이 마지막에 직접 누르는 자리다.
'
' 채우는 범위는 표의 몸통 전체다. Base 행 아래부터 표의 끝까지, 전체 열과 배너
' 열을 모두 본다. 계 행과 【TOP2】 같은 요약 행도 표의 내용이므로 함께 채운다.
'==========================================================================
Option Explicit

Public Const DASH As String = "-"

'--------------------------------------------------------------------------
Public Sub FillDashes()
    On Error GoTo Fail
    Dim ws As Worksheet
    If Not modUtil.RequireSheet(ws) Then Exit Sub

    Dim tabs() As TTable, n As Long
    tabs = modTable.ScanTables(ws, n)
    If n = 0 Then
        modConfig.Say "이 시트에서 통계표를 찾지 못했습니다.", vbExclamation
        Exit Sub
    End If

    ' 1~3 을 아직 하지 않은 것으로 보이면 더 세게 말린다. 순서를 거꾸로 하면
    ' 막대도 형광도 얻지 못한 채 표만 글자로 굳어 버리기 때문이다.
    Dim warn As String, notYet As Boolean
    notYet = Not LooksProcessed(ws, tabs, n)

    warn = "표 안의 빈 칸을 " & DASH & " 로 채웁니다. (표 " & n & "장)" & vbCrLf & vbCrLf
    warn = warn & "이 작업은 맨 마지막에 하십시오." & vbCrLf & vbCrLf
    warn = warn & DASH & " 는 숫자가 아니라 글자입니다. 채우고 나면 그 칸에는" & vbCrLf
    warn = warn & "데이터막대가 그려지지 않고 인사이트에서도 빠집니다." & vbCrLf & vbCrLf
    If notYet Then
        warn = warn & "[확인 필요] 이 시트에는 아직 데이터막대도 형광도 없습니다." & vbCrLf
        warn = warn & "1. 배너 선정 → 2. 데이터막대 → 3. 인사이트 을 먼저" & vbCrLf
        warn = warn & "마치신 뒤에 누르시기를 권합니다." & vbCrLf & vbCrLf
    End If
    warn = warn & "지금 채우시겠습니까?"
    If Not modConfig.Ask(warn, "빈 칸 - 로 채우기") Then Exit Sub

    ' 표 한 장의 몸통을 배열로 받아 고친 뒤 한 번에 되돌려 쓴다. 칸마다 쓰면
    ' 표 하나에 수백 번, 통계표 하나에 수만 번을 엑셀과 주고받게 된다.
    modUtil.BeginWork "빈 칸을 채우는 중입니다"
    Dim i As Long, r As Long, c As Long, filled As Long
    Dim body As Range, arr As Variant, touched As Boolean
    For i = 1 To n
        Set body = ws.Range(ws.Cells(tabs(i).BaseRow + 1, tabs(i).TotalCol), _
                            ws.Cells(tabs(i).EndRow, tabs(i).LastCol))
        arr = body.Value2
        If IsArray(arr) Then
            touched = False
            For r = 1 To UBound(arr, 1)
                For c = 1 To UBound(arr, 2)
                    If IsBlankValue(arr(r, c)) Then
                        arr(r, c) = DASH
                        filled = filled + 1
                        touched = True
                    End If
                Next c
            Next r
            If touched Then
                body.Value2 = arr
                body.HorizontalAlignment = xlRight
            End If
        ElseIf IsBlankValue(arr) Then
            body.Value2 = DASH
            body.HorizontalAlignment = xlRight
            filled = filled + 1
        End If
        modUtil.Progress "빈 칸을 채우는 중입니다", i, n
    Next i
    modUtil.EndWork

    modConfig.Say "빈 칸 " & filled & "개를 " & DASH & " 로 채웠습니다." & vbCrLf & vbCrLf & _
                  "되돌리려면 [빈 칸 되돌리기] 를 누르시면 됩니다.", vbInformation
    Exit Sub
Fail:
    modUtil.EndWork
    modConfig.Say "빈 칸을 채우다가 멈췄습니다. (표 " & i & " 행 " & r & " 열 " & c & ")" & vbCrLf & vbCrLf & _
                  "오류 " & Err.Number & " : " & Err.Description, vbCritical
End Sub

'--------------------------------------------------------------------------
' 채워 넣은 - 를 다시 빈 칸으로 돌린다. 데이터막대나 인사이트을 다시 하려면
' 먼저 이것을 눌러야 한다.
'--------------------------------------------------------------------------
Public Sub ClearDashes()
    On Error GoTo Fail
    Dim ws As Worksheet
    If Not modUtil.RequireSheet(ws) Then Exit Sub

    Dim tabs() As TTable, n As Long
    tabs = modTable.ScanTables(ws, n)
    If n = 0 Then
        modConfig.Say "이 시트에서 통계표를 찾지 못했습니다.", vbExclamation
        Exit Sub
    End If

    modUtil.BeginWork "- 를 걷어내는 중입니다"
    Dim i As Long, r As Long, c As Long, cleared As Long
    Dim body As Range, arr As Variant, touched As Boolean
    For i = 1 To n
        Set body = ws.Range(ws.Cells(tabs(i).BaseRow + 1, tabs(i).TotalCol), _
                            ws.Cells(tabs(i).EndRow, tabs(i).LastCol))
        arr = body.Value2
        If IsArray(arr) Then
            touched = False
            For r = 1 To UBound(arr, 1)
                For c = 1 To UBound(arr, 2)
                    If Trim$(CStr(arr(r, c) & "")) = DASH Then
                        arr(r, c) = Empty
                        cleared = cleared + 1
                        touched = True
                    End If
                Next c
            Next r
            If touched Then body.Value2 = arr
        End If
        modUtil.Progress "- 를 걷어내는 중입니다", i, n
    Next i
    modUtil.EndWork

    modConfig.Say DASH & " " & cleared & "개를 걷어내 빈 칸으로 되돌렸습니다.", vbInformation
    Exit Sub
Fail:
    modUtil.EndWork
    modConfig.Say "- 를 걷어내다가 멈췄습니다. (표 " & i & " 행 " & r & " 열 " & c & ")" & vbCrLf & vbCrLf & _
                  "오류 " & Err.Number & " : " & Err.Description, vbCritical
End Sub

'==========================================================================
' 1~3 을 이미 돌린 시트로 보이는가. 데이터막대나 형광, 배너 음영 가운데
' 하나라도 있으면 손을 댄 시트로 본다. 앞쪽 표 몇 장만 들춰 보고 판단한다.
'--------------------------------------------------------------------------
Private Function LooksProcessed(ws As Worksheet, tabs() As TTable, ByVal n As Long) As Boolean
    Dim i As Long, g As Long, limit As Long
    Dim body As Range, budget As Long

    limit = n
    If limit > 5 Then limit = 5

    ' 이것은 어림으로 가늠하는 자리다. 판정이 조금 빗나가도 안내 문구가 한 줄
    ' 더 뜰 뿐이지만, 여기서 오래 붙들리면 사람은 매크로가 멈춘 것으로 여기고
    ' 엑셀을 강제로 닫는다. 그래서 훑어보는 칸 수에 상한을 둔다.
    budget = 2000

    For i = 1 To limit
        ' 배너 음영
        For g = 1 To tabs(i).GroupCount
            If modBanner.IsShaded(ws, tabs(i), g) Then
                LooksProcessed = True
                Exit Function
            End If
        Next g

        If tabs(i).EndRow > tabs(i).BaseRow Then
            Set body = ws.Range(ws.Cells(tabs(i).BaseRow + 1, tabs(i).TotalCol), _
                                ws.Cells(tabs(i).EndRow, tabs(i).LastCol))
            ' 데이터막대
            On Error Resume Next
            If body.FormatConditions.Count > 0 Then
                LooksProcessed = True
                On Error GoTo 0
                Exit Function
            End If
            On Error GoTo 0
            ' 형광
            If RangeHasInsight(body, budget) Then
                LooksProcessed = True
                Exit Function
            End If
            If budget <= 0 Then Exit Function
        End If
    Next i
End Function

'==========================================================================
' 이 범위 안에 인사이트 형광이 칠해져 있는가.
'
' 칸마다 색을 물으면 표 한 장에 수백 번, 통계표 하나에 수만 번을 엑셀과
' 주고받게 되어 사람이 기다리지 못한다. Interior.ColorIndex 는 범위가 한 색이면
' 그 색을 돌려주고 색이 섞여 있으면 Null 을 돌려주므로, 그 성질을 써서 한 번의
' 물음으로 대부분을 가른다. 아직 아무것도 칠하지 않은 시트는 여기서 끝난다.
' 색이 섞인 표에서만 행 단위로 좁히고, 그래도 섞인 행에서만 칸을 훑는다.
'--------------------------------------------------------------------------
Private Function RangeHasInsight(body As Range, ByRef budget As Long) As Boolean
    Dim ci As Variant, rr As Long, cel As Range

    ci = ColorIndexOf(body)
    If Not IsNull(ci) Then
        If ci = xlColorIndexNone Then Exit Function
        RangeHasInsight = modInsight.IsInsightColor(body.Interior.Color)
        Exit Function
    End If

    For rr = 1 To body.Rows.Count
        ci = ColorIndexOf(body.Rows(rr))
        If IsNull(ci) Then
            For Each cel In body.Rows(rr).Cells
                budget = budget - 1
                If modInsight.IsInsightColor(cel.Interior.Color) Then
                    RangeHasInsight = True
                    Exit Function
                End If
                If budget <= 0 Then Exit Function
            Next cel
        ElseIf ci <> xlColorIndexNone Then
            If modInsight.IsInsightColor(body.Rows(rr).Interior.Color) Then
                RangeHasInsight = True
                Exit Function
            End If
        End If
    Next rr
End Function

'--------------------------------------------------------------------------
' 범위의 ColorIndex 를 안전하게 묻는다. 병합 셀이 걸쳐 있거나 범위가 커서
' 엑셀이 답하지 못하면 Null 로 보아, 부르는 쪽이 더 좁은 범위로 다시 묻게 한다.
'--------------------------------------------------------------------------
Private Function ColorIndexOf(rng As Range) As Variant
    ColorIndexOf = Null
    On Error Resume Next
    ColorIndexOf = rng.Interior.ColorIndex
    If Err.Number <> 0 Then
        Err.Clear
        ColorIndexOf = Null
    End If
    On Error GoTo 0
End Function

'==========================================================================
' 빈 칸인가 - 수식이 들어 있어 겉보기만 비어 있는 칸은 건드리지 않는다.
'--------------------------------------------------------------------------
Private Function IsBlankValue(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then
        IsBlankValue = True
    ElseIf VarType(v) = vbString Then
        IsBlankValue = (Len(Trim$(CStr(v))) = 0)
    End If
End Function

' 빌드가 모든 모듈을 실제로 컴파일하도록 부르는 자리다. VBA 는 필요할 때만
' 컴파일하기 때문에, 한 번도 불리지 않은 모듈의 오류는 사람이 그 단추를 누르는
' 순간에야 드러난다. 이 함수가 그 구멍을 막는다.
Public Function Ping() As String
    Ping = "modDash"
End Function
