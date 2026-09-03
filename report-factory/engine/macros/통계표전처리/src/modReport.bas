Attribute VB_Name = "modReport"
'==========================================================================
' modReport — 검토 산출물 세 가지
'
' 매크로가 무엇을 했는지 사람이 확인하고 남에게 설명할 수 있어야 실무에서 쓴다.
' 색칠만 해 놓으면 「왜 이 배너를 뺐느냐」는 물음에 답할 자리가 없다.
'
'   배너군 목록   이 통계표에 어떤 배너군이 몇 번 나오고 모수는 어떤지 한 장에
'                 모은다. 보존 목록에 적을 이름을 여기서 그대로 복사한다.
'   선정 보고서   표마다 배너군을 남겼는지 뺐는지, 그 사유가 무엇인지 적는다.
'   인사이트 목록 형광으로 칠한 칸을 차이가 큰 순서로 늘어놓는다. 요약을 쓸 때
'                 이 표를 보고 고른다.
'
' 세 산출물은 모두 기능 1·3 과 같은 판정 함수를 부르므로 결과가 어긋날 수 없다.
'==========================================================================
Option Explicit

'==========================================================================
' 1) 배너군 목록
'==========================================================================
Public Sub BannerCatalog()
    Dim ws As Worksheet
    If Not modUtil.RequireSheet(ws) Then Exit Sub
    Dim tabs() As TTable, n As Long
    tabs = modTable.ScanTables(ws, n)
    If n = 0 Then
        modConfig.Say "이 시트에서 통계표를 찾지 못했습니다.", vbExclamation
        Exit Sub
    End If

    Dim minBase As Long
    minBase = modConfig.MinBase()

    ' 이름별로 모은다. 순서는 시트에 나온 차례를 지킨다.
    Dim names As Object, stat As Object, order As Object
    Set names = CreateObject("Scripting.Dictionary")
    Set stat = CreateObject("Scripting.Dictionary")
    Set order = CreateObject("Scripting.Dictionary")

    modUtil.BeginWork "배너군을 모으는 중입니다"
    Dim i As Long, g As Long, key As String, okC As Long, allC As Long
    For i = 1 To n
        For g = 1 To tabs(i).GroupCount
            key = modConfig.Squash(tabs(i).GroupName(g))
            If Not names.Exists(key) Then
                names(key) = tabs(i).GroupName(g)
                order(key) = order.Count
                stat(key) = Array(0#, 0#, 0#, 0#, "")   ' 표 수, 항목 수, 통과 수, base 합, 보기
            End If
            modBanner.BaseTally ws, tabs(i), g, minBase, okC, allC

            Dim a As Variant
            a = stat(key)
            a(0) = a(0) + 1
            a(1) = a(1) + allC
            a(2) = a(2) + okC
            Dim c As Long
            For c = tabs(i).GroupFirst(g) To tabs(i).GroupLast(g)
                a(3) = a(3) + modTable.BaseOf(ws, tabs(i), c)
            Next c
            If Len(a(4)) = 0 Then a(4) = SampleItems(ws, tabs(i), g)
            stat(key) = a
        Next g
        modUtil.Progress "배너군을 모으는 중입니다", i, n
    Next i

    Dim rpt As Worksheet
    Set rpt = modUtil.FreshSheet(ws.Parent, "배너군 목록")
    rpt.Cells(1, 1).Value = "배너군 목록  (원본 시트 : " & ws.name & " · 표 " & n & "장 · Base 하한 " & minBase & ")"
    rpt.Cells(1, 1).Font.Bold = True
    modUtil.HeaderRow rpt, 3, Array("배너군 이름", "나온 표 수", "항목 수(평균)", _
                                    "Base " & minBase & " 통과율", "평균 Base", "판정", "항목 보기")

    Dim r As Long, k As Variant, keepAll As String
    r = 4
    For Each k In names.Keys
        Dim s As Variant
        s = stat(k)
        Dim ratio As Double, avgBase As Double, avgItems As Double
        If s(1) > 0 Then ratio = s(2) / s(1)
        If s(1) > 0 Then avgBase = s(3) / s(1)
        If s(0) > 0 Then avgItems = s(1) / s(0)

        rpt.Cells(r, 1).Value = names(k)
        rpt.Cells(r, 2).Value = s(0)
        rpt.Cells(r, 3).Value = Round(avgItems, 1)
        rpt.Cells(r, 4).Value = ratio
        rpt.Cells(r, 4).NumberFormat = "0%"
        rpt.Cells(r, 5).Value = Round(avgBase, 0)
        If ratio >= 0.5 Then
            rpt.Cells(r, 6).Value = "읽을 수 있음"
            If Len(keepAll) > 0 Then keepAll = keepAll & ", "
            keepAll = keepAll & names(k)
        Else
            rpt.Cells(r, 6).Value = "모수 부족"
            rpt.Cells(r, 6).Font.Color = RGB(160, 0, 0)
        End If
        rpt.Cells(r, 7).Value = s(4)
        r = r + 1
    Next k

    ' 바로 복사해 [보존 배너군] 칸에 붙여 넣을 수 있게 한 줄로도 적어 둔다.
    rpt.Cells(r + 1, 1).Value = "모수를 갖춘 배너군 (복사해서 리본의 [보존 배너군] 칸에 붙여 넣으세요)"
    rpt.Cells(r + 1, 1).Font.Bold = True
    rpt.Cells(r + 2, 1).Value = keepAll
    rpt.Cells(r + 2, 1).Font.Color = RGB(0, 90, 160)

    Finish rpt, 7, 3
    modUtil.EndWork
    rpt.Activate
    modConfig.Say "배너군 " & names.Count & "종을 '배너군 목록' 시트에 모았습니다." & vbCrLf & vbCrLf & _
                  "맨 아래 한 줄을 복사해 리본의 [보존 배너군] 칸에 붙여 넣은 뒤," & vbCrLf & _
                  "보고서에 쓰지 않을 이름만 지우시면 됩니다.", vbInformation
End Sub

'==========================================================================
' 2) 선정 보고서
'==========================================================================
Public Sub SelectionReport()
    Dim ws As Worksheet
    If Not modUtil.RequireSheet(ws) Then Exit Sub
    Dim tabs() As TTable, n As Long
    tabs = modTable.ScanTables(ws, n)
    If n = 0 Then
        modConfig.Say "이 시트에서 통계표를 찾지 못했습니다.", vbExclamation
        Exit Sub
    End If

    Dim keep As String, minBase As Long
    keep = modConfig.KeepGroups()
    minBase = modConfig.MinBase()

    Dim rpt As Worksheet
    modUtil.BeginWork "선정 보고서를 만드는 중입니다"
    Set rpt = modUtil.FreshSheet(ws.Parent, "선정 보고서")
    rpt.Cells(1, 1).Value = "배너 선정 보고서  (원본 시트 : " & ws.name & ")"
    rpt.Cells(1, 1).Font.Bold = True
    rpt.Cells(2, 1).Value = "기준 - 보존 목록 : " & IIf(Len(Trim$(keep)) = 0, "(지정 안 함)", keep) & _
                            "   /   Base 하한 : " & minBase & " 이상이 항목의 절반 이상" & _
                            "   /   판정 단위 : 통계표 전체 (배너군 이름별로 한 번)"
    modUtil.HeaderRow rpt, 4, Array("표 번호", "표 제목", "배너군", "이 표의 항목 수", _
                                    "이 표의 Base 통과", "전체 Base 통과", "보존 목록", "판정", "사유")

    ' 판정은 표마다 따로 내리지 않는다. 시트 전체를 모은 결과를 그대로 쓴다.
    Dim dec As Object
    Set dec = modBanner.BuildDecisions(ws, tabs, n, keep, minBase)

    Dim r As Long, i As Long, g As Long, okC As Long, allC As Long
    Dim kept As Long, dropped As Long
    r = 5
    For i = 1 To n
        For g = 1 To tabs(i).GroupCount
            modBanner.BaseTally ws, tabs(i), g, minBase, okC, allC

            Dim a As Variant, inKeep As Boolean, pass As Boolean, allOk As Long, allAll As Long
            a = dec(modConfig.Squash(tabs(i).GroupName(g)))
            allOk = CLng(a(1)): allAll = CLng(a(2))
            inKeep = CBool(a(3)): pass = CBool(a(4))

            rpt.Cells(r, 1).Value = i
            rpt.Cells(r, 2).Value = tabs(i).Title
            rpt.Cells(r, 3).Value = tabs(i).GroupName(g)
            rpt.Cells(r, 4).Value = allC
            rpt.Cells(r, 5).Value = okC & " / " & allC
            rpt.Cells(r, 6).Value = allOk & " / " & allAll
            rpt.Cells(r, 7).Value = IIf(Len(Trim$(keep)) = 0, "(미지정)", IIf(inKeep, "포함", "제외"))
            rpt.Cells(r, 8).Value = IIf(pass, "선정", "미사용")
            If pass Then
                rpt.Cells(r, 9).Value = "두 관문 통과"
                kept = kept + 1
            Else
                dropped = dropped + 1
                rpt.Cells(r, 8).Font.Color = RGB(160, 0, 0)
                If Not inKeep Then
                    rpt.Cells(r, 9).Value = "보존 목록에 없음 (보고서에 쓰지 않기로 한 배너)"
                Else
                    rpt.Cells(r, 9).Value = "통계표 전체에서 Base " & minBase & " 이상인 항목이 " & _
                                            allOk & "/" & allAll & " 로 절반에 못 미침"
                    If okC * 2 >= allC And allC > 0 Then
                        rpt.Cells(r, 9).Value = rpt.Cells(r, 9).Value & _
                            "  (이 표에는 모수를 갖춘 항목만 실려 있으나, 배너군 전체로는 못 미침)"
                    End If
                End If
            End If
            r = r + 1
        Next g
        modUtil.Progress "선정 보고서를 만드는 중입니다", i, n
    Next i

    rpt.Cells(2, 8).Value = "선정 " & kept & "건 · 미사용 " & dropped & "건"
    rpt.Cells(2, 8).Font.Bold = True
    Finish rpt, 9, 4
    modUtil.EndWork
    rpt.Activate
    modConfig.Say "'선정 보고서' 시트를 만들었습니다. (선정 " & kept & "건 · 미사용 " & dropped & "건)" & vbCrLf & vbCrLf & _
                  "이 시트는 판정 근거를 남기려는 것이므로, 제작팀에 보낼 때는" & vbCrLf & _
                  "지우고 보내셔도 됩니다.", vbInformation
End Sub

'==========================================================================
' 3) 인사이트 목록
'==========================================================================
Public Sub InsightList()
    Dim ws As Worksheet
    If Not modUtil.RequireSheet(ws) Then Exit Sub
    Dim rpt As Worksheet
    modUtil.BeginWork "인사이트 목록을 만드는 중입니다"
    Set rpt = modUtil.FreshSheet(ws.Parent, "인사이트 목록")
    rpt.Cells(1, 1).Value = "인사이트 목록  (원본 시트 : " & ws.name & " · 기준 : " & modInsight.RuleText() & _
                            " · Base 하한 " & modConfig.MinBase() & ")"
    rpt.Cells(1, 1).Font.Bold = True
    modUtil.HeaderRow rpt, 3, Array("표 제목", "보기(행)", "배너군", "배너 항목", _
                                    "Base", "전체", "이 집단", "차이(%p)", "방향", "셀 위치")

    ' 두 번 훑는다. 처음에는 몇 줄인지만 세고, 그 크기로 자리를 잡은 뒤 두 번째에
    ' 채운다. 시트에 한 줄씩 쓰는 것보다 이 편이 훨씬 빠르다.
    Dim nTab As Long, nHit As Long
    modInsight.ScanInsights ws, False, False, nTab, nHit
    If nHit > 0 Then
        modInsight.PrepareBuffer nHit
        Dim nTab2 As Long, nHit2 As Long
        modInsight.ScanInsights ws, False, True, nTab2, nHit2
        rpt.Range(rpt.Cells(4, 1), rpt.Cells(3 + nHit, 10)).Value = modInsight.Buffer()
    End If
    modUtil.EndWork

    If nTab = 0 Then
        modConfig.Say "이 시트에서 통계표를 찾지 못했습니다.", vbExclamation
        Exit Sub
    End If

    If nHit > 0 Then
        ' 차이가 큰 순서로 세운다. 요약에 쓸 것을 위에서부터 고르면 된다.
        With rpt.Sort
            .SortFields.Clear
            .SortFields.Add key:=rpt.Range(rpt.Cells(4, 8), rpt.Cells(3 + nHit, 8)), Order:=xlDescending
            .SetRange rpt.Range(rpt.Cells(3, 1), rpt.Cells(3 + nHit, 10))
            .Header = xlYes
            .Apply
        End With
        rpt.Range(rpt.Cells(4, 6), rpt.Cells(3 + nHit, 8)).NumberFormat = "0.0"
    End If

    Finish rpt, 10, 3
    rpt.Activate
    modConfig.Say "인사이트 " & nHit & "건을 '인사이트 목록' 시트에 모았습니다." & vbCrLf & vbCrLf & _
                  "차이가 큰 순서로 세워 두었습니다. 다만 순서는 크기일 뿐이고," & vbCrLf & _
                  "요약에 올릴 값어치는 의뢰인의 물음이 정합니다.", vbInformation
End Sub

'==========================================================================
Private Function SampleItems(ws As Worksheet, t As TTable, g As Long) As String
    Dim c As Long, s As String, cnt As Long, v As String
    For c = t.GroupFirst(g) To t.GroupLast(g)
        If t.HeadRow > 0 Then v = Trim$(CStr(ws.Cells(t.HeadRow, c).Value & ""))
        If Len(v) > 0 Then
            If Len(s) > 0 Then s = s & " · "
            s = s & v & "(" & modTable.BaseOf(ws, t, c) & ")"
            cnt = cnt + 1
            If cnt >= 4 Then
                s = s & " …"
                Exit For
            End If
        End If
    Next c
    SampleItems = s
End Function

' 열 너비를 맞추고 머리글에 자동 필터를 건다. 시트를 고르지 않고 처리하므로
' 화면 갱신을 멈춘 상태에서도 안전하다.
Private Sub Finish(rpt As Worksheet, ByVal lastCol As Long, ByVal headRow As Long)
    On Error Resume Next
    rpt.Range(rpt.Cells(1, 1), rpt.Cells(1, lastCol)).EntireColumn.AutoFit
    Dim c As Long
    For c = 1 To lastCol
        If rpt.Columns(c).ColumnWidth > 42 Then rpt.Columns(c).ColumnWidth = 42
        If rpt.Columns(c).ColumnWidth < 8 Then rpt.Columns(c).ColumnWidth = 8
    Next c
    rpt.Range(rpt.Cells(headRow, 1), rpt.Cells(headRow, lastCol)).AutoFilter
    On Error GoTo 0
End Sub

' 빌드가 모든 모듈을 실제로 컴파일하도록 부르는 자리다. VBA 는 필요할 때만
' 컴파일하기 때문에, 한 번도 불리지 않은 모듈의 오류는 사람이 그 단추를 누르는
' 순간에야 드러난다. 이 함수가 그 구멍을 막는다.
Public Function Ping() As String
    Ping = "modReport"
End Function
