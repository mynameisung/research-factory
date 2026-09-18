Attribute VB_Name = "modReport"
'==========================================================================
' modReport — 배너군 목록
'
' 매크로가 무엇을 했는지 사람이 확인하고 남에게 설명할 수 있어야 실무에서 쓴다.
'
'   배너군 목록   이 통계표에 어떤 배너군이 몇 번 나오고 모수는 어떤지 한 장에
'                 모은다. 보존 목록에 적을 이름을 여기서 그대로 복사한다.
'                 문항이 여럿이면 팝업에서 볼 문항을 고른다(기본은 전체 선택).
'                 보통은 전체 문항 기준으로 뽑지만, 특정 문항 하나의 배너군
'                 구성만 보고 싶을 때는 그 문항만 골라 좁혀 볼 수 있다.
'
' [배너 선정] 과 같은 판정 함수(modBanner.BaseTally)를 쓰므로 결과가 어긋날 수 없다.
'==========================================================================
Option Explicit

' [배너군 목록] 팝업이 뜬 동안 쓰는 임시 자리. 스캔은 [배너군 목록] 을 누른
' 그 순간 한 번만 하고, 팝업에서 [확인] 을 누르면 그 결과를 그대로 이어 쓴다.
Private gCatWs As Worksheet
Private gCatTabs() As TTable
Private gCatN As Long

'==========================================================================
' 배너군 목록 — 문항을 고르는 팝업을 먼저 띄운다.
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

    Set gCatWs = ws
    gCatTabs = tabs
    gCatN = n

    Dim titles() As String, i As Long
    ReDim titles(1 To n)
    For i = 1 To n
        titles(i) = tabs(i).Title
    Next i

    frmCatalogPick.Populate titles
    frmCatalogPick.Show
End Sub

'==========================================================================
' [배너군 목록] 팝업에서 [확인] 을 누르면 여기로 온다. picked(i) 가 True 인
' 표만 모아 배너군을 뽑는다. BannerCatalog 가 스캔해 둔 gCatWs/gCatTabs 를
' 그대로 쓰므로 다시 스캔하지 않는다.
'==========================================================================
Public Sub RunCatalogFiltered(picked() As Boolean)
    If gCatN = 0 Or gCatWs Is Nothing Then Exit Sub
    Dim ws As Worksheet, tabs() As TTable, n As Long
    Set ws = gCatWs
    tabs = gCatTabs
    n = gCatN

    Dim minBase As Long
    minBase = modConfig.MinBase()

    ' 이름별로 모은다. 순서는 시트에 나온 차례를 지킨다.
    Dim names As Object, stat As Object, order As Object
    Set names = CreateObject("Scripting.Dictionary")
    Set stat = CreateObject("Scripting.Dictionary")
    Set order = CreateObject("Scripting.Dictionary")

    modUtil.BeginWork "배너군을 모으는 중입니다"
    Dim i As Long, g As Long, key As String, okC As Long, allC As Long, pickedN As Long
    For i = 1 To n
        If picked(i) Then
            pickedN = pickedN + 1
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
        End If
        modUtil.Progress "배너군을 모으는 중입니다", i, n
    Next i

    If pickedN = 0 Then
        modUtil.EndWork
        modConfig.Say "고른 문항이 없어 배너군 목록을 만들지 못했습니다.", vbExclamation
        Exit Sub
    End If

    Dim rpt As Worksheet
    Set rpt = modUtil.FreshSheet(ws.Parent, "배너군 목록")
    rpt.Cells(1, 1).Value = "배너군 목록  (원본 시트 : " & ws.name & " · 문항 " & pickedN & "/" & n & _
                            "개 선택 · Base 하한 " & minBase & ")"
    rpt.Cells(1, 1).Font.Bold = True
    modUtil.HeaderRow rpt, 3, Array("배너군 이름", "나온 표 수", "항목 수(평균)", _
                                    "Base " & minBase & " 통과율", "평균 Base", "판정", "항목 보기")

    ' 배너 선정과 같은 판정(그 배너군 항목의 절반 이상이 Base 를 넘는가)을
    ' 여기서도 그대로 쓴다. 이름별로 한 번만 정해, 뒤에서 시트 전체에
    ' 똑같이 적용한다(표마다 다르게 판정하면 같은 배너가 남았다 없어졌다 한다).
    Dim passByName As Object
    Set passByName = CreateObject("Scripting.Dictionary")

    Dim r As Long, k As Variant, keepAll As String
    r = 4
    For Each k In names.Keys
        Dim s As Variant
        s = stat(k)
        Dim ratio As Double, avgBase As Double, avgItems As Double
        If s(1) > 0 Then ratio = s(2) / s(1)
        If s(1) > 0 Then avgBase = s(3) / s(1)
        If s(0) > 0 Then avgItems = s(1) / s(0)
        passByName(k) = (ratio >= 0.5)

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

    ' [배너 선정] 처럼 이 판정을 시트에도 그대로 입힐지 물은 뒤 적용한다.
    ' 원본 시트를 건드리는 일이라 보고서만 볼 때와 달리 한 번 확인을 받는다.
    Dim shadeMsg As String
    If modConfig.Ask( _
        "지금 나온 판정(읽을 수 있음 / 모수 부족)을 [배너 선정] 처럼" & vbCrLf & _
        "원본 시트에도 회색 음영으로 입힐까요?" & vbCrLf & vbCrLf & _
        "문항 " & pickedN & "개 기준으로 정한 것이라, [설정] 의 [보존 배너군]" & vbCrLf & _
        "텍스트 목록과는 다른 결과가 나올 수 있습니다. 이미 칠해 둔 음영이" & vbCrLf & _
        "있다면 이 판정으로 다시 칠해 덮어씁니다.", "배너군 목록 → 음영 적용") Then

        modUtil.BeginWork "판정대로 음영을 입히는 중입니다"
        Dim kept2 As Long, dropped2 As Long
        For i = 1 To n
            For g = 1 To tabs(i).GroupCount
                Dim gk As String
                gk = modConfig.Squash(tabs(i).GroupName(g))
                If passByName.Exists(gk) Then
                    If passByName(gk) Then
                        modBanner.ClearShade ws, tabs(i), g
                        kept2 = kept2 + 1
                    Else
                        modBanner.Shade ws, tabs(i), g
                        modBars.ClearBarsIn ws.Range(ws.Cells(tabs(i).BaseRow, tabs(i).GroupFirst(g)), _
                                                      ws.Cells(tabs(i).EndRow, tabs(i).GroupLast(g)))
                        dropped2 = dropped2 + 1
                    End If
                End If
            Next g
            modUtil.Progress "판정대로 음영을 입히는 중입니다", i, n
        Next i
        modUtil.EndWork

        shadeMsg = vbCrLf & vbCrLf & "시트에도 적용했습니다 — 선정 " & kept2 & "개 · 음영 " & dropped2 & "개." & vbCrLf & _
                   "[제외 배너 삭제] 로 지울 수 있습니다."
    End If

    modConfig.Say "문항 " & pickedN & "개(전체 " & n & "개 가운데)에서 배너군 " & names.Count & "종을 " & _
                  "'배너군 목록' 시트에 모았습니다." & vbCrLf & vbCrLf & _
                  "맨 아래 한 줄을 복사해 리본의 [보존 배너군] 칸에 붙여 넣은 뒤," & vbCrLf & _
                  "보고서에 쓰지 않을 이름만 지우시면 됩니다." & shadeMsg, vbInformation
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
