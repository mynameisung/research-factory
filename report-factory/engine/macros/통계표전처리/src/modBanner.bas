Attribute VB_Name = "modBanner"
'==========================================================================
' modBanner — 기능 1. 배너 선정
'
' 문항마다 배너군을 하나씩 훑어 보고서에 쓸 것과 쓰지 않을 것을 가른다.
' 판정은 두 관문을 함께 적용한다.
'
'   1) 보존 목록  리본에 등록해 둔 배너군 이름(예: 용도·권역·관리 운영 방식)에
'                 들어 있어야 한다. 목록이 비어 있으면 이 관문은 건너뛴다.
'   2) Base 하한   그 배너군에 속한 항목 가운데 Base 가 기준값(기본 30) 이상인
'                 항목이 절반 이상이어야 한다. 모수가 모자란 군은 % 를 읽을 수
'                 없으므로 전체 열만 읽는다.
'
' 두 관문은 하는 일이 다르다. 어느 배너를 쓸지는 연구팀이 정하고(보존 목록),
' 그 배너를 읽을 수 있는지는 모수가 정한다(Base 하한).
'
' 통과하지 못한 배너군은 회색으로 음영 처리한다. 음영은 표시일 뿐이고, 실제로
' 열을 지우는 것은 [제외 배너 삭제] 를 눌렀을 때다. 삭제는 시트 전체의 열을
' 지우지 않고 표 한 장의 범위만 왼쪽으로 당긴다. 표마다 배너 열의 위치가 다르기
' 때문에 시트 단위로 지우면 다른 표가 깨진다.
'==========================================================================
Option Explicit

Public Const SHADE_COLOR As Long = 14277081     ' RGB(217, 217, 217) 회색
Private Const BACKUP_PREFIX As String = "삭제전백업_"

'--------------------------------------------------------------------------
' 선정하고 음영을 칠한다.
'--------------------------------------------------------------------------
Public Sub SelectBanners()
    Dim ws As Worksheet
    If Not modUtil.RequireSheet(ws) Then Exit Sub

    Dim keep As String, minBase As Long
    keep = modConfig.KeepGroups()
    minBase = modConfig.MinBase()

    Dim tabs() As TTable, n As Long
    tabs = modTable.ScanTables(ws, n)
    If n = 0 Then
        modConfig.Say "이 시트에서 통계표를 찾지 못했습니다." & vbCrLf & vbCrLf & _
               "A열 근처에 'Base for %' 가 있고 그 위에 '전체' 라벨이 있는" & vbCrLf & _
               "원본 통계표의 Table 시트에서 실행해 주세요.", vbExclamation
        Exit Sub
    End If

    modUtil.BeginWork "배너를 고르는 중입니다"

    ' 판정은 표마다 따로 하지 않고 시트 전체를 모아 배너군 이름 단위로 한 번 내린다.
    Dim dec As Object
    Set dec = BuildDecisions(ws, tabs, n, keep, minBase)

    Dim i As Long, g As Long, kept As Long, dropped As Long, unknown As String
    For i = 1 To n
        For g = 1 To tabs(i).GroupCount
            If DecidedPass(dec, tabs(i).GroupName(g)) Then
                ClearShade ws, tabs(i), g
                kept = kept + 1
            Else
                Shade ws, tabs(i), g
                ' 순서를 바꿔 눌렀을 때를 대비해, 미사용 배너에 남아 있던 막대를 거둔다.
                modBars.ClearBarsIn ws.Range(ws.Cells(tabs(i).BaseRow, tabs(i).GroupFirst(g)), _
                                             ws.Cells(tabs(i).EndRow, tabs(i).GroupLast(g)))
                dropped = dropped + 1
            End If
        Next g
        modUtil.Progress "배너를 고르는 중입니다", i, n
    Next i
    modUtil.EndWork

    ' 보존 목록에 적었는데 시트에서 한 번도 안 나온 이름은 오타일 가능성이 높다.
    unknown = UnmatchedKeepNames(ws, tabs, n, keep)

    Dim msg As String
    msg = "표 " & n & "장을 훑었습니다." & vbCrLf & vbCrLf & _
          "선정한 배너군 : " & kept & "개" & vbCrLf & _
          "음영 처리한 배너군 : " & dropped & "개" & vbCrLf & vbCrLf & _
          "이번에 적용한 기준" & vbCrLf & _
          "1. 보존 목록 : " & IIf(Len(Trim$(keep)) = 0, "(지정 안 함 - Base 기준만 적용)", keep) & vbCrLf & _
          "2. Base 하한 : 배너 항목의 절반 이상이 " & minBase & " 이상이어야 남긴다" & vbCrLf & _
          "3. 판정 단위 : 표 한 장이 아니라 통계표 전체(배너군 이름별로 한 번)." & vbCrLf & _
          "              같은 배너군이 표마다 남았다 없어졌다 하지 않는다"
    If Len(unknown) > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "[확인 필요] 보존 목록에 적었지만 이 시트에 없는 이름이 있습니다." & vbCrLf & _
              "  " & unknown & vbCrLf & _
              "[검토] → [배너군 목록] 을 눌러 실제 이름을 확인해 주세요."
    End If
    If kept = 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "[확인 필요] 남은 배너군이 하나도 없습니다. 보존 목록의 이름이나" & vbCrLf & _
              "Base 하한을 다시 보시기 바랍니다."
    End If
    modConfig.Say msg, vbInformation
End Sub

'--------------------------------------------------------------------------
' 음영이 칠해진 배너군을 표마다 지우고 왼쪽으로 당긴다.
'--------------------------------------------------------------------------
Public Sub DeleteShadedBanners()
    Dim ws As Worksheet
    If Not modUtil.RequireSheet(ws) Then Exit Sub

    Dim tabs() As TTable, n As Long
    tabs = modTable.ScanTables(ws, n)
    If n = 0 Then
        modConfig.Say "이 시트에서 통계표를 찾지 못했습니다.", vbExclamation
        Exit Sub
    End If

    Dim i As Long, g As Long, target As Long
    For i = 1 To n
        For g = 1 To tabs(i).GroupCount
            If IsShaded(ws, tabs(i), g) Then target = target + 1
        Next g
    Next i

    If target = 0 Then
        modConfig.Say "지울 음영 배너가 없습니다. [배너 선정] 을 먼저 실행해 주세요.", vbInformation
        Exit Sub
    End If

    Dim warn As String
    warn = "음영 처리된 배너군 " & target & "개를 지웁니다." & vbCrLf & vbCrLf
    If modConfig.BackupBeforeDelete() Then
        warn = warn & "지우기 전에 이 시트의 사본을 한 장 만들어 둡니다." & vbCrLf
    Else
        warn = warn & "설정의 [삭제 전 백업] 이 꺼져 있어 되돌릴 수 없습니다." & vbCrLf
    End If
    warn = warn & "매크로가 한 일은 Ctrl+Z 로 되돌아가지 않습니다." & vbCrLf & vbCrLf & _
                  "계속하시겠습니까?"
    If Not modConfig.Ask(warn, "배너 미사용 삭제") Then Exit Sub

    Dim backupName As String
    If modConfig.BackupBeforeDelete() Then backupName = MakeBackup(ws)

    modUtil.BeginWork "미사용 배너를 지우는 중입니다"
    Dim done As Long, failed As Long, oldLastCol As Long
    ' 지우기 전 시트의 오른쪽 끝을 기억해 둔다. 나중에 표 오른쪽에 남는 테두리와
    ' 음영을 그 자리까지 걷어내야 하기 때문이다.
    oldLastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    ' 아래 표부터, 그리고 오른쪽 배너군부터 지워야 앞서 잡아 둔 좌표가 어긋나지 않는다.
    For i = n To 1 Step -1
        For g = tabs(i).GroupCount To 1 Step -1
            If IsShaded(ws, tabs(i), g) Then
                Dim rng As Range
                ' 제목행부터 함께 당긴다. 제목행에는 표 폭만큼 테두리와 음영이
                ' 깔려 있어서, 몸통만 지우면 제목 줄만 예전 폭으로 남는다.
                Set rng = ws.Range(ws.Cells(tabs(i).TitleRow, tabs(i).GroupFirst(g)), _
                                   ws.Cells(tabs(i).EndRow, tabs(i).GroupLast(g)))
                modUtil.UnmergeIfNeeded rng      ' 병합이 걸쳐 있으면 삭제가 멎는다
                On Error Resume Next
                rng.Delete Shift:=xlToLeft
                If Err.Number <> 0 Then
                    failed = failed + 1
                    Err.Clear
                Else
                    done = done + 1
                End If
                On Error GoTo 0
            End If
        Next g
        modUtil.Progress "미사용 배너를 지우는 중입니다", n - i + 1, n
    Next i

    ' 표가 좁아지면서 오른쪽에 테두리와 음영만 남는다. 표마다 새 폭을 다시 재고
    ' 그 바깥을 걷어낸다. 값은 이미 없으므로 서식만 지운다.
    Dim tabs2() As TTable, n2 As Long, k As Long
    tabs2 = modTable.ScanTables(ws, n2)
    For k = 1 To n2
        If oldLastCol > tabs2(k).LastCol Then
            ws.Range(ws.Cells(tabs2(k).TitleRow, tabs2(k).LastCol + 1), _
                     ws.Cells(tabs2(k).EndRow, oldLastCol)).ClearFormats
        End If
        modUtil.Progress "표 오른쪽을 정리하는 중입니다", k, n2
    Next k
    modUtil.EndWork

    Dim msg As String
    msg = "배너군 " & done & "개를 지웠습니다. 선정한 배너가 전체 열 오른쪽으로 모였습니다."
    If failed > 0 Then
        msg = msg & vbCrLf & vbCrLf & "[확인 필요] " & failed & "개는 지우지 못했습니다." & vbCrLf & _
                    "병합된 셀이나 보호된 범위가 걸쳐 있을 수 있습니다."
    End If
    If Len(backupName) > 0 Then
        msg = msg & vbCrLf & vbCrLf & "지우기 전 사본을 '" & backupName & "' 시트에 남겼습니다." & vbCrLf & _
                    "확인이 끝나면 그 시트는 지우고 저장하시면 됩니다."
    End If
    modConfig.Say msg, vbInformation
End Sub

'--------------------------------------------------------------------------
' 음영만 모두 걷어낸다 (기준을 바꿔 다시 선정할 때).
'--------------------------------------------------------------------------
Public Sub ClearAllShade()
    Dim ws As Worksheet, tabs() As TTable, n As Long, i As Long, g As Long
    If Not modUtil.RequireSheet(ws) Then Exit Sub

    modUtil.BeginWork "음영을 걷어내는 중입니다"
    tabs = modTable.ScanTables(ws, n)
    For i = 1 To n
        For g = 1 To tabs(i).GroupCount
            ClearShade ws, tabs(i), g
        Next g
        modUtil.Progress "음영을 걷어내는 중입니다", i, n
    Next i
    modUtil.EndWork
    modConfig.Say "배너 음영을 모두 걷어냈습니다.", vbInformation
End Sub

'==========================================================================
' 판정 — 배너군 이름 하나에 대해 통계표 전체를 놓고 한 번만 정한다.
'
' 표 한 장만 보고 정하면 안 된다. DP 는 표마다 해당 없는 배너 범주를 빼고
' 보내므로, 어떤 표에는 그 배너군의 항목이 하나만 실려 온다. 예를 들어
' 'A1-1. 기업 형태' 표에는 '건물 소유 형태' 배너군이 '기업 소유(54)' 한 칸만
' 붙어 있다. 표 단위로 보면 1/1 통과라 남지만, 그 배너군 전체로 보면
' 개인(9)·구분(13)·국공유(3)·기타(7) 가 모수에 못 미쳐 보고서에 쓸 수 없는
' 배너다. 쓰지 않기로 한 배너가 표마다 남았다 없어졌다 하면 안 되므로,
' 시트 전체를 모아 한 번 정하고 그 결정을 모든 표에 똑같이 적용한다.
'
' 돌려주는 것은 사전이다.  키 = 납작하게 만든 배너군 이름
'                          값 = Array(원래이름, 통과항목수, 전체항목수,
'                                     보존목록해당, 최종판정)
'==========================================================================
Public Function BuildDecisions(ws As Worksheet, tabs() As TTable, ByVal n As Long, _
                               ByVal keep As String, ByVal minBase As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    If n = 0 Then
        Set BuildDecisions = d
        Exit Function
    End If

    Dim i As Long, g As Long, key As String, okC As Long, allC As Long, a As Variant
    For i = 1 To n
        For g = 1 To tabs(i).GroupCount
            key = modConfig.Squash(tabs(i).GroupName(g))
            If Len(key) > 0 Then
                BaseTally ws, tabs(i), g, minBase, okC, allC
                If Not d.Exists(key) Then d(key) = Array(tabs(i).GroupName(g), 0&, 0&, False, False)
                a = d(key)
                a(1) = a(1) + okC
                a(2) = a(2) + allC
                d(key) = a
            End If
        Next g
    Next i

    ' 모아 놓은 집계로 최종 판정을 내린다.
    Dim k As Variant, inKeep As Boolean, passBase As Boolean
    For Each k In d.Keys
        a = d(k)
        inKeep = (Len(Trim$(keep)) = 0) Or modConfig.InList(CStr(a(0)), keep)
        passBase = (a(2) > 0) And (a(1) * 2 >= a(2))
        a(3) = inKeep
        a(4) = (inKeep And passBase)
        d(k) = a
    Next k

    Set BuildDecisions = d
End Function

' 판정 사전에서 이 배너군의 결과를 꺼낸다.
Public Function DecidedPass(d As Object, ByVal groupName As String) As Boolean
    Dim key As String
    If d Is Nothing Then Exit Function
    key = modConfig.Squash(groupName)
    If Not d.Exists(key) Then Exit Function
    DecidedPass = CBool(d(key)(4))
End Function

' 그 배너군에서 Base 하한을 넘는 항목이 몇 개인지 센다.
Public Sub BaseTally(ws As Worksheet, t As TTable, g As Long, minBase As Long, _
                     ByRef okCells As Long, ByRef allCells As Long)
    Dim c As Long
    okCells = 0: allCells = 0
    For c = t.GroupFirst(g) To t.GroupLast(g)
        allCells = allCells + 1
        If modTable.BaseOf(ws, t, c) >= minBase Then okCells = okCells + 1
    Next c
End Sub

Public Function IsShaded(ws As Worksheet, t As TTable, g As Long) As Boolean
    IsShaded = (ws.Cells(t.GroupRow, t.GroupFirst(g)).Interior.Color = SHADE_COLOR)
End Function

'==========================================================================
' [배너군 목록] 이 문항을 좁혀 뽑을 때도 이 두 곳을 그대로 쓰므로 Public 이다.
Public Sub Shade(ws As Worksheet, t As TTable, g As Long)
    ws.Range(ws.Cells(t.GroupRow, t.GroupFirst(g)), _
             ws.Cells(t.EndRow, t.GroupLast(g))).Interior.Color = SHADE_COLOR
End Sub

Public Sub ClearShade(ws As Worksheet, t As TTable, g As Long)
    ' 음영이 없는 배너군은 곧바로 지나간다.
    If Not IsShaded(ws, t, g) Then Exit Sub
    ' 범위를 통째로 한 번에 지운다. 칸마다 색을 물어 가며 지우면 표가 많을 때
    ' 그것만으로 몇 초가 사라진다. 회색으로 칠해 둔 배너군에는 애초에 인사이트를
    ' 칠하지 않으므로, 통째로 지워도 잃을 색이 없다.
    ws.Range(ws.Cells(t.GroupRow, t.GroupFirst(g)), _
             ws.Cells(t.EndRow, t.GroupLast(g))).Interior.ColorIndex = xlColorIndexNone
End Sub

'--------------------------------------------------------------------------
' 지우기 직전의 시트를 한 장 떠 둔다.
'--------------------------------------------------------------------------
Private Function MakeBackup(ws As Worksheet) As String
    Dim nm As String, i As Long
    On Error GoTo Fail
    nm = BACKUP_PREFIX & Format$(Now, "mmdd_hhnn")
    ' 같은 분에 두 번 눌렀을 때를 위해 뒤에 번호를 붙인다.
    Dim cand As String
    cand = nm
    i = 1
    Do While SheetExists(ws.Parent, cand)
        i = i + 1
        cand = nm & "_" & i
    Loop
    Application.ScreenUpdating = False
    ws.Copy After:=ws.Parent.Worksheets(ws.Parent.Worksheets.Count)
    ws.Parent.Worksheets(ws.Parent.Worksheets.Count).name = Left$(cand, 31)
    ws.Activate
    Application.ScreenUpdating = True
    MakeBackup = Left$(cand, 31)
    Exit Function
Fail:
    Application.ScreenUpdating = True
    MakeBackup = ""
End Function

Private Function SheetExists(wb As Workbook, ByVal nm As String) As Boolean
    Dim s As Object
    On Error Resume Next
    Set s = wb.Sheets(Left$(nm, 31))
    On Error GoTo 0
    SheetExists = Not s Is Nothing
End Function

'--------------------------------------------------------------------------
' 보존 목록에 적었지만 시트에 없는 이름을 찾아 준다. 대개 오타나 표기 차이다.
'--------------------------------------------------------------------------
Private Function UnmatchedKeepNames(ws As Worksheet, tabs() As TTable, n As Long, _
                                    keep As String) As String
    If Len(Trim$(keep)) = 0 Or n = 0 Then Exit Function

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim i As Long, g As Long
    For i = 1 To n
        For g = 1 To tabs(i).GroupCount
            seen(modConfig.Squash(tabs(i).GroupName(g))) = True
        Next g
    Next i

    Dim parts() As String, k As Long, miss As String
    parts = Split(Replace(Replace(Replace(keep, "·", ","), "|", ","), ";", ","), ",")
    For k = LBound(parts) To UBound(parts)
        If Len(modConfig.Squash(parts(k))) > 0 Then
            If Not seen.Exists(modConfig.Squash(parts(k))) Then
                If Len(miss) > 0 Then miss = miss & ", "
                miss = miss & Trim$(parts(k))
            End If
        End If
    Next k
    UnmatchedKeepNames = miss
End Function

' 빌드가 모든 모듈을 실제로 컴파일하도록 부르는 자리다. VBA 는 필요할 때만
' 컴파일하기 때문에, 한 번도 불리지 않은 모듈의 오류는 사람이 그 단추를 누르는
' 순간에야 드러난다. 이 함수가 그 구멍을 막는다.
Public Function Ping() As String
    Ping = "modBanner"
End Function
