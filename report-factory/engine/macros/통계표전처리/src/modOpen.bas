Attribute VB_Name = "modOpen"
'==========================================================================
' modOpen — 기능 0. 통계표 열기
'
' 다른 폴더에 있는 통계표(.xls · .xlsx · .xlsm)를 골라 곧바로 작업을 시작한다.
' 고른 파일을 그대로 열지 않고 표가 담긴 시트를 새 통합 문서로 떠서 연다.
' 뒤이어 쓰는 [음영 배너 일괄삭제] 는 되돌릴 수 없기 때문에, 원본은 손대지
' 않은 채로 두고 사본에서만 작업하게 만든 것이다.
'==========================================================================
Option Explicit

' 어느 통계표에서 떠 온 사본인지 적어 두는 문서 속성의 이름.
' 모듈 수준 상수는 반드시 이 선언부에 두어야 한다. 프로시저 사이에 끼워 넣으면
' VBA 가 모듈 수준으로 보지 않아 「변수가 정의되지 않았습니다」로 멎는다.
Private Const SRC_PROP As String = "전처리_원본경로"

Public Sub OpenTableFile()
    Dim path As Variant
    path = Application.GetOpenFilename( _
        FileFilter:="통계표 파일 (*.xls; *.xlsx; *.xlsm), *.xls; *.xlsx; *.xlsm", _
        Title:="전처리할 통계표를 고르세요")
    If VarType(path) = vbBoolean Then Exit Sub          ' 취소
    OpenTablePath CStr(path)
End Sub

' 파일 고르는 일과 여는 일을 갈라 두었다. 대화상자 없이도 부를 수 있어야
' 자동 검증에서 이 경로를 실제로 확인할 수 있기 때문이다.
Public Sub OpenTablePath(ByVal path As String)
    Dim srcWb As Workbook, srcWs As Worksheet
    Application.ScreenUpdating = False
    On Error GoTo Fail
    Set srcWb = Application.Workbooks.Open(FileName:=CStr(path), ReadOnly:=True, UpdateLinks:=0)

    Set srcWs = FindTableSheet(srcWb)
    If srcWs Is Nothing Then
        srcWb.Close SaveChanges:=False
        Application.ScreenUpdating = True
        modConfig.Say "이 파일에서 통계표 시트를 찾지 못했습니다." & vbCrLf & _
               "A열에 'Base for %' 가 있는 시트가 있어야 합니다.", vbExclamation
        Exit Sub
    End If

    ' 표 시트를 새 통합 문서로 떠낸다. 표 제목을 찾아볼 수 있도록 '목 차' 시트가
    ' 있으면 함께 가져온다.
    Dim toc As Worksheet, s As Worksheet
    For Each s In srcWb.Worksheets
        If Replace(Trim$(s.name), " ", "") = "목차" Then Set toc = s
    Next s

    If toc Is Nothing Then
        srcWs.Copy
    Else
        srcWb.Worksheets(Array(toc.name, srcWs.name)).Copy
    End If
    Dim newWb As Workbook
    Set newWb = ActiveWorkbook
    On Error Resume Next
    newWb.Worksheets(srcWs.name).Activate
    On Error GoTo Fail

    RememberSource newWb, CStr(path)

    ' 원본은 아무것도 바꾸지 않고 닫는다.
    srcWb.Close SaveChanges:=False
    Application.ScreenUpdating = True

    Dim baseName As String
    baseName = Mid$(CStr(path), InStrRev(CStr(path), "\") + 1)
    If InStrRev(baseName, ".") > 0 Then baseName = Left$(baseName, InStrRev(baseName, ".") - 1)

    modConfig.Say "작업 사본을 열었습니다." & vbCrLf & vbCrLf & _
           "원본 : " & baseName & vbCrLf & _
           "시트 : " & ActiveSheet.Name & vbCrLf & vbCrLf & _
           "원본 파일은 바뀌지 않았습니다. 이 사본에서 1 → 2 → 3 순서로 실행하시고," & vbCrLf & _
           "끝나면 [내보내기] 를 눌러 xlsx 파일로 저장해 주세요.", vbInformation

    Exit Sub
Fail:
    Application.ScreenUpdating = True
    modConfig.Say "파일을 여는 중에 문제가 생겼습니다." & vbCrLf & Err.Description, vbCritical
End Sub

'==========================================================================
' 내보내기 — 전처리를 마친 통합 문서를 xlsx 로 저장한다.
'
' 데이터막대는 조건부서식, 배너 음영과 인사이트 형광은 셀 배경색이라 xlsx 에
' 그대로 담긴다. 저장 뒤에도 계속 손볼 수 있도록 저장한 파일을 열어 둔다.
'--------------------------------------------------------------------------
Public Sub ExportXlsx()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub

    ' 가져오기로 연 사본이면 원본 파일 이름을 그대로 쓰고 뒤에 _전처리후 를 붙인다.
    ' 그래야 어느 통계표를 처리한 결과인지 파일 이름만 보고 알 수 있다.
    Dim src As String, suggest As String
    src = SourcePathOf(wb)
    If Len(src) > 0 Then
        Dim nm As String, folder As String
        nm = Mid$(src, InStrRev(src, "\") + 1)
        folder = Left$(src, InStrRev(src, "\"))
        If InStrRev(nm, ".") > 0 Then nm = Left$(nm, InStrRev(nm, ".") - 1)
        suggest = folder & nm & "_전처리후.xlsx"
    Else
        suggest = "통계표 전처리.xlsx"
    End If

    Dim path As Variant
    path = Application.GetSaveAsFilename( _
        InitialFileName:=suggest, _
        FileFilter:="Excel 통합 문서 (*.xlsx), *.xlsx", _
        Title:="전처리한 통계표를 저장할 위치를 고르세요")
    If VarType(path) = vbBoolean Then Exit Sub          ' 취소

    On Error GoTo Fail
    Application.DisplayAlerts = False
    wb.SaveAs FileName:=CStr(path), FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True

    modConfig.Say "저장했습니다." & vbCrLf & vbCrLf & CStr(path), vbInformation
    Exit Sub
Fail:
    Application.DisplayAlerts = True
    modConfig.Say "저장하지 못했습니다." & vbCrLf & Err.Description, vbCritical
End Sub

'--------------------------------------------------------------------------
' 어느 통계표에서 떠 온 사본인지 문서 속성에 적어 둔다. 통합 문서에 붙어
' 다니므로 저장했다 다시 열어도 남는다.
'--------------------------------------------------------------------------
Private Sub RememberSource(wb As Workbook, ByVal path As String)
    On Error Resume Next
    wb.CustomDocumentProperties(SRC_PROP).Delete
    wb.CustomDocumentProperties.Add Name:=SRC_PROP, LinkToContent:=False, _
                                    Type:=msoPropertyTypeString, Value:=path
    On Error GoTo 0
End Sub

Private Function SourcePathOf(wb As Workbook) As String
    Dim s As String
    On Error Resume Next
    s = CStr(wb.CustomDocumentProperties(SRC_PROP).Value)
    On Error GoTo 0
    SourcePathOf = s
End Function

'--------------------------------------------------------------------------
' 표가 담긴 시트를 찾는다. 이름이 'Table' 인 시트를 먼저 보고, 없으면
' A열에 'Base for %' 가 있는 시트를 앞에서부터 찾는다.
'--------------------------------------------------------------------------
Private Function FindTableSheet(wb As Workbook) As Worksheet
    Dim ws As Worksheet, hit As Range

    For Each ws In wb.Worksheets
        If LCase$(Trim$(ws.Name)) = "table" Then
            Set FindTableSheet = ws
            Exit Function
        End If
    Next ws

    For Each ws In wb.Worksheets
        Set hit = Nothing
        On Error Resume Next
        Set hit = ws.Columns(1).Find(What:="Base for %", LookAt:=xlWhole, MatchCase:=False)
        On Error GoTo 0
        If Not hit Is Nothing Then
            Set FindTableSheet = ws
            Exit Function
        End If
    Next ws
End Function

' 빌드가 모든 모듈을 실제로 컴파일하도록 부르는 자리다. VBA 는 필요할 때만
' 컴파일하기 때문에, 한 번도 불리지 않은 모듈의 오류는 사람이 그 단추를 누르는
' 순간에야 드러난다. 이 함수가 그 구멍을 막는다.
Public Function Ping() As String
    Ping = "modOpen"
End Function
