Attribute VB_Name = "modInstall"
'==========================================================================
' modInstall — 더블클릭 한 번으로 설치 (1천 명 배포 기준)
'
' 이 파일을 직접 열면(더블클릭) 스스로 추가 기능 폴더로 복사하고 등록한 뒤
' 닫힌다. ThisWorkbook.Workbook_Open 이 이 모듈의 OfferInstall 을 부른다.
'
' 설치되는 이름은 이 파일 자신의 이름과 무관하게 늘 고정된 AI_NAME(통계표
' 전처리.xlam)이다. 나눠 주는 파일은 일부러 다른 이름(통계표전처리_설치.xlam)
' 을 쓴다 - 엑셀은 이름이 같은 문서를 두 개 동시에 열지 못하는데, 이미 설치
' 되어 자동으로 불려온 옛 버전과 새로 받은 파일의 이름이 같으면 그 순간
' 「같은 이름의 통합 문서를 열 수 없습니다」로 막혀 이 코드가 실행될 기회조차
' 없다. 배포 파일 이름을 다르게 두면 이 문제 자체가 생기지 않는다 - 업데이트를
' 나눠 줄 때마다 예전 버전을 쓰던 사람이 걸리는 일이 없어야 하기 때문이다.
'
' 자동화(빌드 검증 등)로 조용히 열릴 때는 아무 일도 하지 않는다 - 그 판단은
' ThisWorkbook 쪽에서 Application.Visible 로 먼저 가른다. 이미 추가 기능
' 폴더에서 열린 것(정상적으로 등록되어 자동으로 불려온 경우)도 여기까지
' 오지 않는다. AddIns.Add·Installed 는 엑셀 창이 실제로 보이는 상태에서만
' 동작한다 - 이 함수가 자동화로는 불리지 않도록 가른 이유이기도 하다.
'==========================================================================
Option Explicit

Public Const AI_NAME As String = "통계표전처리.xlam"

Public Sub OfferInstall(wb As Workbook)
    On Error GoTo Fail

    If Val(Application.Version) < 14 Then
        modConfig.Say "이 추가 기능은 엑셀 2010(버전 14) 이상에서 동작합니다." & vbCrLf & vbCrLf & _
                      "지금 쓰시는 엑셀은 버전 " & Application.Version & " 입니다. 이 파일은" & vbCrLf & _
                      "그냥 닫으셔도 됩니다.", vbExclamation
        Exit Sub
    End If

    Dim addinsDir As String, dst As String
    addinsDir = Environ$("APPDATA") & "\Microsoft\AddIns\"
    dst = addinsDir & AI_NAME

    ' 등록된 이름으로 예전 버전이 있는지 본다. 경로가 아니라 이름으로 찾는다 -
    ' 레지스트리에 남은 경로가 어긋나 있어도(예전에 다른 자리에 설치됐던 경우)
    ' 이름은 늘 같기 때문이다.
    Dim a As AddIn, oldAi As AddIn, found As Boolean
    For Each a In Application.AddIns
        If StrComp(a.Name, AI_NAME, vbTextCompare) = 0 Then
            found = True
            Set oldAi = a
            Exit For
        End If
    Next a

    Dim msg As String
    If found Then
        msg = "통계표 전처리는 이미 설치되어 있습니다." & vbCrLf & vbCrLf & _
              "지금 연 파일로 덮어써서 새 버전으로 올릴까요?"
    Else
        msg = "통계표 전처리를 추가 기능으로 설치할까요?" & vbCrLf & vbCrLf & _
              "설치하면 앞으로 어떤 통계표 파일을 열어도 리본에" & vbCrLf & _
              "[통계표 전처리] 탭이 뜹니다. 이 창은 설치 뒤 저절로 닫힙니다."
    End If

    If Not modConfig.Ask(msg, "통계표 전처리 설치") Then
        modConfig.Say "설치하지 않았습니다. 이 파일을 열어 둔 동안만 쓸 수 있고," & vbCrLf & _
                      "다른 통계표 파일에는 리본 탭이 뜨지 않습니다."
        Exit Sub
    End If

    If Dir$(addinsDir, vbDirectory) = "" Then MkDir addinsDir

    ' 이미 등록되어 켜져 있는 예전 버전이 있으면 먼저 꺼서 파일 잠금을 푼다.
    ' 등록은 있는데 파일이 이미 없어졌거나 하는 어긋난 상태라도(On Error Resume
    ' Next 로 감싸) 여기서 멈추지 않고 새로 까는 쪽으로 계속 진행한다.
    If found Then
        On Error Resume Next
        If oldAi.Installed Then oldAi.Installed = False
        On Error GoTo Fail
    End If

    ' 방금 끈 파일의 운영체제 쪽 잠금이 곧바로 풀리지 않을 때가 있어, 짧게
    ' 나눠 몇 번 다시 시도한다. 그래도 안 풀리면 사람에게 원인을 짚어 준다.
    Dim tries As Long, copyErr As Long
    Do
        On Error Resume Next
        Err.Clear
        wb.SaveCopyAs dst
        copyErr = Err.Number
        On Error GoTo Fail
        If copyErr = 0 Then Exit Do
        tries = tries + 1
        If tries >= 5 Then
            Err.Raise copyErr, , "이 파일이 계속 사용 중이라 복사하지 못했습니다."
        End If
        Application.Wait Now + TimeValue("0:00:01")
    Loop

    ' 이미 등록된 이름이면 그 AddIn 개체를 그대로 다시 쓴다. 엑셀이 이미 아는
    ' 경로를 Add() 로 또 추가하려 하면 「Add 속성을 구할 수 없습니다」로 막힌다
    ' - 방금 파일만 새로 덮어썼을 뿐 등록 자체는 그대로이므로 다시 더할 이유가
    ' 없다.
    Dim newAi As AddIn
    If found Then
        Set newAi = oldAi
    Else
        Set newAi = Application.AddIns.Add(dst)
    End If

    ' Installed 를 켜는 시점이 파일을 막 바꾼 직후라 아주 드물게 한 번에
    ' 안 먹을 때가 있어 몇 번 다시 시도한다.
    Dim installTries As Long, instErr As Long
    Do
        On Error Resume Next
        Err.Clear
        newAi.Installed = True
        instErr = Err.Number
        On Error GoTo Fail
        If instErr = 0 Then Exit Do
        installTries = installTries + 1
        If installTries >= 3 Then Err.Raise instErr, , "추가 기능 등록을 켜지 못했습니다."
        Application.Wait Now + TimeValue("0:00:01")
    Loop

    modConfig.Say "설치를 마쳤습니다." & vbCrLf & vbCrLf & _
                  "엑셀을 새로 열면 [통계표 전처리] 탭이 보입니다." & vbCrLf & _
                  "이 파일은 이제 닫으셔도 됩니다 - 다시 열 필요는 없습니다."

    wb.Close False
    Exit Sub

Fail:
    modConfig.Say "설치하다가 멈췄습니다." & vbCrLf & vbCrLf & _
                  "오류 " & Err.Number & " : " & Err.Description & vbCrLf & vbCrLf & _
                  "다른 통계표 파일과 엑셀 창을 모두 닫은 뒤 이 파일만 다시 열어 보세요.", vbCritical
End Sub

' 빌드가 모든 모듈을 실제로 컴파일하도록 부르는 자리다. VBA 는 필요할 때만
' 컴파일하기 때문에, 한 번도 불리지 않은 모듈의 오류는 사람이 그 단추를 누르는
' 순간에야 드러난다. 이 함수가 그 구멍을 막는다.
Public Function Ping() As String
    Ping = "modInstall"
End Function
