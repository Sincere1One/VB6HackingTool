Attribute VB_Name = "Module_Process"
'Programmed by Chun Dong Hyuk

Option Explicit

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb&)
'프로그램을 초기화한다.
Public Sub Initialize()
    On Error GoTo ErrorProcess
    '중복실행일 경우 종료한다.
    If App.PrevInstance Then End
    Dim Temp As String
    With Form_Main
           '포트번호 세팅을 위해 소켓을 닫는다.
           .Winsock1.Close
           '컨트롤 소켓 포트번호 30308
           .Winsock1.LocalPort = 30308
           '접속을 대기한다.
           .Winsock1.Listen
    End With
    '스크린 핸들과 DC를 구한다.
    ScreenInitialize
    Exit Sub
ErrorProcess:
    '소켓 접속대기의 오류가 발생하면 프로그램을 종료한다.
    End
    
End Sub

'소켓으로 컨트롤 데이터를 전송한다.
Public Sub SendData(Buffer() As Byte)
    On Error Resume Next
    Form_Main.Winsock1.SendData Buffer()
    
End Sub

'소켓으로부터 수신한 데이터를 병합/처리한다.
Public Sub ReceiveBufferProcess(ByRef Buffer() As Byte, BufferSize As Long)
    Dim Address As Long
    Do Until Address = BufferSize
         Select Case Buffer(Address)
                    '디스플레이 정보 전송 초기화
                    Case 1
                             'NumLock/CapsLock/ScrollLock 키를 원격PC 상태와 동기화한다.
                             SetToggleKey Buffer(Address + 4), Buffer(Address + 5), Buffer(Address + 6)
                             '디스플레이 데이터를 전송한다.
                             SendDisplayInfo Buffer(Address + 1)
                             Address = Address + 7
                    '디스플레이 데이터 요청
                    Case 2
                             '디스플레이 데이터를 전송한다.
                             SendDisplayInfo Buffer(Address + 1)
                             Address = Address + 2
                    '한글/영문 모드 변환
                    Case 3
                             '한글/영문 변환코드가 들어오면 모드를 바꾼다.
                             ConvertCharactorMode Buffer(Address + 1)
                             Address = Address + 2
                    '키보드 정보
                    Case 4
                             '키보드 정보를 처리한다.
                             SetKeyBoardInfo Buffer(Address + 1), Buffer(Address + 2)
                             Address = Address + 3
                    '마우스 정보
                    Case 5
                             '마우스 정보를 처리한다.
                             CopyMemory MouseInfo, Buffer(Address + 1), 5
                             SetMouseInfo
                             Address = Address + 6
         End Select
    Loop
   
End Sub

'최상위 윈도우로 설정한다.
Public Sub SetTopWindow(Handle As Long)
    SetWindowPos Handle, -1, 0, 0, 10, 10, &H1

End Sub

'Control Move를 제어하기 위한 함수
Public Sub ControlDrag(ByVal Hwnd As Long)
    ReleaseCapture
    SendMessage Hwnd, &HA1, 2, 0&

End Sub

'Control Resize를 제어하기 위한 함수
Public Sub ControlResize(ByVal Hwnd As Long)
    ReleaseCapture
    SendMessage Hwnd, &HA1, 17, 0&

End Sub

'레지스트리에서 값을 읽어들인다.
Public Function GetRegistry(Root As Long, SubKey As String, ValueTitle As String) As String
    On Error Resume Next
    Dim Buffer1 As Long
    Dim Buffer2 As Long
    Dim Buffer3 As String * 255
    If RegOpenKeyEx(Root, SubKey, 0, 983103, Buffer2) <> 0 Then
      GetRegistry = ""
      RegCloseKey Buffer2
      Exit Function
    End If
    If RegQueryValueEx(Buffer2, ValueTitle, 0, Buffer1, ByVal Buffer3, 255) = 0 And Buffer1 = 1 Then
      GetRegistry = Mid$(Buffer3, 1, InStr(Buffer3, Chr$(0)) - 1)
    Else
      GetRegistry = ""
    End If
    RegCloseKey Buffer2
           
End Function

'레지스트리에 값을 저장한다.
Public Sub SetRegistry(Root As Long, SubKey As String, ValueTitle As String, Value As String)
    On Error Resume Next
    Dim Buffer As Long
    RegCreateKey Root, SubKey, Buffer
    RegSetValueEx Buffer, ValueTitle, 0, 1, ByVal Value, XLen(Value) + 1
    RegCloseKey Buffer
          
End Sub

'유니코드 문자열 Len 함수
Public Function XLen(Buffer As String) As Integer
    On Error Resume Next
    XLen = Val(LenB(StrConv(Buffer, vbFromUnicode)))

End Function

