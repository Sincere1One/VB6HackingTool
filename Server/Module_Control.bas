Attribute VB_Name = "Module_Control"
'Programmed by Chun Dong Hyuk

Private Const MOUSEEVENTF_ABSOLUTE = &H8000
Private Const MOUSEEVENTF_MOVE = &H1
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const MOUSEEVENTF_RIGHTDOWN = &H8
Private Const MOUSEEVENTF_RIGHTUP = &H10
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20
Private Const MOUSEEVENTF_MIDDLEUP = &H40
Private Const MOUSEEVENTF_WHEEL = &H800
Private Const MOUSEEVENTF_XDOWN = &H100
Private Const MOUSEEVENTF_XUP = &H200
Private Const MOUSEEVENTF_WHEELDELTA = 120

Private Type PointAPI
    X As Long
    Y As Long
End Type

Private Type MouseInfoType
    X As Integer
    Y As Integer
    Button As Byte
End Type

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetMessageExtraInfo Lib "user32" () As Long
Private Declare Function SetCursorPos Lib "user32.dll" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ImmGetContext Lib "imm32.dll" (ByVal Hwnd As Long) As Long
Private Declare Function ImmSetConversionStatus Lib "imm32.dll" (ByVal hIME As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
Private Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAuction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Sub mouse_event Lib "user32.dll" (ByVal dwFlags As Long, ByVal dX As Long, ByVal dY As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb&)

'마지막 마우스 버튼의 상태를 저장하는 변수
Private MouseButtonInfo As Byte
'기존 바탕화면 이미지의 경로(Wallpaper/OriginalWallpaper)
Private DesktopImageInfo(1) As String
'연결종료시 바탕화면 복구여부를 저장하는 변수
Private RestoreDesktopImage As Byte

'마우스 좌표 및 버튼 정보를 저장하는 변수
Public MouseInfo As MouseInfoType

'토글키를 원격PC와 동기화한다.
Public Sub SetToggleKey(NumLock As Byte, CapsLock As Byte, ScrollLock As Byte)
    'NumLock키를 설정한다.
    If NumLock = 1 Then
      If GetKeyState(&H90) = 0 Then
        keybd_event &H90, 0, 1, GetMessageExtraInfo
        keybd_event &H90, 0, 3, GetMessageExtraInfo
      End If
    Else
      If GetKeyState(&H90) = 1 Then
        keybd_event &H90, 0, 1, GetMessageExtraInfo
        keybd_event &H90, 0, 3, GetMessageExtraInfo
      End If
    End If
    'CapsLock키를 설정한다.
    If CapsLock = 1 Then
      If GetKeyState(&H14) = 0 Then
        keybd_event &H14, 0, 1, GetMessageExtraInfo
        keybd_event &H14, 0, 3, GetMessageExtraInfo
      End If
    Else
      If GetKeyState(&H14) = 1 Then
        keybd_event &H14, 0, 1, GetMessageExtraInfo
        keybd_event &H14, 0, 3, GetMessageExtraInfo
      End If
    End If
    'ScrollLock키를 설정한다.
    If ScrollLock = 1 Then
      If GetKeyState(&H91) = 0 Then
        keybd_event &H91, 0, 1, GetMessageExtraInfo
        keybd_event &H91, 0, 3, GetMessageExtraInfo
      End If
    Else
      If GetKeyState(&H91) = 1 Then
        keybd_event &H91, 0, 1, GetMessageExtraInfo
        keybd_event &H91, 0, 3, GetMessageExtraInfo
      End If
    End If

End Sub

'키보드 정보를 처리한다.
Public Sub SetKeyBoardInfo(ButtonFlag As Byte, KeyCode As Byte)
    Select Case ButtonFlag
              'KeyDown
              Case 1
                       keybd_event KeyCode, 0, 1, GetMessageExtraInfo
              'KeyUp
              Case 2
                       keybd_event KeyCode, 0, 3, GetMessageExtraInfo
    End Select

End Sub

'마우스 정보를 처리한다.
Public Sub SetMouseInfo()
    '원하는 좌표에 마우스 커서를 이동한다.
    If (MouseInfo.Button <> 8) And (MouseInfo.Button <> 9) Then
       SetCursorPos MouseInfo.X, MouseInfo.Y
    End If
    Select Case MouseInfo.Button
            Case 0   'None
                     '기존에 왼쪽버튼이 눌려있었다면, 눌림을 해제한다.
                     If MouseButtonInfo = 1 Then
                       mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, GetMessageExtraInfo
                     '기존에 오른쪽버튼이 눌려있었다면, 눌림을 해제한다.
                     ElseIf MouseButtonInfo = 2 Then
                       mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, GetMessageExtraInfo
                     End If
                     MouseButtonInfo = 0
            Case 1   'Left
                     '기존에 왼쪽버튼이 눌려있지 않았다면, 버튼을 눌림상태로 설정한다.
                     If MouseButtonInfo <> 1 Then
                       mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, GetMessageExtraInfo
                     End If
                     MouseButtonInfo = 1
            Case 2   'Right
                     '기존에 오른쪽버튼이 눌려있지 않았다면, 버튼을 눌림상태로 설정한다.
                     If MouseButtonInfo <> 2 Then
                       mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, GetMessageExtraInfo
                     End If
                     MouseButtonInfo = 2
            Case 3   'Both
                     MouseButtonInfo = 3
            Case 8   'Wheel
                     mouse_event MOUSEEVENTF_WHEEL, 0, 0, MOUSEEVENTF_WHEELDELTA, GetMessageExtraInfo
            Case 9   'Wheel
                     mouse_event MOUSEEVENTF_WHEEL, 0, 0, -MOUSEEVENTF_WHEELDELTA, GetMessageExtraInfo
    End Select

End Sub

'키보드 한글/영문으로 변환한다.
Public Sub ConvertCharactorMode(Flag As Byte)
    Dim PointType As PointAPI
    Dim Handle As Long
    Dim hIME As Long
    '현재 마우스 커서위치를 구한다.
    GetCursorPos PointType
    '마우스 커서가 위치한 곳의 핸들을 구한다.
    Handle = WindowFromPointXY(PointType.X, PointType.Y)
    '핸들을 이용하여 IME를 변환한다.
    hIME = ImmGetContext(Handle)
    '한글상태로 설정한다.
    If Flag = 1 Then
      ImmSetConversionStatus hIME, &H1, &H0
    '영문상태로 설정한다.
    ElseIf Flag = 2 Then
      ImmSetConversionStatus hIME, &H0, &H0
    End If

End Sub



