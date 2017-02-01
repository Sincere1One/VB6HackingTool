Attribute VB_Name = "Module2"
Option Explicit

' -- ¹ÙÅÁÈ­¸é Ä¸ÃÄ
Const SRCCOPY = &HCC0020

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal Hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type

Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long
Private Declare Function GetCursor Lib "user32" () As Long
Private Declare Function DrawIcon Lib "user32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Public Type PointAPI
        X As Long
        Y As Long
End Type

Public Sub Cap()
    Dim HDC As Long
    Dim Handle As Long

    Dim MPos As PointAPI
    Dim SysCur As Long
    Dim SysIcon As ICONINFO

    With Form_Main.Picture1
        .Cls
        HDC = GetDC(GetDesktopWindow())
        .Width = Screen.Width
        .Height = Screen.Height

        BitBlt .HDC, 0, 0, Screen.Width, Screen.Height, HDC, 0, 0, SRCCOPY

        GetCursorPos MPos
        SysCur = GetCursor
        GetIconInfo SysCur, SysIcon
        DrawIcon .HDC, MPos.X - SysIcon.xHotspot, MPos.Y - SysIcon.yHotspot, SysCur

       SaveJPG .Image, "C:\WINDOWS\gogi.csr"

        .Cls
    End With
End Sub

