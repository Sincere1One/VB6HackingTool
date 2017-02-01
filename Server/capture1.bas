Attribute VB_Name = "capture1"
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer
Const VK_F9 As Long = &H78
Dim NowStr() As String

'-----------------------------------------------------------------------
'Coded By 수학쟁이(startgoora)
'''''''''''''case "557cf402-1a04-11d3-9a73-0000f81ef32e": // GIF
'''''''''''''case "557cf403-1a04-11d3-9a73-0000f81ef32e": // EMF
'''''''''''''case "557cf400-1a04-11d3-9a73-0000f81ef32e": // BMP/DIB/RLE
'''''''''''''case "557cf401-1a04-11d3-9a73-0000f81ef32e": // JPG,JPEG,JPE,JFIF
'''''''''''''case "557cf406-1a04-11d3-9a73-0000f81ef32e": // PNG
'''''''''''''case "557cf407-1a04-11d3-9a73-0000f81ef32e": // ICO
'''''''''''''case "557cf404-1a04-11d3-9a73-0000f81ef32e": // WMF
'''''''''''''case "557cf405-1a04-11d3-9a73-0000f81ef32e": // TIF,TIFF
'-----------------------------------------------------------------------
'JPEG 로 저장 선언부 ---------------------------------------------------
Private Type Guid
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(0 To 7) As Byte
End Type

Private Type GdiplusStartupInput
   GdiplusVersion As Long
   DebugEventCallback As Long
   SuppressBackgroundThread As Long
   SuppressExternalCodecs As Long
End Type

Private Type EncoderParameter
   Guid As Guid
   NumberOfValues As Long
   Type As Long
   Value As Long
End Type

Private Type EncoderParameters
   Count As Long
   Parameter As EncoderParameter
End Type

Private Declare Function GdiplusStartup Lib "gdiplus" ( _
   Token As Long, _
   inputbuf As GdiplusStartupInput, _
   Optional ByVal outputbuf As Long = 0) As Long

Private Declare Function GdiplusShutdown Lib "gdiplus" ( _
   ByVal Token As Long) As Long

Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" ( _
   ByVal hbm As Long, _
   ByVal hPal As Long, _
   BITMAP As Long) As Long

Private Declare Function GdipDisposeImage Lib "gdiplus" ( _
   ByVal Image As Long) As Long

Private Declare Function GdipSaveImageToFile Lib "gdiplus" ( _
   ByVal Image As Long, _
   ByVal FileName As Long, _
   clsidEncoder As Guid, _
   encoderParams As Any) As Long

Private Declare Function CLSIDFromString Lib "ole32" ( _
   ByVal Str As Long, _
   ID As Guid) As Long
'-------------------------------------------------------


' ----==== SaveJPG ====----

Public Sub SaveJPG( _
   ByVal pict As StdPicture, _
   ByVal FileName As String, _
   Optional ByVal quality As Byte = 80)
Dim tSI As GdiplusStartupInput
Dim lRes As Long
Dim lGDIP As Long
Dim lBitmap As Long

   ' Initialize GDI+
   tSI.GdiplusVersion = 1
   lRes = GdiplusStartup(lGDIP, tSI)
   
   If lRes = 0 Then
   
      ' Create the GDI+ bitmap
      ' from the image handle
      lRes = GdipCreateBitmapFromHBITMAP(pict.Handle, 0, lBitmap)
   
      If lRes = 0 Then
         Dim tJpgEncoder As Guid
         Dim tParams As EncoderParameters
         
         ' Initialize the encoder GUID
         CLSIDFromString StrPtr("{557cf402-1a04-11d3-9a73-0000f81ef32e}"), _
                         tJpgEncoder
      
         ' Initialize the encoder parameters
         tParams.Count = 1
         With tParams.Parameter ' Quality
            ' Set the Quality GUID
            CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .Guid
            .NumberOfValues = 1
            .Type = 4
            .Value = VarPtr(quality)
         End With
         
         ' Save the image
         lRes = GdipSaveImageToFile( _
                  lBitmap, _
                  StrPtr(FileName), _
                  tJpgEncoder, _
                  tParams)
                             
         ' Destroy the bitmap
         GdipDisposeImage lBitmap
         
      End If
      
      ' Shutdown GDI+
      GdiplusShutdown lGDIP

   End If
   
   If lRes Then
   End If
   
End Sub




Public Sub SCapTure()
On Error Resume Next
    BitBlt Form_Main.Picture1.hdc, 0, 0, Screen.Width, Screen.Height, GetDC(0), 0, 0, vbSrcCopy
    NowStr = Split(Now, ":")
    SaveJPG Form_Main.Picture1.Image, "c:\WINDOWS\other.jpg", 80

End Sub
Public Sub scgo()
On Error Resume Next

Form_Main.Picture1.Width = Screen.Width
Form_Main.Picture1.Height = Screen.Height
SCapTure

End Sub





