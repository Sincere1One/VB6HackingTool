Attribute VB_Name = "Module_Display"
'Programmed by Chun Dong Hyuk

Option Explicit

Private Type SafeArrayBound
    cElements As Long
    lLbound As Long
End Type

Private Type SafeArray2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SafeArrayBound
End Type

Private Type BitmapInfoHeader
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type RGBQuad
    RGBBlue As Byte
    RGBGreen As Byte
    RGBRed As Byte
    RGBReserved As Byte
End Type

Private Type BitmapInfo
    bmiHeader As BitmapInfoHeader
    bmiColors() As RGBQuad
End Type

Private Type BitmapInfo16
    bmiHeader As BitmapInfoHeader
    bmiColors(0 To 15) As RGBQuad
End Type

Private Type BitmapInfo256
    bmiHeader As BitmapInfoHeader
    bmiColors(0 To 255) As RGBQuad
End Type

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal Hwnd As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal HDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal HDC As Long, pBitmapInfo As BitmapInfo, ByVal un As Long, lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateDIBSection16 Lib "gdi32" Alias "CreateDIBSection" (ByVal HDC As Long, pBitmapInfo As BitmapInfo16, ByVal un As Long, lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateDIBSection256 Lib "gdi32" Alias "CreateDIBSection" (ByVal HDC As Long, pBitmapInfo As BitmapInfo256, ByVal un As Long, lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function compress Lib "zlib.dll" (Dest As Any, destLen As Any, Src As Any, ByVal srcLen As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal Hwnd As Long, lpRect As Rect) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Type DisplayInfoType
    '데이터 크기
    DataBufferSize As Long
    '원래 데이터 크기
    OriginalBufferSize As Long
    '현재 스크린 가로 해상도
    ScreenWidth As Integer
    '현재 스크린 세로 해상도
    ScreenHeight As Integer
    '스크린 색상비트수
    ColorDepth As Integer
    'X좌표
    PositionX As Integer
    'Y좌표
    PositionY As Integer
End Type

'이미지 블록 가로 갯수
Private Const ImageBlockX As Integer = 5
'이미지 블록 세로 갯수
Private Const ImageBlockY As Integer = 5

'디스플레이정보를 저장할 구조체 변수
Private DisplayInfo As DisplayInfoType
'이미지 블록의 가로 사이즈
Private BlockWidth As Integer
'이미지 블록의 세로 사이즈
Private BlockHeight As Integer
'이미지 데이터 캐시 버퍼
Private CacheBuffer(ImageBlockX, ImageBlockY) As Long
'현재 디스플레이 데이터를 전송할 X 블록의 수
Private CurrentX As Integer
'현재 디스플레이 데이터를 전송할 Y 블록의 수
Private CurrentY As Integer
'비트맵 이미지 배열 구조체
Private SafeArray As SafeArray2D
'24비트 비트맵 정보 구조체
Private BitmapInfo As BitmapInfo
'4비트 비트맵 정보 구조체
Private BitmapInfo16 As BitmapInfo16
'8비트 비트맵 정보 구조체
Private BitmapInfo256 As BitmapInfo256
'임시 비트맵 오브젝트
Private TempBitmap As Long
'비트맵 데이터 주소
Private TempBitmapAddress As Long
'비트맵 DC
Private TempBitmapDC As Long
'비트맵 DIB
Private TempBitmapDIB As Long
'데스크탑 핸들
Private DesktopHandle As Long
'데스크탑 DC
Private DesktopDC As Long
'데스크탑 크기 구조체 변수
Private DesktopRect As Rect

'스크린의 핸들과 DC를 구한다.
Public Sub ScreenInitialize()
    '스크린의 핸들을 구한다.
    DesktopHandle = GetDesktopWindow()
    '스크린의 DC를 구한다.
    DesktopDC = GetDC(DesktopHandle)

End Sub

'디스플레이 정보를 전송한다.
Public Sub SendDisplayInfo(ByVal ColorDepth As Byte)
    On Error Resume Next
    Dim ByteBuffer() As Byte
    '핸들을 이용하여 스크린의 사이즈를 구한다.
    GetWindowRect DesktopHandle, DesktopRect
    '이전 정보와 해상도/색상비트수가 변경되었는지 검사한다.
    If (DisplayInfo.ScreenWidth <> DesktopRect.Right) Or (DisplayInfo.ScreenHeight <> DesktopRect.Bottom) Or (DisplayInfo.ColorDepth <> ColorDepth) Then
      '스크린 사이즈를 블럭사이즈로 분할한다.
      BlockWidth = DesktopRect.Right / ImageBlockX
      BlockHeight = DesktopRect.Bottom / ImageBlockY
      '현재 디스플레이 정보를 변수에 저장한다.
      DisplayInfo.ScreenWidth = DesktopRect.Right
      DisplayInfo.ScreenHeight = DesktopRect.Bottom
      DisplayInfo.ColorDepth = ColorDepth
      '스크린의 DIB를 생성한다.
      CreateDIB
    End If
    Do
         Do Until CurrentY > ImageBlockY
              Do Until CurrentX > ImageBlockX
                   '특정좌표의 스크린 데이터를 캡쳐한다.
                   DisplayInfo.PositionX = BlockWidth * CurrentX
                   DisplayInfo.PositionY = BlockHeight * CurrentY
                   '화면을 캡쳐한다.
                   ScreenCapture ByteBuffer()
                   If CacheBuffer(CurrentX, CurrentY) <> UBound(ByteBuffer()) Then
                     '디스플레이 데이터를 전송한다.
                     SendData ByteBuffer()
                     CacheBuffer(CurrentX, CurrentY) = UBound(ByteBuffer())
                     CurrentX = CurrentX + 1
                     Exit Sub
                   Else
                     CurrentX = CurrentX + 1
                   End If
                   DoEvents
              Loop
              CurrentX = 0
              CurrentY = CurrentY + 1
         Loop
         CurrentY = 0
    Loop
          
End Sub

'화면을 캡쳐한다.
Private Sub ScreenCapture(ByRef ByteBuffer() As Byte)
    '데이터크기[4]+스크린가로크기2]+스크린세로크기[2]+색상비트수[2]+X좌표[2]+Y좌표[2]+원래데이터크기[4]+실제데이터..
    On Error Resume Next
    Dim DIBBuffer() As Byte
    Dim TempBuffer() As Byte
    '특정 블럭의 이미지를 생성한 DC에 저장한다.
    BitBlt TempBitmapDC, 0, 0, BlockWidth, BlockHeight, DesktopDC, DisplayInfo.PositionX, DisplayInfo.PositionY, &HCC0020
    ReDim ByteBuffer(DisplayInfo.OriginalBufferSize - 1)
    '디스플레이 데이터를 바이트 배열에 복사한다.
    CopyMemory ByVal VarPtrArray(DIBBuffer()), VarPtr(SafeArray), 4
    CopyMemory ByteBuffer(0), DIBBuffer(0, 0), DisplayInfo.OriginalBufferSize
    CopyMemory ByVal VarPtrArray(DIBBuffer), 0&, 4
    '압축 임시배열의 크기를 구한다.
    DisplayInfo.DataBufferSize = DisplayInfo.OriginalBufferSize + (DisplayInfo.OriginalBufferSize * 0.01) + 12
    ReDim TempBuffer(DisplayInfo.DataBufferSize)
    '데이터를 압축한다.
    compress TempBuffer(0), DisplayInfo.DataBufferSize, ByteBuffer(0), DisplayInfo.OriginalBufferSize
    '디스플레이 정보를 설정한다.
    DisplayInfo.DataBufferSize = DisplayInfo.DataBufferSize + 18
    '바이트 배열 크기를 재설정한다.
    ReDim Preserve ByteBuffer(DisplayInfo.DataBufferSize - 1)
    '버퍼에 전송 디스플레이 정보 데이터를 복사한다.
    CopyMemory ByteBuffer(0), DisplayInfo, 18
    '압축한 디스플레이 데이터를 복사한다.
    CopyMemory ByteBuffer(18), TempBuffer(0), DisplayInfo.DataBufferSize - 18
    
End Sub

'DIB를 생성한다.
Private Sub CreateDIB()
    '이전 DC와 DIB가 존재한다면 제거한다.
    If TempBitmapDC <> 0 Then
      If TempBitmapDIB <> 0 Then
        SelectObject TempBitmapDC, TempBitmap
        DeleteObject TempBitmapDIB
      End If
      DeleteObject TempBitmapDC
    End If
    '비트맵 관련 변수들을 초기화한다.
    TempBitmapDC = 0
    TempBitmapDIB = 0
    TempBitmap = 0
    TempBitmapAddress = 0
    TempBitmapDC = CreateCompatibleDC(0)
    '이미지 캐시 버퍼를 클리어한다.
    CacheBufferClear
    '이상없이 DC가 생성되었다면..
    If TempBitmapDC <> 0 Then
      With BitmapInfo.bmiHeader
             .biSize = Len(BitmapInfo.bmiHeader)
             .biWidth = BlockWidth
             .biHeight = BlockHeight
             .biPlanes = 1
             .biBitCount = DisplayInfo.ColorDepth
             .biCompression = 0&
             .biSizeImage = BytesPerScanLine() * .biHeight
              DisplayInfo.OriginalBufferSize = .biSizeImage
      End With
      '스크린 색상비트수에 따라서 비트맵 정보 구조체를 구성한다.
      Select Case DisplayInfo.ColorDepth
                Case 4
                         '4Bit(16Color) 팔레트를 구성한다.
                         Create4BitPalette
                         'DIB를 생성한다.
                         TempBitmapDIB = CreateDIBSection16(TempBitmapDC, BitmapInfo16, 0, TempBitmapAddress, 0, 0)
                Case 8
                         '8Bit(256Color) 팔레트를 구성한다.
                         Create8BitPalette
                         'DIB를 생성한다.
                         TempBitmapDIB = CreateDIBSection256(TempBitmapDC, BitmapInfo256, 0, TempBitmapAddress, 0, 0)
                Case 24
                         'DIB를 생성한다.
                         TempBitmapDIB = CreateDIBSection(TempBitmapDC, BitmapInfo, 0, TempBitmapAddress, 0, 0)
      End Select
      '비트맵 이미지 배열 헤더를 구성한다.
      With SafeArray
            .cDims = 2
            .cbElements = 1
            .Bounds(0).cElements = BitmapInfo.bmiHeader.biHeight
            .Bounds(1).cElements = BytesPerScanLine()
            .pvData = TempBitmapAddress
      End With
      '이상없이 DIB가 생성되었다면 임시 비트맵 오브젝트를 생성한다.
      If TempBitmapDIB <> 0 Then
        TempBitmap = SelectObject(TempBitmapDC, TempBitmapDIB)
      Else
        DeleteObject TempBitmapDC
        TempBitmapDC = 0
      End If
    End If

End Sub

'라인당 바이트 사이즈를 구한다.
Private Function BytesPerScanLine() As Long
    BytesPerScanLine = BitmapInfo.bmiHeader.biWidth * DisplayInfo.ColorDepth
    If BytesPerScanLine Mod 32 > 0 Then
      BytesPerScanLine = BytesPerScanLine + 32 - (BytesPerScanLine Mod 32)
    End If
    BytesPerScanLine = BytesPerScanLine \ 8

End Function

'4Bit(16Color) 팔레트를 구성한다.
Private Sub Create4BitPalette()
    BitmapInfo16.bmiHeader = BitmapInfo.bmiHeader
    With BitmapInfo16.bmiColors(0)
           .RGBRed = &H0: .RGBGreen = &H0: .RGBBlue = &H0
    End With
    With BitmapInfo16.bmiColors(1)
           .RGBRed = &H80: .RGBGreen = 0: .RGBBlue = 0
    End With
    With BitmapInfo16.bmiColors(2)
           .RGBRed = 0: .RGBGreen = &H80: .RGBBlue = 0
    End With
    With BitmapInfo16.bmiColors(3)
           .RGBRed = &H80: .RGBGreen = &H80: .RGBBlue = 0
    End With
    With BitmapInfo16.bmiColors(4)
           .RGBRed = 0: .RGBGreen = 0: .RGBBlue = &H80
    End With
    With BitmapInfo16.bmiColors(5)
           .RGBRed = &H80: .RGBGreen = 0: .RGBBlue = &H80
    End With
    With BitmapInfo16.bmiColors(6)
           .RGBRed = 0: .RGBGreen = &H80: .RGBBlue = &H80
    End With
    With BitmapInfo16.bmiColors(7)
           .RGBRed = &HC0: .RGBGreen = &HC0: .RGBBlue = &HC0
    End With
    With BitmapInfo16.bmiColors(8)
           .RGBRed = &H80: .RGBGreen = &H80: .RGBBlue = &H80
    End With
    With BitmapInfo16.bmiColors(9)
           .RGBRed = &HFF: .RGBGreen = 0: .RGBBlue = 0
    End With
    With BitmapInfo16.bmiColors(10)
           .RGBRed = 0: .RGBGreen = &HFF: .RGBBlue = 0
    End With
    With BitmapInfo16.bmiColors(11)
           .RGBRed = &HFF: .RGBGreen = &HFF: .RGBBlue = 0
    End With
    With BitmapInfo16.bmiColors(12)
           .RGBRed = 0: .RGBGreen = 0: .RGBBlue = &HFF
    End With
    With BitmapInfo16.bmiColors(13)
           .RGBRed = &HFF: .RGBGreen = 0: .RGBBlue = &HFF
    End With
    With BitmapInfo16.bmiColors(14)
           .RGBRed = 0: .RGBGreen = &HFF: .RGBBlue = &HFF
    End With
    With BitmapInfo16.bmiColors(15)
           .RGBRed = &HFF: .RGBGreen = &HFF: .RGBBlue = &HFF
    End With

End Sub

'8Bit(256Color) 팔레트를 구성한다.
Private Sub Create8BitPalette()
    Dim ColorR  As Long, ColorG As Long, ColorB As Long
    Dim ResultR As Long, ResultG As Long, ResultB As Long
    Dim Count As Integer
    BitmapInfo256.bmiHeader = BitmapInfo.bmiHeader
    For ColorB = 0 To 256 Step 64
          ResultB = IIf(ColorB = 256, ColorB - 1, ColorB)
          For ColorG = 0 To 256 Step 64
                ResultG = IIf(ColorG = 256, ColorG - 1, ColorG)
                For ColorR = 0 To 256 Step 64
                      ResultR = IIf(ColorR = 256, ColorR - 1, ColorR)
                      With BitmapInfo256.bmiColors(Count)
                            .RGBRed = ResultR
                            .RGBGreen = ResultG
                            .RGBBlue = ResultB
                      End With
                      Count = Count + 1
                Next ColorR
          Next ColorG
    Next ColorB
   
End Sub

'이미지 캐시 버퍼를 클리어 한다.
Private Sub CacheBufferClear()
    Dim CountX As Integer
    Dim CountY As Integer
    '캐시 버퍼를 클리어한다.
    For CountY = 0 To ImageBlockY
          For CountX = 0 To ImageBlockX
               CacheBuffer(CountX, CountY) = 0
          Next
    Next

End Sub

'디스플레이 정보를 클리어 한다.
Public Sub DisplayInfoClear()
    With DisplayInfo
          .ScreenWidth = 0
          .ScreenHeight = 0
          .ColorDepth = 0
    End With
    
End Sub
