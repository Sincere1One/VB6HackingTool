VERSION 5.00
Begin VB.UserControl LunaButton 
   Appearance      =   0  '평면
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   1170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   855
   ForeColor       =   &H00000000&
   LockControls    =   -1  'True
   PaletteMode     =   4  '없음
   ScaleHeight     =   78
   ScaleMode       =   3  '픽셀
   ScaleWidth      =   57
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   390
      Top             =   60
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   0
      Left            =   60
      Top             =   60
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   1
      Left            =   60
      Top             =   330
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   2
      Left            =   60
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   3
      Left            =   60
      Top             =   870
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "LunaButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Programmed by Chun Dong Hyuk

Option Explicit

'Mouse Over 이벤트를 위한 API를 선언한다.
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal Hwnd As Long, lpRect As Rect) As Long

'API 에서 쓰일 구조체
Private Type PointAPI
    X As Long
    Y As Long
End Type

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'사용자 정의 컨트롤에서 사용할 이벤트를 정의한다.
Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseOver()
Public Event MouseExit()

'버튼 MouseOver 지원을 위한 플래그 변수
Private ButtonMouseOverFlag As Integer

'컨트롤의 핸들을 반환한다.
Public Property Get Hwnd() As Long
    Hwnd = UserControl.Hwnd

End Property

'컨트롤의 기본 이미지를 반환한다.
Public Property Get Picture() As Picture
    Set Picture = UserControl.Image1(0).Picture
    
End Property

'컨트롤의 기본 이미지를 설정한다.
Public Property Set Picture(ByVal NewPicture As Picture)
    Set UserControl.Image1(0).Picture = NewPicture
    PropertyChanged "Picture"
    UserControl.Picture = Image1(0).Picture
    
End Property

'컨트롤의 MouseOver 이미지를 반환한다.
Public Property Get MouseOverPicture() As Picture
    Set MouseOverPicture = UserControl.Image1(1).Picture

End Property

'컨트롤의 MouseOver 이미지를 설정한다.
Public Property Set MouseOverPicture(ByVal NewPicture As Picture)
    Set UserControl.Image1(1).Picture = NewPicture
    PropertyChanged "MouseOverPicture"
    
End Property

'컨트롤의 Disable 이미지를 반환한다.
Public Property Get DisablePicture() As Picture
    Set DisablePicture = UserControl.Image1(3).Picture

End Property

'컨트롤의 Disable 이미지를 설정한다.
Public Property Set DisablePicture(ByVal NewPicture As Picture)
    Set UserControl.Image1(3).Picture = NewPicture
    PropertyChanged "DisablePicture"
    
End Property

'컨트롤의 MouseDown 이미지를 반환한다.
Public Property Get MouseDownPicture() As Picture
    Set MouseDownPicture = UserControl.Image1(2).Picture

End Property

'컨트롤의 MouseDown 이미지를 설정한다.
Public Property Set MouseDownPicture(ByVal NewPicture As Picture)
    Set UserControl.Image1(2).Picture = NewPicture
    PropertyChanged "MouseDownPicture"
    
End Property

'컨트롤의 Enabled 속성을 반환한다.
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled

End Property

'컨트롤의 Enabled 속성을 설정한다.
Public Property Let Enabled(ByVal Flag As Boolean)
    UserControl.Enabled = Flag
    If Flag = True Then
      UserControl.Picture = Image1(0).Picture
    Else
      UserControl.Picture = Image1(3).Picture
    End If
    PropertyChanged "Enabled"
    
End Property

'MouseOver 이벤트 처리를 위해 타이머를 사용한다.
Private Sub Timer1_Timer()
    On Error Resume Next
    Dim WindowPointAPI As PointAPI
    Dim WindowRect As Rect
    GetCursorPos WindowPointAPI
    GetWindowRect UserControl.Hwnd, WindowRect
    If (WindowRect.Left <= WindowPointAPI.X And WindowRect.Right >= WindowPointAPI.X And WindowRect.Top <= WindowPointAPI.Y And WindowRect.Bottom >= WindowPointAPI.Y) Then
      If ButtonMouseOverFlag = 1 Then Exit Sub
      UserControl.Picture = Image1(1).Picture
      ButtonMouseOverFlag = 1
      RaiseEvent MouseOver
    Else
      UserControl.Picture = Image1(0).Picture
      ButtonMouseOverFlag = 0
      Timer1.Enabled = False
      RaiseEvent MouseExit
    End If

End Sub

'컨트롤 클릭시 클릭이벤트를 발생한다.
Private Sub UserControl_Click()
    RaiseEvent Click

End Sub

'초기 컨트롤 크기를 설정한다.
Private Sub UserControl_InitProperties()
    UserControl.Height = 255
    UserControl.Width = 255

End Sub

'컨트롤 MouseDown 이벤트를 처리한다.
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = False
    UserControl.Picture = Image1(2).Picture
    RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

'컨트롤 MouseMove 이벤트를 처리한다.
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ButtonMouseOverFlag = 0 Then
      Timer1.Enabled = True
    End If

End Sub

'컨트롤 MouseUp 이벤트를 처리한다.
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl.Picture = Image1(1).Picture
    Timer1.Enabled = True
    RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

'컨트롤 속성을 읽어서 설정한다.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    Set DisablePicture = PropBag.ReadProperty("DisablePicture", Nothing)
    Set MouseOverPicture = PropBag.ReadProperty("MouseOverPicture", Nothing)
    Set MouseDownPicture = PropBag.ReadProperty("MouseDownPicture", Nothing)
    Enabled = PropBag.ReadProperty("Enabled", True)
    
End Sub

'컨트롤 속성을 적용한다.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Picture", Picture, Nothing
    PropBag.WriteProperty "DisablePicture", DisablePicture, Nothing
    PropBag.WriteProperty "MouseOverPicture", MouseOverPicture, Nothing
    PropBag.WriteProperty "MouseDownPicture", MouseDownPicture, Nothing
    PropBag.WriteProperty "Enabled", Enabled, True
    
End Sub


