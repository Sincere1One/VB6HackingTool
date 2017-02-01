VERSION 5.00
Begin VB.Form SC 
   Caption         =   "image"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4680
   Icon            =   "SC.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   3240
   ScaleWidth      =   4680
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton Command2 
      Caption         =   "읽기 중지"
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.Timer Read 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1200
      Top             =   480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "그림 읽기"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   2415
      Left            =   120
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "SC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Read.Enabled = True
End Sub

Private Sub Command2_Click()
Read.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Read.Enabled = True Then
Read.Enabled = False
End If
End Sub

Private Sub Read_Timer()
On Error Resume Next
Form1.Winsock1.SendData "/스크린"
Image1.Picture = LoadPicture(App.Path & "\win.jpg")
Kill App.Path & "\win.jpg"
End Sub
