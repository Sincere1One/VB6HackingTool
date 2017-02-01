VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Linears Hacking Tool 1.0"
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4680
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   3210
   ScaleWidth      =   4680
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton Command2 
      Caption         =   "kill"
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '가운데 맞춤
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Text            =   "프로세서 이름"
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "프로세서 보기"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   1575
   End
   Begin VB.ListBox proList 
      Height          =   2580
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Winsock1.SendData "/pr좀"
End Sub

Private Sub Command2_Click()
Form1.Winsock1.SendData "/킬시켜:" & Text1.Text
End Sub
