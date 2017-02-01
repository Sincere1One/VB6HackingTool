VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  '없음
   Caption         =   "LH Tools"
   ClientHeight    =   3585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3720
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":581A
   ScaleHeight     =   3585
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3240
      Top             =   0
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
'(ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Private Sub Form_Load()
'URLDownloadToFile 0, "http://hosting.ohseon.com/gksthf1226/MSWINSCK.OCX", "C:\WINDOWS\system32\MSWINSCK.OCX", 0, 0
End Sub


Private Sub Form_Unload(Cancel As Integer)
Form1.con.Enabled = True
Form1.접속자목록.Enabled = True
Form1.Text1.Enabled = True
Form1.Text2.Enabled = True
Form1.con.Enabled = True
End Sub

Private Sub Timer1_Timer()
Form1.Show
Unload Me
End Sub
