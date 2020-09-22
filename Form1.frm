VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About WinMaster 4.0"
   ClientHeight    =   3825
   ClientLeft      =   2520
   ClientTop       =   1725
   ClientWidth     =   4830
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   0
      Picture         =   "Form1.frx":0ECA
      ScaleHeight     =   3495
      ScaleWidth      =   5250
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   5250
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   1920
      Picture         =   "Form1.frx":626B
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   1
      Top             =   1920
      Width           =   720
   End
   Begin VB.Timer Timer1 
      Interval        =   2
      Left            =   240
      Top             =   1200
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5400
      Left            =   120
      Picture         =   "Form1.frx":7135
      ScaleHeight     =   5400
      ScaleWidth      =   4500
      TabIndex        =   0
      Top             =   2400
      Width           =   4500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "        General Corporation               Number One In The Software                Technology"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   960
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_DblClick()
Picture3.Visible = True
End Sub

Private Sub Picture3_Click()
Picture3.Visible = False
End Sub

Private Sub Timer1_Timer()
If Picture1.Top = -5500 Then
Label1.Visible = True
Else

Picture1.Top = Picture1.Top - 5
Picture2.Top = Picture2.Top - 5
End If
End Sub
