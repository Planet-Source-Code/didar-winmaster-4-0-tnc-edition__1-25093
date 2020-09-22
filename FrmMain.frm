VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WinMaster4.0"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin VB.CommandButton Command8 
         Caption         =   "Close App"
         Height          =   375
         Left            =   3240
         TabIndex        =   17
         ToolTipText     =   "To Close Current Running Application."
         Top             =   2880
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Add"
         Height          =   375
         Left            =   5760
         TabIndex        =   16
         ToolTipText     =   "To Set The Selected *.exe Program For Startup"
         Top             =   3600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command6 
         Caption         =   "System"
         Height          =   375
         Left            =   2280
         TabIndex        =   14
         ToolTipText     =   "Drive And System Property Window"
         Top             =   2880
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Remove"
         Height          =   375
         Left            =   1320
         TabIndex        =   12
         ToolTipText     =   "To Remove The Application From Startup Menu.."
         Top             =   2880
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Info"
         Height          =   375
         Left            =   4200
         TabIndex        =   11
         ToolTipText     =   "Information."
         Top             =   2880
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   480
         Picture         =   "FrmMain.frx":0ECA
         ScaleHeight     =   540
         ScaleWidth      =   540
         TabIndex        =   10
         Top             =   360
         Width           =   540
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   360
         TabIndex        =   5
         Top             =   1560
         Width           =   5655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Set"
         Height          =   375
         Left            =   5760
         TabIndex        =   4
         ToolTipText     =   "To Set The Selected *.exe File for Startup."
         Top             =   3720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   360
         TabIndex        =   3
         Text            =   "0"
         Top             =   2280
         Width           =   5655
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Add"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         ToolTipText     =   "Select Any *.exe File To Load At StartUp"
         Top             =   2880
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Exit"
         Height          =   375
         Left            =   5160
         TabIndex        =   1
         ToolTipText     =   "Exit"
         Top             =   2880
         Width           =   855
      End
      Begin MSComDlg.CommonDialog cmdlg 
         Left            =   360
         Top             =   4320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Help"
         Height          =   255
         Left            =   5520
         TabIndex        =   15
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "CopyRight By General Corporation 2000"
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Program Id"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Parameter  (*.exe)"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "GSI Windows System Administration"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label4 
         Caption         =   "A Powerul System Program         Written By Didar"
         Height          =   495
         Left            =   2160
         TabIndex        =   6
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Menu MnuMain 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu MnuMainShow 
         Caption         =   "&Show"
      End
      Begin VB.Menu MnuMainHide 
         Caption         =   "&Hide"
      End
      Begin VB.Menu MnuMainS1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMainNext 
         Caption         =   "&Next"
      End
      Begin VB.Menu MnuMainBack 
         Caption         =   "&Back"
      End
      Begin VB.Menu MnuMainS2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMainClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

 
 Const HKEY_LOCAL_MACHINE = &H80000002
 
 


Public Sub SaveString(hKey As Long, StrPath As String, StrValue As String, StrData As String)
   Dim KeyH&
    r = RegCreateKey(hKey, StrPath, KeyH&)
    r = RegSetValueEx(KeyH&, StrValue, 0, 1, ByVal StrData, Len(StrData))
    r = RegCloseKey(KeyH&)
End Sub




Private Sub Command1_Click()
If Command1.Value = 1 Then
SaveSetting App.Title, App.Title, "RunWithSystem", 1
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "General System" & Text2.Text, "General System" & Text2.Text
Else
SaveSetting App.Title, App.Title, "RunWithSystem", 0
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "General System " & Text2.Text, Text1.Text
End If

End Sub


Private Sub Command2_Click()
cmdlg.Filter = "*.exe|*.exe"
cmdlg.ShowOpen
Text1.Text = cmdlg.FileName
Command7_Click
End Sub

Private Sub Command3_Click()
Dim a As Integer
a = MsgBox("Do You Really Want To Quit?", 49, "Quit")
If a = vbCancel Then
Load Me
Else
End
End If
End Sub

Private Sub Command4_Click()
Form1.Show
End Sub

Private Sub Command5_Click()
Load Form2
Form2.Show
End Sub

Private Sub Command6_Click()
Unload Me
Form3.Show
End Sub



Private Sub Command7_Click()
If Text1.Text = "" Then
MsgBox "No File Selected To Load At Startup", 16, "No File"
Text1.SetFocus
Else
If Text2.Text = "" Then
MsgBox "Program Id Is Empty. You Must Enter Any Numeric Number(0,1,2,3...)", 16, "Invalid"
Text2.SetFocus
Else
Command1_Click
End If
End If
End Sub

Private Sub Command8_Click()
Unload Me
Form5.Show
End Sub

Private Sub Label6_Click()
MsgBox "Select Any Exe File To Run At StartUp.This GSI System Administration Tools Is A PowerFul System Utility Tool And It's Designed For Only Expert User.Proper Use Of This Software Can Change Your System Speed By The Function Of Windows DNA Structure.                                                                                                                                                                                                  General Corporation <<<>>>Always Tries To Do Something For You.", 32, "Help"
End Sub
