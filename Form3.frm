VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General Security System"
   ClientHeight    =   5640
   ClientLeft      =   1890
   ClientTop       =   645
   ClientWidth     =   6000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.TextBox Text6 
         Height          =   2415
         Left            =   4680
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   56
         Text            =   "Form3.frx":0ECA
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command43 
         Caption         =   "Disable"
         Height          =   255
         Left            =   3000
         TabIndex        =   55
         Top             =   3000
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command42 
         Caption         =   "Enable"
         Height          =   255
         Left            =   1560
         TabIndex        =   54
         Top             =   3000
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Height          =   2775
         Left            =   5400
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   53
         Text            =   "Form3.frx":0F5B
         Top             =   -480
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton Command41 
         Caption         =   "TNC Shape"
         Height          =   255
         Left            =   2880
         TabIndex        =   52
         ToolTipText     =   "Windows 3D Appearance Effect.."
         Top             =   4440
         Width           =   1215
      End
      Begin VB.CommandButton Command39 
         Caption         =   "E:\,F:\"
         Height          =   255
         Left            =   3720
         TabIndex        =   51
         Top             =   2160
         Width           =   615
      End
      Begin VB.CommandButton Command38 
         Caption         =   "D:\,E:\"
         Height          =   255
         Left            =   2880
         TabIndex        =   50
         Top             =   2160
         Width           =   735
      End
      Begin VB.CommandButton Command37 
         Caption         =   "A:\,C:\"
         Height          =   255
         Left            =   2040
         TabIndex        =   49
         Top             =   2160
         Width           =   735
      End
      Begin VB.CommandButton Command36 
         Caption         =   "K:\"
         Height          =   255
         Left            =   1440
         TabIndex        =   48
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton Command35 
         Caption         =   "J:\"
         Height          =   255
         Left            =   840
         TabIndex        =   47
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton Command34 
         Caption         =   "I:\"
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton Command33 
         Caption         =   "A:\"
         Height          =   255
         Left            =   1560
         TabIndex        =   45
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox Text4 
         Height          =   2535
         Left            =   5520
         TabIndex        =   43
         Text            =   """NoDrives""=hex:"
         Top             =   5160
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.CommandButton Command32 
         Caption         =   "Command32"
         Height          =   495
         Left            =   2880
         TabIndex        =   42
         Top             =   840
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton Command31 
         Caption         =   " Min\Max Spe"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         ToolTipText     =   "Speeding Up The Maximizing And Minimizing.."
         Top             =   4440
         Width           =   1215
      End
      Begin VB.CommandButton Command30 
         Caption         =   "Desktop Icon"
         Height          =   255
         Left            =   2880
         TabIndex        =   39
         ToolTipText     =   "To Hide Desktop Icon.."
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CommandButton Command29 
         Caption         =   "Back Setting"
         Height          =   255
         Left            =   4200
         TabIndex        =   38
         ToolTipText     =   "To Restrict Background Setting.."
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton Command28 
         Caption         =   "Disable"
         Height          =   255
         Left            =   3000
         TabIndex        =   37
         Top             =   3000
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command27 
         Caption         =   "Enable"
         Height          =   255
         Left            =   1560
         TabIndex        =   36
         Top             =   3000
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command26 
         Caption         =   "Control Panel"
         Height          =   255
         Left            =   4200
         TabIndex        =   35
         ToolTipText     =   "To Disable Control Panel"
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CommandButton Command25 
         Caption         =   "File Menu"
         Height          =   255
         Left            =   2880
         TabIndex        =   34
         ToolTipText     =   "To Disable File Menu"
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Log Off"
         Height          =   255
         Left            =   1560
         TabIndex        =   33
         ToolTipText     =   "To Disable Log off Menu"
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Run"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         ToolTipText     =   "To Disable Run Menu"
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton Command22 
         Caption         =   " Find"
         Height          =   255
         Left            =   4200
         TabIndex        =   31
         ToolTipText     =   "To Disable Find Menu"
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Favourite"
         Height          =   255
         Left            =   2880
         TabIndex        =   30
         ToolTipText     =   "To Disable Favourite Menu"
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Documents"
         Height          =   255
         Left            =   1560
         TabIndex        =   29
         ToolTipText     =   "To Disable Documents Menu"
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Disable"
         Height          =   255
         Left            =   2880
         TabIndex        =   28
         Top             =   3000
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Enable"
         Height          =   255
         Left            =   1560
         TabIndex        =   27
         Top             =   3000
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command17 
         Caption         =   "ShutDown"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         ToolTipText     =   "To Disable The ShutDown Command"
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   4920
         TabIndex        =   25
         Top             =   5640
         Width           =   1575
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Command16"
         Height          =   255
         Left            =   3840
         TabIndex        =   24
         Top             =   5640
         Width           =   1095
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Restart"
         Height          =   255
         Left            =   960
         TabIndex        =   23
         ToolTipText     =   "Restart The Computer To See The Effect."
         Top             =   4800
         Width           =   1815
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Clean Boot"
         Height          =   255
         Left            =   1560
         TabIndex        =   22
         ToolTipText     =   "Clean The Autoexec.Bat File To Quick Start."
         Top             =   4440
         Width           =   1215
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Close App"
         Height          =   255
         Left            =   1560
         TabIndex        =   20
         ToolTipText     =   "To Close Any Running Application."
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CommandButton Command12 
         Caption         =   "B  A  C  K"
         Height          =   255
         Left            =   4200
         TabIndex        =   19
         ToolTipText     =   "Back To Previous Screen."
         Top             =   4440
         Width           =   1215
      End
      Begin VB.CommandButton Command11 
         Caption         =   "E  X  I  T"
         Height          =   255
         Left            =   2880
         TabIndex        =   18
         ToolTipText     =   "Exit To Windows System"
         Top             =   4800
         Width           =   1815
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Speed Menu"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         ToolTipText     =   " To Change The Speed Of StartMenu..."
         Top             =   4080
         Width           =   1215
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         Height          =   780
         Left            =   360
         Picture         =   "Form3.frx":0FEA
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   16
         Top             =   360
         Width           =   780
      End
      Begin VB.CommandButton Command9 
         Caption         =   "G:\"
         Height          =   255
         Left            =   4560
         TabIndex        =   12
         ToolTipText     =   "To Hide G:\ Drive"
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton Command8 
         Caption         =   "D:\"
         Height          =   255
         Left            =   2760
         TabIndex        =   11
         ToolTipText     =   "To Hide D,E,F drive"
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton Command7 
         Caption         =   "C:\"
         Height          =   255
         Left            =   2160
         TabIndex        =   10
         ToolTipText     =   "To Hide C,E,F drive"
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton Command6 
         Caption         =   "H:\"
         Height          =   255
         Left            =   5160
         TabIndex        =   9
         ToolTipText     =   "To Hide A,E,F drive"
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton Command5 
         Caption         =   "F:\"
         Height          =   255
         Left            =   3960
         TabIndex        =   8
         ToolTipText     =   "To Hide F:\ Drive"
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton Command4 
         Caption         =   "E:\ "
         Height          =   255
         Left            =   3360
         TabIndex        =   7
         ToolTipText     =   "To Hide F:\ And E:\ Drive"
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Text            =   "Nodrives"
         Top             =   5640
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   5640
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   375
         Left            =   2280
         TabIndex        =   4
         Top             =   5640
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Show All Drive"
         Height          =   255
         Left            =   4440
         TabIndex        =   3
         ToolTipText     =   "To Show All Available Drive"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Hide  All Drive"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         ToolTipText     =   "To Hide All Drive"
         Top             =   1800
         Width           =   1215
      End
      Begin VB.OLE OLE1 
         Height          =   735
         Left            =   1680
         TabIndex        =   44
         Top             =   840
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   "HELP"
         Height          =   195
         Left            =   5040
         TabIndex        =   41
         Top             =   720
         Width           =   420
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Change Your System..."
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "CopyRight By  General Corporation 2000"
         Height          =   255
         Left            =   1560
         TabIndex        =   15
         Top             =   5280
         Width           =   3015
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Drive Security System For Windows"
         Height          =   255
         Left            =   1800
         TabIndex        =   14
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "General Security Information"
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
         Left            =   1800
         TabIndex        =   13
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Which Drive You Want To Hide......"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   1440
         Width           =   3255
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tahmina As String

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

 
 Const HKEY_CURRENT_USER = &H80000001
 
 


Public Sub SaveString(hKey As Long, StrPath As String, StrValue As String, StrData As String)
   Dim KeyH&
    r = RegCreateKey(hKey, StrPath, KeyH&)
    r = RegSetValueEx(KeyH&, StrValue, 0, 1, ByVal StrData, Len(StrData))
    r = RegCloseKey(KeyH&)
End Sub
Public Sub SaveString1(hKey As Long, StrPath As String, StrValue As String, StrData As String)
   Dim KeyH&
    r = RegCreateKey(hKey, StrPath, KeyH&)
    r = RegSetValueEx(KeyH&, StrValue, 0, 0, ByVal StrData, Len(StrData))
    r = RegCloseKey(KeyH&)
End Sub




Private Sub Command1_Click()
Text1.Text = "}"
Command3_Click
End Sub

Private Sub Command10_Click()
If Command1.Value = 10 Then
SaveSetting App.Title, App.Title, "RunWithSystem", 1
SaveString HKEY_CURRENT_USER, "Control Panle\Desktop", "MenuShowDelay", "MenuShowDelay"
Else
SaveSetting App.Title, App.Title, "RunWithSystem", 0
SaveString HKEY_CURRENT_USER, "Control Panel\Desktop", "MenuShowDelay", 0
End If

End Sub

Private Sub Command11_Click()
Dim a As Integer
a = MsgBox("Do You Really Want To Quit?", 49, "Quit")
If a = vbCancel Then
Load Me
Else
End
End If
End Sub

Private Sub Command12_Click()
Unload Me
FrmMain.Show
End Sub

Private Sub Command13_Click()
Unload Me
Form5.Show
End Sub

Private Sub Command14_Click()
On Error Resume Next
Kill ("c:\autoexec.bat")
End Sub

Private Sub Command15_Click()
MsgBox "Please Be Sure That All Other Application Is Closed.", 32, "APP Close.."
i = Shell("c:\windows\rundll.exe user.exe,exitwindowsexec", vbNormalFocus)
End Sub

Private Sub Command16_Click()
If Command1.Value = 6 Then
SaveSetting App.Title, App.Title, "RunWithSystem", 1
SaveString HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", Text3.Text, Text3.Text
Else
SaveSetting App.Title, App.Title, "RunWithSystem", 0
SaveString HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", Text3.Text, Text1.Text
End If
End Sub

Private Sub Command17_Click()
Text3.Text = "NoClose"
Command18.Visible = True
Command19.Visible = True
Command42.Visible = False
Command43.Visible = False

End Sub

Private Sub Command18_Click()
Text1.Text = "11111"
Command18.Visible = False
Command19.Visible = False
Command16_Click
End Sub

Private Sub Command19_Click()
Text1.Text = "2"
Command18.Visible = False
Command19.Visible = False
Command16_Click
End Sub

Private Sub Command2_Click()
Text1.Text = "11111"
Command3_Click
End Sub

Private Sub Command20_Click()
Text3.Text = "NoRecentDocsHistory"
Command18.Visible = True
Command19.Visible = True
Command42.Visible = False
Command43.Visible = False

End Sub

Private Sub Command21_Click()
Text3.Text = "NoFavoritesMenu"
Command18.Visible = True
Command19.Visible = True
Command42.Visible = False
Command43.Visible = False

End Sub

Private Sub Command22_Click()
Text3.Text = "NoFind"
Command18.Visible = True
Command19.Visible = True
Command42.Visible = False
Command43.Visible = False

End Sub

Private Sub Command23_Click()
Text3.Text = "NoRun"
Command18.Visible = True
Command19.Visible = True
Command42.Visible = False
Command43.Visible = False

End Sub

Private Sub Command24_Click()
Text3.Text = "NoLogOff"
Command18.Visible = True
Command19.Visible = True
Command42.Visible = False
Command43.Visible = False

End Sub

Private Sub Command25_Click()
Text3.Text = "NoFileMenu"
Command18.Visible = True
Command19.Visible = True
Command42.Visible = False
Command43.Visible = False

End Sub

Private Sub Command26_Click()
Text3.Text = "NoSetFolders"
Command18.Visible = True
Command19.Visible = True
Command42.Visible = False
Command43.Visible = False

End Sub

Private Sub Command27_Click()
If Command27.Value = 1 Then
SaveSetting App.Title, App.Title, "RunWithSystem", 1
SaveString HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\policies\system", "NoDispBackgroundPage", "NoDispBackgroundPage"
Command27.Visible = False
Command28.Visible = False
Else
SaveSetting App.Title, App.Title, "RunWithSystem", 0
SaveString HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\policies\system", "NoDispBackgroundPage", "11111"
Command27.Visible = False
Command28.Visible = False
End If
End Sub

Private Sub Command28_Click()
If Command28.Value = 1 Then
SaveSetting App.Title, App.Title, "RunWithSystem", 1
SaveString HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\policies\system", "NoDispBackgroundPage", "NoDispBackgroundPage"
Command28.Visible = False
Command27.Visible = False
Else
SaveSetting App.Title, App.Title, "RunWithSystem", 0
SaveString HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\policies\system", "NoDispBackgroundPage", "0"
Command27.Visible = False
Command28.Visible = False
End If
End Sub

Private Sub Command29_Click()
Command27.Visible = True
Command28.Visible = True
Command42.Visible = False
Command43.Visible = False

End Sub

Private Sub Command3_Click()
If Command1.Value = 1 Then
SaveSetting App.Title, App.Title, "RunWithSystem", 1
SaveString HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", Text2.Text, Text2.Text
Else
SaveSetting App.Title, App.Title, "RunWithSystem", 0
SaveString HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", Text2.Text, Text1.Text
End If

End Sub


Private Sub Command111_Click()
Dim a As Integer
a = MsgBox("Do You Really Want To Quit?", 49, "Quit")
If a = vbCancel Then
Load Me
Else
End
End If
End Sub

Private Sub Command30_Click()
Text3.Text = "NoDesktop"
Command18.Visible = True
Command19.Visible = True
Command42.Visible = False
Command43.Visible = False

End Sub

Private Sub Command31_Click()
If Command31.Value = 10 Then
SaveSetting App.Title, App.Title, "RunWithSystem", 1
SaveString HKEY_CURRENT_USER, "Control Panle\Desktop\WindowMetrics", "minanimate", "minanimate"
Else
SaveSetting App.Title, App.Title, "RunWithSystem", 0
SaveString HKEY_CURRENT_USER, "Control Panel\Desktop\WindowMetrics", "minanimate", 0
End If

End Sub

Private Sub Command32_Click()
On Error Resume Next

abc$ = "REGEDIT4" & Chr(13) & Chr(10) & "[HKEY_USERS\.Default\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer]" & Chr(13) & Chr(10)
dolby$ = abc & Text4.Text & tahmina
Kill (App.Path & "\Tahmina.reg")
FileNumber = FreeFile
FileName = App.Path & "\Tahmina.reg"
Open FileName For Append As #FileNumber
Print #FileNumber, dolby
Close #FileNumber


OLE1.Delete
OLE1.CreateLink App.Path & "\Tahmina.reg"
OLE1.DoVerb



End Sub

Private Sub Command33_Click()
tahmina$ = "01,00,00,00"
Command32_Click
End Sub

Private Sub Command34_Click()
tahmina$ = "00,01,00,00"
Command32_Click
End Sub

Private Sub Command35_Click()
tahmina$ = "00,02,00,00"
Command32_Click
End Sub

Private Sub Command36_Click()
tahmina$ = "00,04,00,00"
Command32_Click
End Sub

Private Sub Command37_Click()
tahmina$ = "05,00,00,00"
Command32_Click
End Sub

Private Sub Command38_Click()
tahmina$ = "18,00,00,00"
Command32_Click
End Sub

Private Sub Command39_Click()
tahmina$ = "30,00,00,00"
Command32_Click
End Sub

Private Sub Command4_Click()
tahmina$ = "10,00,00,00"
Command32_Click
End Sub

Private Sub Command40_Click()

End Sub

Private Sub Command41_Click()
Command43.Visible = True
Command42.Visible = True

End Sub

Private Sub Command42_Click()
On Error Resume Next

Kill (App.Path & "\Tahmina.reg")
FileNumber = FreeFile
FileName = App.Path & "\Tahmina.reg"
Open FileName For Append As #FileNumber
Print #FileNumber, Text5.Text
Close #FileNumber
Command42.Visible = False
Command43.Visible = False

OLE1.Delete
OLE1.CreateLink App.Path & "\Tahmina.reg"
OLE1.DoVerb

End Sub

Private Sub Command43_Click()
On Error Resume Next

Kill (App.Path & "\Tahmina.reg")
FileNumber = FreeFile
FileName = App.Path & "\Tahmina.reg"
Open FileName For Append As #FileNumber
Print #FileNumber, Text6.Text
Close #FileNumber
Command42.Visible = False
Command43.Visible = False

OLE1.Delete
OLE1.CreateLink App.Path & "\Tahmina.reg"
OLE1.DoVerb

End Sub

Private Sub Command5_Click()
tahmina$ = "20,00,00,00"
Command32_Click
End Sub

Private Sub Command6_Click()
tahmina$ = "80,00,00,00"
Command32_Click
End Sub

Private Sub Command7_Click()
tahmina$ = "04,00,00,00"
Command32_Click
End Sub

Private Sub Command8_Click()
tahmina$ = "08,00,00,00"
Command32_Click
End Sub

Private Sub Command9_Click()
tahmina$ = "40,00,00,00"
Command32_Click
End Sub

Private Sub Label6_Click()
MsgBox "Don't Change Anything, If You Are Not Confirm On About It.", 32, "Info"
End Sub
