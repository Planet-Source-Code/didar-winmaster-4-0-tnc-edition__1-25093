VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Running Application's"
   ClientHeight    =   3990
   ClientLeft      =   1665
   ClientTop       =   1755
   ClientWidth     =   6630
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   6375
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   3000
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.ListBox List1 
         Height          =   2205
         ItemData        =   "Form5.frx":0ECA
         Left            =   120
         List            =   "Form5.frx":0ECC
         TabIndex        =   2
         Top             =   480
         Width           =   3975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         Height          =   375
         Left            =   5160
         TabIndex        =   1
         ToolTipText     =   "To Terminate Current Select Application."
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "This Part Is Only For System  Administrator..."
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   3600
         Width           =   3135
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   B   A  C  K"
         Height          =   255
         Left            =   5040
         TabIndex        =   6
         ToolTipText     =   "Back  Window"
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "1.Kernel32.dll   2. MsgSrv32.exe 3.Mprexe.exe      4.Mmtask.tsk     5.Explorer.exe 6.Wsloader.exe  7.Systray.exe"
         Height          =   1575
         Left            =   5040
         TabIndex        =   5
         ToolTipText     =   "Warning!!"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "    Remember   Don't Close This.."
         Height          =   375
         Left            =   4920
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Current Running Application.."
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   180
         Width           =   3615
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const MAX_PATH& = 260
Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * MAX_PATH
    End Type

Private Declare Function RegisterServiceProcess Lib "kernel32" (ByVal ProcessID As Long, ByVal ServiceFlags As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long


Private Sub Command1_Click()
If Text1.Text = "c:\windows\desktop\general.exe" Then
MsgBox "Windows Is Unable To Close This Application.", 16, "Error!!"
Else
If Text1.Text = "c:\windows\system\kernel32.dll" Then
MsgBox "Windows Is Unable To Close This Application.", 16, "Error!!"
Else
If Text1.Text = "c:\windows\system\msgsrv32.exe" Then
MsgBox "Windows Is Unable To Close This Application.", 16, "Error!!"
Else
If Text1.Text = "c:\windows\system\mprexe.exe" Then
MsgBox "Windows Is Unable To Close This Application.", 16, "Error!!"
Else
If Text1.Text = "c:\windows\system\mmtask.tsk" Then
MsgBox "Windows Is Unable To Close This Application.", 16, "Error!!"
Else
If Text1.Text = "c:\windows\explorer.exe" Then
MsgBox "Windows Is Unable To Close This Application.", 16, "Error!!"
Else
If Text1.Text = "c:\windows\system\wsloader.exe" Then
MsgBox "Windows Is Unable To Close This Application.", 16, "Error!!"
Else
If Text1.Text = "c:\windows\system\systray.exe" Then
MsgBox "Windows Is Unable To Close This Application.", 16, "Error!!"
Else
KillApp (Text1.Text)
End If
End If
End If
End If
End If
End If
End If
End If
End Sub
Public Function KillApp(myName As String) As Boolean

    Const PROCESS_ALL_ACCESS = 0
    Dim uProcess As PROCESSENTRY32
    Dim rProcessFound As Long
    Dim hSnapshot As Long
    Dim szExename As String
    Dim exitCode As Long
    Dim myProcess As Long
    Dim AppKill As Boolean
    Dim appCount As Integer
    Dim i As Integer
    On Local Error GoTo Finish
    appCount = 0
    
    Const TH32CS_SNAPPROCESS As Long = 2&
    
    uProcess.dwSize = Len(uProcess)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    rProcessFound = ProcessFirst(hSnapshot, uProcess)
    List1.Clear
    
    Do While rProcessFound
        i = InStr(1, uProcess.szexeFile, Chr(0))
        szExename = LCase$(Left$(uProcess.szexeFile, i - 1))
        List1.AddItem (szExename)
        If Right$(szExename, Len(myName)) = LCase$(myName) Then
            KillApp = True
            appCount = appCount + 1
            myProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
            AppKill = TerminateProcess(myProcess, exitCode)
            Call CloseHandle(myProcess)
        End If


        rProcessFound = ProcessNext(hSnapshot, uProcess)
    Loop


    Call CloseHandle(hSnapshot)
Finish:
End Function

Private Sub Form_Load()
KillApp ("none")
RegisterServiceProcess GetCurrentProcessId, 1 'Hide app

End Sub


Private Sub Form_Resize()
List1.Width = Form1.Width - 400
List1.Height = Form1.Height - 1000
End Sub

Private Sub Form_Unload(Cancel As Integer)
RegisterServiceProcess GetCurrentProcessId, 0 'Remove service flag

End Sub

Private Sub Label4_Click()
Unload Me
Form3.Show
End Sub

Private Sub List1_Click()
Text1.Text = List1.List(List1.ListIndex)

End Sub
Private Sub List1_dblClick()
If Text1.Text = "c:\windows\desktop\general.exe" Then
MsgBox "Windows Is Unable To Close This Application.", 16, "Error!!"
Else
If Text1.Text = "c:\windows\system\kernel32.dll" Then
MsgBox "Windows Is Unable To Close This Application.", 16, "Error!!"
Else
If Text1.Text = "c:\windows\system\msgsrv32.exe" Then
MsgBox "Windows Is Unable To Close This Application.", 16, "Error!!"
Else
If Text1.Text = "c:\windows\system\mprexe.exe" Then
MsgBox "Windows Is Unable To Close This Application.", 16, "Error!!"
Else
If Text1.Text = "c:\windows\system\mmtask.tsk" Then
MsgBox "Windows Is Unable To Close This Application.", 16, "Error!!"
Else
If Text1.Text = "c:\windows\explorer.exe" Then
MsgBox "Windows Is Unable To Close This Application.", 16, "Error!!"
Else
If Text1.Text = "c:\windows\system\wsloader.exe" Then
MsgBox "Windows Is Unable To Close This Application.", 16, "Error!!"
Else
If Text1.Text = "c:\windows\system\systray.exe" Then
MsgBox "Windows Is Unable To Close This Application.", 16, "Error!!"
Else

Text1.Text = List1.List(List1.ListIndex)
KillApp (Text1.Text)
End If
End If
End If
End If
End If
End If
End If
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = "13" Then
KillApp (Text1.Text)
End If
End Sub


