VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   ClientHeight    =   3180
   ClientLeft      =   2580
   ClientTop       =   3240
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   7260
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1680
      Top             =   1320
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "EXE NAME!"
      Top             =   120
      Width           =   5415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "KILL"
      Height          =   255
      Left            =   5760
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
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
KillApp (Text1.Text)
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
Text1.Width = Form1.Width - Command1.Width - 300
Command1.Left = Text1.Width + 150
End Sub

Private Sub Form_Unload(Cancel As Integer)
RegisterServiceProcess GetCurrentProcessId, 0 'Remove service flag

End Sub

Private Sub List1_Click()
Text1.Text = List1.List(List1.ListIndex)

End Sub
Private Sub List1_dblClick()
Text1.Text = List1.List(List1.ListIndex)
KillApp (Text1.Text)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = "13" Then
KillApp (Text1.Text)
End If
End Sub

