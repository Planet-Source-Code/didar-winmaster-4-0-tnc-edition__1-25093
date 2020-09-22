VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Windows All Startup Objects."
   ClientHeight    =   3840
   ClientLeft      =   1710
   ClientTop       =   1890
   ClientWidth     =   7095
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List5 
      Height          =   840
      Left            =   6000
      TabIndex        =   11
      Top             =   3840
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4440
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Text            =   "0"
      Top             =   4200
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.CommandButton Command9 
         Caption         =   "Command9"
         Height          =   735
         Left            =   3480
         TabIndex        =   18
         Top             =   1440
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         Caption         =   "System"
         Height          =   375
         Left            =   5040
         TabIndex        =   17
         ToolTipText     =   "Drive Property Window"
         Top             =   2880
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         Caption         =   "DrvClean"
         Height          =   375
         Left            =   3120
         TabIndex        =   15
         ToolTipText     =   "Windows Advanced Startup Objects Cleaning"
         Top             =   2880
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "DrvShow"
         Height          =   375
         Left            =   2160
         TabIndex        =   14
         ToolTipText     =   "Windows Advanced Startup Objects"
         Top             =   2880
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Info"
         Height          =   375
         Left            =   5880
         TabIndex        =   13
         ToolTipText     =   "Information.."
         Top             =   2880
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Back"
         Height          =   375
         Left            =   4080
         TabIndex        =   12
         ToolTipText     =   "Back Window"
         Top             =   2880
         Width           =   855
      End
      Begin VB.ListBox List4 
         Height          =   2205
         ItemData        =   "Form2.frx":0ECA
         Left            =   120
         List            =   "Form2.frx":0ECC
         TabIndex        =   10
         Top             =   600
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   375
         Left            =   3120
         TabIndex        =   8
         Top             =   3720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ListBox List3 
         Height          =   1815
         ItemData        =   "Form2.frx":0ECE
         Left            =   6240
         List            =   "Form2.frx":0ED0
         TabIndex        =   7
         Top             =   3240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.ListBox List2 
         Height          =   2205
         ItemData        =   "Form2.frx":0ED2
         Left            =   2040
         List            =   "Form2.frx":0ED4
         TabIndex        =   5
         Top             =   600
         Width           =   4575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Exe.Clean"
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         ToolTipText     =   "Windows Normal Startup Objects Cleaning"
         Top             =   2880
         Width           =   855
      End
      Begin VB.ListBox List1 
         Height          =   1230
         ItemData        =   "Form2.frx":0ED6
         Left            =   120
         List            =   "Form2.frx":0ED8
         TabIndex        =   2
         Top             =   720
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exe.Show"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Windows Normal StartUp Objects"
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "GSI Powerful Tools For Windows"
         Height          =   375
         Left            =   2040
         TabIndex        =   16
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Windows All StartUp Objects..."
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
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   3015
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Type a4
    a As String * 4
End Type
Private Type l4
    l As Long
End Type





 Const REG_NONE = 0
 Const REG_SZ = 1
 Const REG_EXPAND_SZ = 2
 Const REG_BINARY = 3
 Const REG_DWORD = 4
 Const REG_DWORD_LITTLE_ENDIAN = 4
 Const REG_DWORD_BIG_ENDIAN = 5
 Const REG_LINK = 6
 Const REG_MULTI_SZ = 7
 Const REG_RESOURCE_LIST = 8
 Const REG_FULL_RESOURCE_DESCRIPTOR = 9

 Const HKEY_CLASSES_ROOT = &H80000000
 Const HKEY_CURRENT_USER = &H80000001
 Const HKEY_LOCAL_MACHINE = &H80000002
 Const HKEY_USERS = &H80000003

 Const ERROR_NONE = 0
 Const ERROR_BADDB = 1
 Const ERROR_BADKEY = 2
 Const ERROR_CANTOPEN = 3
 Const ERROR_CANTREAD = 4
 Const ERROR_CANTWRITE = 5
 Const ERROR_OUTOFMEMORY = 6
 Const ERROR_INVALID_PARAMETER = 7
 Const ERROR_ACCESS_DENIED = 8
 Const ERROR_INVALID_PARAMETERS = 87
 Const ERROR_NO_MORE_ITEMS = 259

 Const KEY_ALL_ACCESS = &H3F

 Const SYNCHRONIZE = &H100000
 Const STANDARD_RIGHTS_ALL = &H1F0000
' Reg Key Security Options
 Const KEY_QUERY_VALUE = &H1
 Const KEY_SET_VALUE = &H2
 Const KEY_CREATE_SUB_KEY = &H4
 Const KEY_ENUMERATE_SUB_KEYS = &H8
 Const KEY_NOTIFY = &H10
 Const KEY_CREATE_LINK = &H20
 Const REG_OPTION_NON_VOLATILE = 0
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
 Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
 Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
 Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

 


Public Sub SaveString(hKey As Long, StrPath As String, StrValue As String, StrData As String)
   Dim KeyH&
    r = RegCreateKey(hKey, StrPath, KeyH&)
    r = RegSetValueEx(KeyH&, StrValue, 0, 1, ByVal StrData, Len(StrData))
    r = RegCloseKey(KeyH&)
End Sub





Public Sub delString(hKey As Long, StrPath As String)
   Dim KeyH&
    r = RegDeleteKey(hKey, StrPath)
     r = RegCloseKey(KeyH&)
   End Sub



Public Sub delvalue(hKey As Long, StrPath As String)
   Dim KeyH&
   r = RegOpenKey(hKey, StrPath, KeyH&)
    r = RegDeleteValue(KeyH&, "(Default)")
    r = RegCloseKey(KeyH&)
   End Sub




Sub CreateNewKey(sNewKeyName As String, lPredefinedKey As Long)
    Dim hNewKey As Long
    Dim lRetVal As Long
    
    lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, _
              vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
              0&, hNewKey, lRetVal)
    RegCloseKey (hNewKey)
End Sub






Sub SetKeyValue(sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As Long)
    Dim Zero As Long, IRetVal As Long, hKey As Long, OrigKeyNam As String
   IRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKeyName, Zero, KEY_ALL_ACCESS, hKey)
    If IRetVal Then MsgBox "RegOpenKey error - " & IRetVal
    IRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
    If IRetVal Then MsgBox "SetValue error - " & IRetVal
    RegCloseKey (hKey)
End Sub




    Sub QueryValue(sKeyName As String, sValueName As String)
       Dim lRetVal As Long
       Dim hKey As Long
       Dim vValue As Variant

       lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKeyName, 0, KEY_ALL_ACCESS, hKey)
       lRetVal = QueryValueEx(hKey, sValueName, vValue)
       MsgBox vValue
       RegCloseKey (hKey)
   End Sub


Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long

    Dim lValue As Long
    Dim sValue As String
    Select Case lType
        Case REG_SZ
            sValue = vValue & Chr$(0)
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
        Case REG_DWORD
            lValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)

        End Select

End Function

Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long

    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String

    On Error GoTo QueryValueExError

   
    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc <> ERROR_NONE Then Error 5

    Select Case lType
   
        Case REG_SZ:
            sValue = String(cch, 0)

 lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)

            If lrc = ERROR_NONE Then
                vValue = Left$(sValue, cch)
            Else
                vValue = Empty
            End If
   
        Case REG_DWORD:



lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)

            If lrc = ERROR_NONE Then vValue = lValue
        Case Else
   
            lrc = -1
    End Select




QueryValueExExit:

    QueryValueEx = lrc
    Exit Function



QueryValueExError:

    Resume QueryValueExExit
End Function







Private Sub Command1_Click()
    Dim lRet As Long, hKey As Long
    List4.Clear
    List2.Clear
    lRet = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", 0, KEY_ALL_ACCESS, hKey)
    If lRet Then MsgBox "Error Accessing System's DNA Structure.": Exit Sub
    
    Dim lIndex As Long, aVName$, lVName As Long, lType As Long, aData$, lData As Long
    Dim aAdd$, l As Long, a4 As a4, l4 As l4, a$
    lVName = 100
    aVName$ = Space$(lVName)
    lData = 100
    aData$ = Space$(lData)
    lRet = RegEnumValue(hKey, lIndex, aVName$, lVName, 0, lType, aData$, lData)
    Do Until lRet = ERROR_NO_MORE_ITEMS
        aAdd$ = Left$(aVName$, lVName) & vbTab
        Do Until TextWidth(aAdd$) > List1.Width \ 4: aAdd$ = aAdd$ & " ": Loop
        aAdd$ = aAdd$ & vbTab
        Select Case lType
            Case REG_BINARY
                For l = 1 To lData
                    a$ = Hex$(Asc(Mid$(aData$, l, 1)))
                    If Len(a$) = 1 Then a$ = "0" & a$
                    aAdd$ = aAdd$ & a$ & " "
                                        
                Next
            Case REG_DWORD
                a4.a = Left$(aData$, lData)
                LSet l4 = a4
                aAdd$ = aAdd$ & l4.l
                
            Case Else
            dna$ = aAdd$
            
            aAdd$ = aAdd$ & Left$(aData$, lData)
                
        End Select
        List1.AddItem aAdd$
        List2.AddItem Left$(aData$, lData)
        List4.AddItem dna
                
        lVName = 100
        lData = 100
        lIndex = lIndex + 1
        lRet = RegEnumValue(hKey, lIndex, aVName$, lVName, 0, lType, aData$, lData)
    Loop
        
    lRet = RegCloseKey(hKey)

End Sub

Private Sub Command2_Click()
Dim i As Integer
On Error Resume Next
For i = 0 To List4.ListCount - 1
If List4.Selected(i) Then
List5.AddItem List4.List(i)
End If


Next
Text2.Text = List5.List(0)
k = Len(Text2.Text)
Text2.Text = Left(Text2.Text, (k - 2))
Text1.Text = List3.List(0)
Command9_Click
End Sub









Private Sub Command3_Click()
If Command1.Value = 1 Then
SaveSetting App.Title, App.Title, "RunWithSystem", 1
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", Text2.Text, Text2.Text
Else
SaveSetting App.Title, App.Title, "RunWithSystem", 0
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", Text2.Text, 0
End If

End Sub


Private Sub Command4_Click()
Unload Me
FrmMain.Show
End Sub

Private Sub Command5_Click()
Form1.Show
End Sub

Private Sub Command6_Click()
Unload Me
Form3.Show
End Sub

Private Sub Command7_Click()
    Dim lRet As Long, hKey As Long
    On Error Resume Next
    List4.Clear
    List2.Clear
    lRet = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices", 0, KEY_ALL_ACCESS, hKey)
    If lRet Then MsgBox "Error Accessing System's DNA Structure.": Exit Sub
    
    Dim lIndex As Long, aVName$, lVName As Long, lType As Long, aData$, lData As Long
    Dim aAdd$, l As Long, a4 As a4, l4 As l4, a$
    lVName = 100
    aVName$ = Space$(lVName)
    lData = 100
    aData$ = Space$(lData)
    lRet = RegEnumValue(hKey, lIndex, aVName$, lVName, 0, lType, aData$, lData)
    Do Until lRet = ERROR_NO_MORE_ITEMS
        aAdd$ = Left$(aVName$, lVName) & vbTab
        Do Until TextWidth(aAdd$) > List1.Width \ 4: aAdd$ = aAdd$ & " ": Loop
        aAdd$ = aAdd$ & vbTab
        Select Case lType
            Case REG_BINARY
                For l = 1 To lData
                    a$ = Hex$(Asc(Mid$(aData$, l, 1)))
                    If Len(a$) = 1 Then a$ = "0" & a$
                    aAdd$ = aAdd$ & a$ & " "
                                        
                Next
            Case REG_DWORD
                a4.a = Left$(aData$, lData)
                LSet l4 = a4
                aAdd$ = aAdd$ & l4.l
                
            Case Else
            dna$ = aAdd$
            
            aAdd$ = aAdd$ & Left$(aData$, lData)
                
        End Select
        List1.AddItem aAdd$
        List2.AddItem Left$(aData$, lData)
        List4.AddItem dna
                
        lVName = 100
        lData = 100
        lIndex = lIndex + 1
        lRet = RegEnumValue(hKey, lIndex, aVName$, lVName, 0, lType, aData$, lData)
    Loop
        
    lRet = RegCloseKey(hKey)


End Sub

Private Sub Command8_Click()
Dim i As Integer
  Dim KeyH&
On Error Resume Next
For i = 0 To List4.ListCount - 1
If List4.Selected(i) Then
List5.AddItem List4.List(i)
End If
Next
Text2.Text = List5.List(0)
k = Len(Text2.Text)
Text2.Text = Left(Text2.Text, (k - 2))
Text1.Text = List3.List(0)



 r = RegOpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices", KeyH&)
    r = RegDeleteValue(KeyH&, Text2.Text)
    r = RegCloseKey(KeyH&)

List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
Command7_Click
End Sub


Private Sub Command9_Click()
   Dim KeyH&
 r = RegOpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", KeyH&)
    r = RegDeleteValue(KeyH&, Text2.Text)
    r = RegCloseKey(KeyH&)
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
Command1_Click
End Sub
