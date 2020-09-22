VERSION 5.00
Begin VB.Form Form4 
   ClientHeight    =   5235
   ClientLeft      =   2685
   ClientTop       =   780
   ClientWidth     =   4155
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   Picture         =   "Form4.frx":0ECA
   ScaleHeight     =   5235
   ScaleWidth      =   4155
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   1200
      Top             =   2160
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Unload Me
FrmMain.Show
Timer1.Enabled = False
End Sub
