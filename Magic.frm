VERSION 5.00
Begin VB.Form Magic 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MagicMK"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3465
   Icon            =   "Magic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Visible         =   0   'False
   Begin VB.Timer mkTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Magic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    On Error Resume Next
    If App.PrevInstance Then End
    EnableKeyHook
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    UnHookMouse
    UnHookKey
End Sub

Private Sub mkTimer_Timer()
    On Error Resume Next
    timeLine = timeLine + 1
    If status = 2 Then PlayFrame
End Sub
