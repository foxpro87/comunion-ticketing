VERSION 5.00
Begin VB.Form frmAccountingSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAccountingSplash.frx":0000
   ScaleHeight     =   303
   ScaleMode       =   0  'User
   ScaleWidth      =   488
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Close"
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   3480
   End
End
Attribute VB_Name = "frmAccountingSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Timer1.Interval = 2000
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAccountingSplash = Nothing
End Sub

Private Sub Timer1_Timer()
    Timer1.Interval = 2000
    If Timer1.Interval = 2000 Then
        Unload Me
    End If
End Sub
