VERSION 5.00
Object = "{B3FB64BF-91F9-11D7-A482-0008A14158BC}#2.22#0"; "ITGControls.ocx"
Begin VB.Form frmOverride 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Override"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4080
   Icon            =   "frmOverride.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ITGControls.ComunionFrames ComunionFrames2 
      Height          =   1365
      Left            =   0
      Top             =   0
      Width           =   4170
      _ExtentX        =   7355
      _ExtentY        =   2408
      FrameColor      =   16777215
      FillColor       =   16777215
      RoundedCorner   =   0   'False
      Caption         =   ""
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   4
      ColorFrom       =   109785
      ColorTo         =   10874879
      Begin ITGControls.ITGTextBox txtname 
         Height          =   285
         Left            =   135
         TabIndex        =   0
         Top             =   510
         Width           =   3855
         _ExtentX        =   6588
         _ExtentY        =   503
         SendKeysTab     =   -1  'True
         BackColor       =   12648447
         LabelBackColor  =   -2147483624
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "sa"
         Label           =   "User ID"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextBoxWidth    =   2295
      End
      Begin ITGControls.ITGTextBox txtpassword 
         Height          =   285
         Left            =   135
         TabIndex        =   1
         Top             =   885
         Width           =   3855
         _ExtentX        =   6588
         _ExtentY        =   503
         SendKeysTab     =   -1  'True
         BackColor       =   12648447
         LabelBackColor  =   -2147483624
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AllCaps         =   -1  'True
         Passwordchar    =   "*"
         Label           =   "Password"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextBoxWidth    =   2295
      End
   End
End
Attribute VB_Name = "frmOverride"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oFormSetup As New clsFormSetup
Public cModule As String
Public cTranNo As String
Public lApproved As Boolean

Private Sub Form_Load()
    Set FrmName = Me
    If LoadOption("COM_THEME", 4) = "1 - Blue" Then
        oFormSetup.FormTheme (1)
    Else
        oFormSetup.FormTheme (2)
    End If
    lApproved = False
End Sub

Private Sub txtname_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
    End Select
End Sub

Private Sub txtpassword_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
    End Select
End Sub

Private Sub txtpassword_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If tmpRS.State = 1 Then tmpRS.Close
        sSQL = "select * from sec_user where UserID = '" & txtname & "' and RoleID = 'SUPERUSER'"
        tmpRS.Open sSQL, cn, adOpenStatic, adLockReadOnly
        If tmpRS.RecordCount <> 0 Then
            If Decrypt(tmpRS!Password) = txtpassword Then
                lApproved = True
                UpdateLogFile cModule, cTranNo, "Override"
                Unload Me
            End If
        End If
    End If
End Sub

