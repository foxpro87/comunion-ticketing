VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{B3FB64BF-91F9-11D7-A482-0008A14158BC}#2.10#0"; "itgcontrols.ocx"
Begin VB.Form frmWait 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4680
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ITGControls.ITGLabel ITGLabel4 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   1800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      Caption         =   "of 0 Records"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ITGControls.ITGLabel ITGLabel3 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   1800
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
      Caption         =   "0 %"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.Animation aniTransmit 
      Height          =   735
      Left            =   480
      TabIndex        =   3
      Top             =   720
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1296
      _Version        =   393216
      AutoPlay        =   -1  'True
      Center          =   -1  'True
      FullWidth       =   257
      FullHeight      =   49
   End
   Begin ITGControls.ITGLabel ITGLabel2 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      Caption         =   "Please Wait..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ITGControls.ITGLabel ITGLabel1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   135
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
      Caption         =   "Loading...."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'IT Group Inc. 2005.09.23

Option Explicit

Private Sub Form_Load()
    DoEvents
    Screen.MousePointer = vbHourglass
    With aniTransmit
        .AutoPlay = True
        .Open Trim(App.Path) & "\Transmit.avi"
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
    aniTransmit.AutoPlay = False
End Sub

