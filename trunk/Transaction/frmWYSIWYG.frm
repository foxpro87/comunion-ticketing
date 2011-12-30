VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{B3FB64BF-91F9-11D7-A482-0008A14158BC}#2.22#0"; "ITGControls.ocx"
Begin VB.Form frmWYSIWYG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WYSIWYG"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6870
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rtfTicket 
      Height          =   6555
      Left            =   45
      TabIndex        =   2
      Top             =   75
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   11562
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmWYSIWYG.frx":0000
   End
   Begin ITGControls.ComunionButton cmdPrint 
      Height          =   375
      Left            =   4665
      TabIndex        =   0
      Top             =   345
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Print"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmWYSIWYG.frx":008B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ITGControls.ComunionButton cmdClose 
      Height          =   375
      Left            =   5685
      TabIndex        =   1
      Top             =   345
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmWYSIWYG.frx":00A7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmWYSIWYG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
'    PrintRTF rtfTicket, 300, 300, 500, 500  ' 1440 Twips = 1 Inch
    'WYSIWYG_RTF rtfTicket, 300,300
End Sub

Private Sub Form_Load()
'    rtfTicket.LoadFile App.Path & "\Reports\TicketFormat.rtf", rtfRTF
'
'    If rs.State = 1 Then rs.Close
'    sSQL = "select * from sales"
'    rs.Open sSQL, cn, adOpenStatic, adLockReadOnly
'    rtfTicket.TextRTF = Replace(rtfTicket.TextRTF, "<cTranNo>", rs!cTranNo & "")
'    rtfTicket.TextRTF = Replace(rtfTicket.TextRTF, "<dDateTime1>", TimeValue(rs!dDateTime) & "")
'    rtfTicket.TextRTF = Replace(rtfTicket.TextRTF, "<dDateTime2>", DateValue(rs!dDateTime) & "")
'    rtfTicket.TextRTF = Replace(rtfTicket.TextRTF, "<cType>", rs!cType & "")
'    rtfTicket.TextRTF = Replace(rtfTicket.TextRTF, "<nNetPrice>", Format(rs!nNetPrice, "00.00") & "")
'
End Sub
