VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{45D48925-E17D-4DB1-BE71-CD38EC107318}#1.0#0"; "SendEmail.tlb"
Begin VB.Form frmReportViewer 
   Caption         =   "Report Viewer"
   ClientHeight    =   9195
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13110
   Icon            =   "frmReportViewer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9195
   ScaleWidth      =   13110
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin SendEmailCtl.CrypPDF CrypPDF1 
      Height          =   315
      Left            =   9330
      TabIndex        =   1
      Top             =   2220
      Visible         =   0   'False
      Width           =   270
      Object.Visible         =   "True"
      Enabled         =   "True"
      ForegroundColor =   "-2147483630"
      BackgroundColor =   "-2147483633"
      BackColor       =   "Control"
      BackgroundImageLayout=   "Center"
      ForeColor       =   "ControlText"
      Location        =   "622, 148"
      Name            =   "CrypPDF"
      Size            =   "18, 21"
      Object.TabIndex        =   "0"
   End
   Begin CRVIEWERLibCtl.CRViewer ITGReportViewer 
      Height          =   8925
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9165
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmReportViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'IT Group Inc. 2005.09.23

Option Explicit

'Object variables

Public sBit As String

'Security Acess Level variables
Public lACNew As Boolean
Public lACEdit As Boolean
Public lACDelete As Boolean
Public lACPost As Boolean
Public lACCancel As Boolean
Public lACPrint As Boolean

Public rpt As New REPORT

Public sPWord As String

'Activate your Toolbar Mode
Private Sub Form_Activate()
    TBBitReload
End Sub

'Release your Object
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    FormWaitShow "Securing PDF File . . ."

    Dim oFile As New Scripting.FileSystemObject
    oFile.CreateTextFile "c:\tmp.pdf", True
    
    CrypPDF1.AddPasswordToPDF OutputFileName, "c:\tmp.pdf", sPWord
    FileCopy "c:\tmp.pdf", OutputFileName
    Kill "c:\tmp.pdf"

    BitEnabled ITGLedgerMain, Me, ITGLedgerMain.tbrMain, , , , , , , , , , True
    BitVisible ITGLedgerMain.tbrMain

    Unload frmWait
    Set frmReportViewer = Nothing
    
End Sub

'Reload menu buttons (do not delete this sub)
Public Sub TBBitReload()
    On Error Resume Next
    BitVisible ITGLedgerMain.tbrMain
End Sub

'Close active window
Public Sub TBCloseWindow()
    Unload Me
End Sub

Public Sub TBFind()
    '**********
End Sub

Public Sub TBFindPrimary()
    '**********
End Sub

Private Sub Form_Load()
    

    'Set FrmName = Me

    'FormSetup
    
'    BitVisible pITGLedgerMDI.tbrMain, True
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Resize()
    ITGReportViewer.Top = 0
    ITGReportViewer.Left = 0
    ITGReportViewer.Height = ScaleHeight
    ITGReportViewer.Width = ScaleWidth
End Sub


Private Sub ITGReportViewer_PrintButtonClicked(UseDefault As Boolean)
'    Dim rpt As Report
'    Dim obj As New Application
'    Set rpt = obj.OpenReport(pRPTPath & "\" & cReceiptName & ".rpt")
'    rpt.PrinterSetup 0
End Sub

