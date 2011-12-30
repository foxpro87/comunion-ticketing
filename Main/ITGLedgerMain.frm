VERSION 5.00
Object = "{45D48925-E17D-4DB1-BE71-CD38EC107318}#1.0#0"; "SendEmail.tlb"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.MDIForm ITGLedgerMain 
   BackColor       =   &H00B49799&
   ClientHeight    =   8895
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   13110
   Icon            =   "ITGLedgerMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13110
      _ExtentX        =   23125
      _ExtentY        =   688
      BandCount       =   2
      FixedOrder      =   -1  'True
      _CBWidth        =   13110
      _CBHeight       =   390
      _Version        =   "6.7.9782"
      Child1          =   "tbrMain"
      MinWidth1       =   12000
      MinHeight1      =   330
      Width1          =   12000
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Caption2        =   "Division"
      Child2          =   "cboSelection"
      MinHeight2      =   315
      Width2          =   5145
      NewRow2         =   0   'False
      Visible2        =   0   'False
      AllowVertical2  =   0   'False
      Begin VB.ComboBox cboSelection 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "ITGLedgerMain.frx":0CCA
         Left            =   12945
         List            =   "ITGLedgerMain.frx":0CCC
         TabIndex        =   3
         Top             =   30
         Width           =   390
      End
      Begin MSComctlLib.Toolbar tbrMain 
         Height          =   330
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   12000
         _ExtentX        =   21167
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ITGImageList"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   26
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "btnNew"
               Object.ToolTipText     =   "New (Ctrl+N)"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "btnEdit"
               Object.ToolTipText     =   "Edit (Ctrl+E)"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "btnDelete"
               Object.ToolTipText     =   "Delete (Ctrl+D)"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "btnNewLine"
               Object.ToolTipText     =   "New Line (Ctrl+Ins)"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "btnDeleteLine"
               Object.ToolTipText     =   "Delete Line (Ctrl+Del)"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "btnSepDetail"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Key             =   "btnPost"
               Object.ToolTipText     =   "Post"
               ImageIndex      =   15
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Key             =   "btnCancel"
               Object.ToolTipText     =   "Cancel"
               ImageIndex      =   16
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "btnSepPostCancel"
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "btnSave"
               Object.ToolTipText     =   "Save (Ctrl+S)"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "btnUndo"
               Object.ToolTipText     =   "Undo Current (Ctrl+Z)"
               ImageIndex      =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "btnUndoCurrent"
                     Text            =   "Undo Current"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Enabled         =   0   'False
                     Key             =   "btnUndoAll"
                     Text            =   "Undo All"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Object.Visible         =   0   'False
                     Key             =   "btnUndoCurrentLine"
                     Text            =   "Undo Current Line"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "btnFind"
               Object.ToolTipText     =   "Find (Crtl+F)"
               ImageIndex      =   6
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Enabled         =   0   'False
                     Key             =   "btnFindP"
                     Text            =   "Primary"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "btnFindC"
                     Text            =   "Comprehensive"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "btnFirst"
               Object.ToolTipText     =   "First Record"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "btnPrevious"
               Object.ToolTipText     =   "Previous Record"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "btnNext"
               Object.ToolTipText     =   "Next Record"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "btnLast"
               Object.ToolTipText     =   "Last Record"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "btnPrint"
               Object.ToolTipText     =   "Print (Ctrl+P)"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "btnClose"
               Object.ToolTipText     =   "Close Window (Ctrl+X)"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "btnAccount"
               Object.ToolTipText     =   "Account Affected"
               ImageIndex      =   17
            EndProperty
            BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "btnReference"
               Object.ToolTipText     =   "Reference Transactions"
               ImageIndex      =   18
            EndProperty
            BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "btnNotify"
               Object.ToolTipText     =   "Alert Message"
               ImageIndex      =   19
            EndProperty
         EndProperty
         Begin VB.PictureBox Picture2 
            Height          =   345
            Left            =   7080
            ScaleHeight     =   285
            ScaleWidth      =   1455
            TabIndex        =   5
            Top             =   -15
            Visible         =   0   'False
            Width           =   1515
            Begin SendEmailCtl.SendEmail SendEmail1 
               Height          =   345
               Left            =   165
               TabIndex        =   6
               Top             =   45
               Width           =   375
               Object.Visible         =   "True"
               Enabled         =   "True"
               ForegroundColor =   "-2147483640"
               BackgroundColor =   "11835289"
               EmailUsername   =   ""
               EmailPassword   =   ""
               EmailSMTP       =   ""
               EmailPORT       =   "0"
               EmailFrom       =   ""
               EmailTo         =   ""
               EmailSubject    =   ""
               EmailBody       =   ""
               EmailAttachment =   ""
               BackColor       =   "153, 151, 180"
               BackgroundImageLayout=   "Center"
               ForeColor       =   "WindowText"
               Location        =   "11, 3"
               Name            =   "SendEmail"
               Size            =   "25, 23"
               Object.TabIndex        =   "0"
            End
         End
         Begin VB.PictureBox Picture1 
            Height          =   300
            Left            =   10695
            ScaleHeight     =   240
            ScaleWidth      =   420
            TabIndex        =   4
            Top             =   15
            Visible         =   0   'False
            Width           =   480
         End
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   300
      Left            =   960
      Top             =   1080
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   540
      Top             =   1080
   End
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   8610
      Width           =   13110
      _ExtentX        =   23125
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   794
            MinWidth        =   794
            Picture         =   "ITGLedgerMain.frx":0CCE
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   6349
            MinWidth        =   6349
            Text            =   "Providing innovative IT solutions to cater to your business needs."
            TextSave        =   "Providing innovative IT solutions to cater to your business needs."
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3686
            MinWidth        =   88
            Text            =   "User"
            TextSave        =   "User"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3600
            MinWidth        =   2
            Text            =   "Department"
            TextSave        =   "Department"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3600
            MinWidth        =   2
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4480
            MinWidth        =   882
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ITGImageList 
      Left            =   0
      Top             =   1035
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITGLedgerMain.frx":2A3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITGLedgerMain.frx":2FD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITGLedgerMain.frx":355E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITGLedgerMain.frx":3AE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITGLedgerMain.frx":406A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITGLedgerMain.frx":4604
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITGLedgerMain.frx":4B8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITGLedgerMain.frx":5124
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITGLedgerMain.frx":56BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITGLedgerMain.frx":5C58
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITGLedgerMain.frx":61F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITGLedgerMain.frx":678C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITGLedgerMain.frx":6D26
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITGLedgerMain.frx":8BA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITGLedgerMain.frx":9482
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITGLedgerMain.frx":9A1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITGLedgerMain.frx":9FB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITGLedgerMain.frx":A408
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ITGLedgerMain.frx":A85A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Enabled         =   0   'False
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuFileDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileNewLine 
         Caption         =   "New Line (Ctrl+Ins)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileDeleteLine 
         Caption         =   "Delete Line (Ctrl+Del)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileUndo 
         Caption         =   "Undo Current"
         Enabled         =   0   'False
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuFileUndoAll 
         Caption         =   "Undo All"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileFind 
         Caption         =   "&Find..."
         Enabled         =   0   'False
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
         Enabled         =   0   'False
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "Close Window"
         Enabled         =   0   'False
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileLogOff 
         Caption         =   "Log-off"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTransaction 
      Caption         =   "&Transaction"
      Begin VB.Menu mnuTicketing 
         Caption         =   "Sales Ticketing"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuEmpReport 
         Caption         =   "Employee Sales Report"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReport 
         Caption         =   "Reports"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuMaint 
      Caption         =   "&Maintenance"
      Begin VB.Menu mnuMaintCompany 
         Caption         =   "Company"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuTicketType 
         Caption         =   "Ticket Type"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuTicketPrice 
         Caption         =   "Ticket Price"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsAuditTrail 
         Caption         =   "Audit Trail"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuToolsSetup 
         Caption         =   "Setup..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEmailSetup 
         Caption         =   "Email Setup"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuToolReprint 
         Caption         =   "Reprint Ticket"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep12 
         Caption         =   "-"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuSec 
      Caption         =   "&Security"
      Begin VB.Menu mnuSecChangePassword 
         Caption         =   "Change Password"
      End
      Begin VB.Menu mnuSecUserProfile 
         Caption         =   "User Profile"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About ComUnion"
      End
   End
   Begin VB.Menu mnuImage 
      Caption         =   "Image"
      Visible         =   0   'False
      Begin VB.Menu mnuImageAdd 
         Caption         =   "Add Image"
      End
      Begin VB.Menu mnuImageClear 
         Caption         =   "Clear Image"
      End
   End
End
Attribute VB_Name = "ITGLedgerMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'IT Group Inc. 2005.09.23
Option Explicit

Private oBar As New clsToolBarMenuBit
Public sBit As String
Private lBlockLogOn As Boolean

Public lSaleLoaded As Boolean
Dim lAdmin As Boolean

Private Sub cboSelection_Change()
    'sDivision = GetValueFrTable("cDivisionID", "PMS_DIVISION", "cDivName = '" & Trim(cboSelection.Text) & "'")
End Sub

Private Sub cboSelection_Click()
    'sDivision = GetValueFrTable("cDivisionID", "PMS_DIVISION", "cDivName = '" & Trim(cboSelection.Text) & "'")
End Sub

Private Sub MDIForm_Activate()
    Me.Caption = "ComUnion" & ": " & cCompany
End Sub

Private Sub MDIForm_Load()
    Dim sLine As String
    Dim lPos As Integer

    Dim strConn As String
    Dim rsFormAccess As New Recordset
    Dim cmdFormAccess As New ADODB.Command

    GetAccessLevel
    lAdmin = IIf(GetValueFrTable("RoleID", "Sec_User", "UserID='" & SecUserID & "'", True) = "SUPERUSER", True, False)

    oBar.BitEnabled Me, tbrMain, , , , , , , , , , True
    oBar.BitVisible tbrMain
    COName = GetValueFrTable("cCompanyName", "Company", " cCompanyID = '" & COID & "' ", False)
    cAddress1 = GetValueFrTable("cAddress1", "Company", " cCompanyID = '" & COID & "' ", False)
    cAddress2 = GetValueFrTable("cAddress2", "Company", " cCompanyID = '" & COID & "' ", False)
    
    sbMain.Panels(3) = SecUserName
    sbMain.Panels(4) = SecUserID
    sbMain.Panels(5) = COID
    
    If lRegistered = False Then
        sbMain.Panels(6) = "Unregistered Version" 'Format(Now, "dddd, mmmm dd, yyyy hh:MM:ss AM/PM")
    End If
    
    'Synchronize terminal time to the server
    If rs.State = 1 Then rs.Close
    rs.Open "select getdate() as sReturn", cn, adOpenStatic, adLockReadOnly
    Date = Format(rs!sReturn, "mm/dd/yyyy")
    Time = Format(rs!sReturn, "HH:MM:SS AM/PM")
    dLoginTime = Now
    rs.Close
    
    SetTerminal
    
'    ShowPicture
    InitEmailTime
    CurrentYear = GetValueFrTable("cValue", "PArameter_User", "cParamName='cSystemCurrentYear'", True)
    CurrentMonth = GetValueFrTable("cValue", "PArameter_User", "cParamName='cSystemCurrentMonth'", True)

    Open App.path & "\reports.ini" For Input Access Read As #1
        Do While Not EOF(1)
            Input #1, sLine
            If InStr(sLine, "USERREPORT") Then
                sLine = Trim(sLine)
                lPos = InStr(sLine, "=")
                EmpReport = Trim(Right(sLine, Len(sLine) - lPos))
            ElseIf InStr(sLine, "REPORTPASSWORD") Then
                sLine = Trim(sLine)
                lPos = InStr(sLine, "=")
                If UCase(Trim(Right(sLine, Len(sLine) - lPos))) = "TRUE" Then
                    lUsePassword = True
                Else
                    lUsePassword = False
                End If
            End If
        Loop
    Close #1


End Sub

Sub SetTerminal()
    sThermalPrinterName = GetValueFrTable("cValue", "parameter_user", "cParamName='Printer Name'", True)

    Dim sHostName As String
    Dim nValue As Integer
    If rs.State = 1 Then rs.Close
    rs.Open "select HOST_NAME() as sString", cn, adOpenStatic, adLockReadOnly
    sHostName = rs!sString: rs.Close
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from parameter_user where cType = 'Terminal' and cCompanyID = '" & COID & "'", cn, adOpenStatic, adLockReadOnly
    nValue = rs.RecordCount + 1: rs.Close
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from parameter_user where cParamName = '" & sHostName & "'", cn, adOpenStatic, adLockReadOnly
    If rs.RecordCount = 0 Then  'no record
        cn.Execute "insert into parameter_user (cCompanyID,cType,cParamName,cValue) values ('" & COID & "','Terminal','" & sHostName & "','" & nValue & "')"
        SecTerminalID = nValue
    Else                                            'get Terminal ID
        SecTerminalID = rs!cValue
    End If
    rs.Close
End Sub

Sub CheckEmailTime()
    Dim i, j As Integer
    If lAdmin = False Then Exit Sub
    For i = 1 To UBound(TimeCollection)
        If IsReportSend(i) = False Then
            If PassTime <> Format(Time, "HH:MM") Then
                If IsNumeric(TimeCollection(i)) = True Then
                    Dim a
                    Dim b
                    Dim c(0 To 1) As Integer
                    PassTime = Format(Time, "HH:MM")
                    a = Split(Format(Time, "HH:MM"), ":")
                    a(0) = a(0) - 1
                    a(1) = a(1) + 60
                    
                    b = Split(TimeCollection(i), ".")
                    If UBound(b) = 1 Then b(1) = 60 * (b(1) * 0.01)
                    
                    c(0) = a(0) Mod b(0)
                    If UBound(b) = 1 Then
                        c(1) = a(1) Mod b(1)
                    Else
                        c(1) = a(1)
                    End If
                    
                    If c(0) + c(1) = 0 Then
                        'We Will Now send the email
                        If rs.State = 1 Then rs.Close
                        rs.Open "select * from MailSetup where (dTime = '" & TimeCollection(i) & "' or dTimer = '" & TimeCollection(i) & "') and cCompanyID = '" & COID & "'", cn, adOpenStatic, adLockReadOnly
                        For j = 0 To rs.RecordCount - 1
                            FormWaitShow "Sending email . . ."
                            With SendEmail1
                                .EmailUsername = rs!cUserID
                                .EmailPassword = rs!cPassword
                                .EmailSMTP = rs!cSMTP
                                .EmailPORT = rs!cPORT
                                .EmailFrom = rs!cFrom
                                .EmailTo = rs!cTo
                                .EmailSubject = rs!cSubject
                                .EmailBody = rs!cBody
                                .EmailAttachment = rs!cAttachment
                                .vSendEmail
                            End With
                            Unload frmWait
                            rs.MoveNext
                        Next j
                    End If
                End If
            Else
                If Format(Time, "HH:MM") = Format(TimeCollection(i), "HH:MM") Then
                    IsReportSend(i) = True
                    'We Will Now send the email
                    If rs.State = 1 Then rs.Close
                    rs.Open "select * from MailSetup where dTime = '" & TimeCollection(i) & "' and cCompanyID = '" & COID & "'", cn, adOpenStatic, adLockReadOnly
                    For j = 0 To rs.RecordCount - 1
                        FormWaitShow "Sending email . . ."
                        With SendEmail1
                            .EmailUsername = rs!cUserID
                            .EmailPassword = rs!cPassword
                            .EmailSMTP = rs!cSMTP
                            .EmailPORT = rs!cPORT
                            .EmailFrom = rs!cFrom
                            .EmailTo = rs!cTo
                            .EmailSubject = rs!cSubject
                            .EmailBody = rs!cBody
                            .EmailAttachment = rs!cAttachment
                            .vSendEmail
                        End With
                        Unload frmWait
                        rs.MoveNext
                    Next j
                End If
            End If
        End If
    Next i
End Sub

Sub ShowPicture()
Dim stm As New ADODB.Stream
Dim rs As New Recordset

    sSQL = "Select iPhoto From COMPANY Where cCompanyID = '" & COID & "'"
    rs.Open sSQL, cn, adOpenKeyset
    
    If (RTrim(LTrim(rs!iPhoto)) & "") <> "" Then
        stm.Type = adTypeBinary
        stm.Open
        stm.Write rs!iPhoto
        stm.SaveToFile "\logo.tmp"
        
        Picture1.Picture = LoadPicture("\logo.tmp")
        Kill ("\logo.tmp")
        stm.Close
    End If
    Set stm = Nothing
    Set rs = Nothing
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next

    If Not (ActiveForm Is Nothing) Then
        Cancel = True
        MsgBox "There are Active form(s)... Unable to exit application.", vbCritical, Caption
        lBlockLogOn = True
    Else
        sSQL = "UPDATE SEC_USER SET PassErrCtr = 0, Locked = 0, InUse = 0 WHERE UserID = '" & Trim(sUserName) & "' "
        cn.Execute sSQL
    End If

End Sub

Private Sub MDIForm_Terminate()
    Set oBar = Nothing
    Set ITGLedgerMain = Nothing
    End
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Set rs = Nothing
    Set cn = Nothing
    Exit Sub
End Sub

Private Sub mnuEmpReport_Click()
'    Dim oRPT As New clsPrinting
'
'    Dim MyDate As Date
'    Dim tmpDate As String
'a:
'    tmpDate = InputBox("Please enter a date", "Report Parameter", Format(dTranDate, "mm/dd/yyyy"))
'    If tmpDate = "" Then
'        Exit Sub
'    ElseIf IsDate(tmpDate) = False Then
'        MsgBox "Please enter a valid date.", vbCritical + vbOKOnly, "ComUnion"
'        GoTo a
'    End If
'    MyDate = Format(tmpDate, "mm/dd/yyyy")
'
'
'    FormWaitShow "Loading Report. . ."
'    With oRPT
'        .pModule = EmpReport
'        .pCOID = COID
'        .sDateTran = MyDate 'Date$
'        .RPTPath = App.path & "\Reports"
'        Set .pCN = cn
'        .PrintReceipt
'    End With
'    Set oRPT = Nothing
'    Unload frmWait
    TimeFrTo.Show
End Sub

Private Sub mnuFileClose_Click()
    CloseWindow
    If ActiveForm Is Nothing Then
        oBar.BitEnabled Me, tbrMain, , , , , , , , , , True
        oBar.BitVisible tbrMain
        tbrMain.Buttons("btnFind").ButtonMenus("btnFindP").Enabled = False
    End If
End Sub

Private Sub mnuFileDelete_Click()
    DeleteRecord
End Sub

Private Sub mnuFileDeleteLine_Click()
    DeleteLine
End Sub

Private Sub mnuFileEdit_Click()
    EditRecord
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub mnuFileFind_Click()
    On Error Resume Next
    ActiveForm.TBFindPrimary
End Sub

Private Sub mnuFileLogOff_Click()
    Unload Me
    If lBlockLogOn Then
        lBlockLogOn = False
        Exit Sub
    End If
    Load frmLogon
    frmLogon.Show
End Sub

Private Sub mnuFileNew_Click()
    AddNewRecord
End Sub

Private Sub mnuFileNewLine_Click()
    NewLine
End Sub

Private Sub mnuFilePrint_Click()
    PrintRecord
End Sub

Private Sub mnuFileSave_Click()
    SaveRecord
End Sub
Private Sub mnuFileUndo_Click()
    UndoAllRecord
End Sub

Private Sub mnuFileUndoAll_Click()
    UndoAllRecord
End Sub

Private Sub mnuImageAdd_Click()
    ActiveForm.TBBrowseImage
End Sub

Private Sub mnuImageClear_Click()
    ActiveForm.TBClearImage
End Sub

Private Sub mnuToolReprint_Click()
    frmToolReprint.Show vbModal
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "btnNew"
            mnuFileNew_Click
        Case "btnEdit"
            mnuFileEdit_Click
        Case "btnDelete"
            mnuFileDelete_Click
        Case "btnNewLine"
            mnuFileNewLine_Click
        Case "btnDeleteLine"
            mnuFileDeleteLine_Click
'        Case "btnPost"
'            mnuFilePost_Click
'        Case "btnCancel"
'            mnuFileCancel_Click
        Case "btnSave"
            mnuFileSave_Click
        Case "btnUndo"
            mnuFileUndo_Click
        Case "btnFind"
            If Button.ButtonMenus("btnFindP").Enabled Then
                mnuFileFind_Click
'            Else
'                mnuFileFindC_Click
            End If
        Case "btnFirst"
            MoveFirst
        Case "btnPrevious"
            MovePrevious
        Case "btnNext"
            MoveNext
        Case "btnLast"
            MoveLast
        Case "btnPrint"
            mnuFilePrint_Click
        Case "btnClose"
            mnuFileClose_Click
'        Case "btnAccount"
'            mnuFileAccount_Click
'        Case "btnReference"
'            mnuFileReference_Click
'        Case "btnNotify"
'            mnuFileNotify_Click
    End Select
End Sub

Private Sub NewLine()
    ActiveForm.TBNewLine
End Sub

Private Sub DeleteLine()
    ActiveForm.TBDeleteLine
End Sub

Private Sub PostRecord()
    ActiveForm.TBPostRecord
End Sub

Private Sub CancelRecord()
    ActiveForm.TBCancelRecord
End Sub

Private Sub PrintRecord()
    ActiveForm.TBPrintRecord
End Sub

Private Sub AddNewRecord()
    ActiveForm.TBNew
End Sub

Private Sub SaveRecord()
    ActiveForm.TBSave
End Sub

Private Sub UndoCurrentRecord()
    ActiveForm.TBUndoCurrent
    ActiveForm.TBUndoCurrent
End Sub

Private Sub UndoAllRecord()
    ActiveForm.TBUndoAll
End Sub

Private Sub EditRecord()
    ActiveForm.TBEdit
End Sub

Private Sub DeleteRecord()
    ActiveForm.TBDelete
End Sub

Private Sub AccountAffected()
    ActiveForm.TBAccountAffected
End Sub

Private Sub ReferenceTrans()
    ActiveForm.TBReference
End Sub

Private Sub NotifyTrans()
    ActiveForm.TBNotify
End Sub

Private Sub FindRecord()
    If ActiveForm Is Nothing Then
        'frmITGSearch.Show 'vbModal
    Else
        ActiveForm.TBFind
    End If
End Sub

Private Sub CloseWindow()
    ActiveForm.TBCloseWindow
End Sub

Private Sub MoveFirst()
    ActiveForm.TBFirstRec
End Sub

Private Sub MovePrevious()
    ActiveForm.TBPrevRec
End Sub

Private Sub MoveNext()
    ActiveForm.TBNextRec
End Sub

Private Sub MoveLast()
    ActiveForm.TBLastRec
End Sub

Private Sub tbrMain_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "btnUndoCurrent"
            mnuFileUndo_Click
        Case "btnUndoAll"
            mnuFileUndoAll_Click
'        Case "btnUndoCurrentLine"
'            mnuDetailUndoCurrent_Click
'        Case "btnFindC"
'            mnuFileFindC_Click
        Case "btnFindP"
            ActiveForm.TBFindPrimary
    End Select
End Sub

Private Sub tbrMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo TheSource
Dim lError As Boolean
    lError = True
        If Me.ActiveForm.Name = "frmARAdjustments" Then
            tbrMain.Buttons("btnPost").ToolTipText = "Approve"
            lError = False
        Else
            tbrMain.Buttons("btnPost").ToolTipText = "Post"
        End If
TheSource:
    If lError Then tbrMain.Buttons("btnPost").ToolTipText = "Post"
    Resume Next
End Sub

Private Sub Timer1_Timer()
    'sbMain.Panels(6) = Format(Now, "dddd, mmmm dd, yyyy") ' hh:MM:ss AM/PM")
    If lRegistered = False Then Exit Sub
    If lSaleLoaded = False Then If sUserName <> "SA" Then lSaleLoaded = True:   frmSaleTicketing.Show vbModal
    CheckEmailTime
    If Format(Time, "HH:MM") = "23:59" Then InitEmailTime
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
    If lCloseWindow Then
        If Not (ActiveForm Is Nothing) Then ActiveForm.TBBitReload
        lCloseWindow = False
    End If
End Sub

Sub GetAccessLevel()
    Dim S As String
    S = GetValueFrTable("RoleID", "SEC_USER", "UserID='" & SecUserID & "'", True)

    If S = "SUPERUSER" Then
        mnuMaintCompany.Enabled = lRegistered
        mnuSecUserProfile.Enabled = lRegistered
        mnuTicketType.Enabled = lRegistered
        mnuTicketPrice.Enabled = lRegistered
        mnuEmailSetup.Enabled = lRegistered
        mnuToolsSetup.Enabled = lRegistered
        mnuToolsAuditTrail.Enabled = lRegistered
        mnuToolReprint.Enabled = lRegistered
        mnuReport.Enabled = lRegistered
        
        mnuTicketing.Enabled = lRegistered
        mnuEmpReport.Enabled = lRegistered
        mnuSecChangePassword.Enabled = lRegistered
    Else
        mnuTicketing.Enabled = lRegistered
        mnuEmpReport.Enabled = lRegistered
        mnuSecChangePassword.Enabled = lRegistered
    End If

End Sub

Private Sub mnuEmailSetup_Click()
    frmMailSetup.Show
    frmMailSetup.Move 0, 0
    frmMailSetup.ZOrder
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuReport_Click()
    frmDatePicker.ReportManagement = False
    frmDatePicker.Show
    frmDatePicker.Move 0, 0
    frmDatePicker.ZOrder
End Sub

Private Sub mnuSecChangePassword_Click()
    frmSecChangePassword.Show
End Sub
 
Private Sub mnuSecUserProfile_Click()
    frmSecUserProfile.Show
    frmSecUserProfile.Move 0, 0
    frmSecUserProfile.ZOrder
End Sub

Private Sub mnuTicketing_Click()
    frmSaleTicketing.Show 1
End Sub

Private Sub mnuTicketPrice_Click()
    frmMaintTicketPricing.Show
    frmMaintTicketPricing.Move 0, 0
    frmMaintTicketPricing.ZOrder
End Sub

Private Sub mnuTicketType_Click()
    frmMaintTicketType.Show
    frmMaintTicketType.Move 0, 0
    frmMaintTicketType.ZOrder
End Sub

Private Sub mnuToolsAuditTrail_Click()
    frmToolAuditTrail.Show
    frmToolAuditTrail.Move 0, 0
    frmToolAuditTrail.ZOrder
End Sub

Private Sub mnuToolsSetup_Click()
    frmHybridAutoNum.Show
    frmHybridAutoNum.Move 0, 0
    frmHybridAutoNum.ZOrder

'    frmSysStructure.Show
'    frmSysStructure.Move 0, 0
'    frmSysStructure.ZOrder
End Sub

Private Sub mnuMaintCompany_Click()
    frmMaintCompany.Show
    frmMaintCompany.Move 0, 0
    frmMaintCompany.ZOrder
End Sub

