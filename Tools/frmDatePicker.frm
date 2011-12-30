VERSION 5.00
Object = "{45D48925-E17D-4DB1-BE71-CD38EC107318}#1.0#0"; "SendEmail.tlb"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B3FB64BF-91F9-11D7-A482-0008A14158BC}#2.22#0"; "ITGControls.ocx"
Begin VB.Form frmDatePicker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Criteria"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7695
   Icon            =   "frmDatePicker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5625
   ScaleWidth      =   7695
   Begin SendEmailCtl.SendEmail SendEmail1 
      Height          =   330
      Left            =   105
      TabIndex        =   8
      Top             =   1860
      Visible         =   0   'False
      Width           =   360
      Object.Visible         =   "True"
      Enabled         =   "True"
      ForegroundColor =   "-2147483630"
      BackgroundColor =   "-2147483633"
      EmailUsername   =   ""
      EmailPassword   =   ""
      EmailSMTP       =   ""
      EmailPORT       =   "0"
      EmailFrom       =   ""
      EmailTo         =   ""
      EmailSubject    =   ""
      EmailBody       =   ""
      EmailAttachment =   ""
      BackColor       =   "Control"
      BackgroundImageLayout=   "Center"
      ForeColor       =   "ControlText"
      Location        =   "7, 124"
      Name            =   "SendEmail"
      Size            =   "24, 22"
      Object.TabIndex        =   "0"
   End
   Begin ITGControls.ITGTab SSTab1 
      Height          =   3870
      Left            =   15
      TabIndex        =   9
      Top             =   45
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   6826
      TabCount        =   2
      TabCaption(0)   =   "              Main              "
      TabContCtrlCnt(0)=   1
      Tab(0)ContCtrlCap(1)=   "RptFrames"
      TabCaption(1)   =   "               List               "
      TabContCtrlCnt(1)=   1
      Tab(1)ContCtrlCap(1)=   "dtgList"
      TabStyle        =   1
      TabTheme        =   1
      ActiveTabBackStartColor=   10874879
      ActiveTabBackEndColor=   16514555
      InActiveTabBackStartColor=   12648447
      InActiveTabBackEndColor=   15397104
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OuterBorderColor=   10198161
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   10526880
      Begin MSDataGridLib.DataGrid dtgList 
         Height          =   3390
         Left            =   -74940
         TabIndex        =   18
         Top             =   390
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   5980
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         Appearance      =   0
         BackColor       =   12648447
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   2
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "cCode"
            Caption         =   "Code"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "cReportHeader"
            Caption         =   "CR Header"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "cTranHeader"
            Caption         =   "SP Header"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "cParam"
            Caption         =   "Parameters"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
               ColumnAllowSizing=   -1  'True
               Locked          =   -1  'True
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   1814.74
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1830.047
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2789.858
            EndProperty
         EndProperty
      End
      Begin ITGControls.ComunionFrames RptFrames 
         Height          =   3390
         Left            =   75
         Top             =   405
         Width           =   5310
         _ExtentX        =   9366
         _ExtentY        =   5980
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
         Begin ITGControls.ITGCheckBox chkSP 
            Height          =   300
            Left            =   165
            TabIndex        =   10
            Top             =   555
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   529
            BackColor       =   16777215
            Caption         =   "Stored Procedure"
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
         Begin ITGControls.ITGTextBox txtcCode 
            Height          =   285
            Left            =   120
            TabIndex        =   11
            Top             =   930
            Width           =   3300
            _ExtentX        =   5609
            _ExtentY        =   503
            SendKeysTab     =   -1  'True
            BackColor       =   14745599
            LabelBackColor  =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mandatory       =   -1  'True
            Label           =   "Report Code"
            BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LabelWidth      =   1700
            TextBoxWidth    =   1540
         End
         Begin ITGControls.ITGTextBox txtCRHeader 
            Height          =   285
            Left            =   120
            TabIndex        =   12
            Top             =   1245
            Width           =   3300
            _ExtentX        =   5609
            _ExtentY        =   503
            SendKeysTab     =   -1  'True
            BackColor       =   14745599
            LabelBackColor  =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mandatory       =   -1  'True
            Label           =   "Report Header Name"
            BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LabelWidth      =   1700
            TextBoxWidth    =   1540
         End
         Begin ITGControls.ITGTextBox txtSPHeader 
            Height          =   285
            Left            =   120
            TabIndex        =   13
            Top             =   1575
            Width           =   3285
            _ExtentX        =   5583
            _ExtentY        =   503
            SendKeysTab     =   -1  'True
            BackColor       =   14745599
            LabelBackColor  =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mandatory       =   -1  'True
            Label           =   "Header Stored Proc"
            BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LabelWidth      =   1700
            TextBoxWidth    =   1525
         End
         Begin ITGControls.ITGTextBox txtCRSubReport 
            Height          =   285
            Left            =   135
            TabIndex        =   14
            Top             =   2340
            Width           =   5040
            _ExtentX        =   8678
            _ExtentY        =   503
            SendKeysTab     =   -1  'True
            BackColor       =   14745599
            LabelBackColor  =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mandatory       =   -1  'True
            Label           =   "CR Sub Report Name/s"
            BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LabelWidth      =   1700
            TextBoxWidth    =   3280
         End
         Begin ITGControls.ITGTextBox txtNoSub 
            Height          =   285
            Left            =   120
            TabIndex        =   15
            Top             =   1950
            Width           =   3285
            _ExtentX        =   5583
            _ExtentY        =   503
            SendKeysTab     =   -1  'True
            BackColor       =   14745599
            LabelBackColor  =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
            DataType        =   1
            Mandatory       =   -1  'True
            Label           =   "No. of Sub Reports"
            BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LabelWidth      =   1700
            TextBoxWidth    =   1525
         End
         Begin ITGControls.ITGTextBox txtSPSubReport 
            Height          =   285
            Left            =   135
            TabIndex        =   16
            Top             =   2655
            Width           =   5040
            _ExtentX        =   8678
            _ExtentY        =   503
            SendKeysTab     =   -1  'True
            BackColor       =   14745599
            LabelBackColor  =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mandatory       =   -1  'True
            Label           =   "SP Sub Report Name/s"
            BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LabelWidth      =   1700
            TextBoxWidth    =   3280
         End
         Begin ITGControls.ITGTextBox txtParameters 
            Height          =   285
            Left            =   135
            TabIndex        =   17
            Top             =   2970
            Width           =   5040
            _ExtentX        =   8678
            _ExtentY        =   503
            SendKeysTab     =   -1  'True
            BackColor       =   14745599
            LabelBackColor  =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mandatory       =   -1  'True
            Label           =   "Parameters"
            BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LabelWidth      =   1700
            TextBoxWidth    =   3280
         End
      End
   End
   Begin VB.ComboBox cboReport 
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
      Left            =   1215
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   630
      Width           =   2280
   End
   Begin MSComCtl2.DTPicker dDate 
      Height          =   315
      Left            =   1215
      TabIndex        =   0
      Top             =   225
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   42729473
      CurrentDate     =   40760
   End
   Begin ITGControls.ComunionButton cmdOK 
      Height          =   330
      Left            =   1680
      TabIndex        =   3
      Top             =   1230
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   582
      BTYPE           =   3
      TX              =   "&OK"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      MICON           =   "frmDatePicker.frx":0CCA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ITGControls.ComunionButton cmdCancel 
      Height          =   330
      Left            =   2640
      TabIndex        =   4
      Top             =   1230
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   582
      BTYPE           =   3
      TX              =   "&Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      MICON           =   "frmDatePicker.frx":0CE6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.StatusBar sbRS 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   5370
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin ITGControls.ComunionButton cmdEmail 
      Height          =   330
      Left            =   90
      TabIndex        =   7
      Top             =   1230
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   582
      BTYPE           =   3
      TX              =   "Email"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      MICON           =   "frmDatePicker.frx":0D02
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   660
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   285
      Width           =   345
   End
End
Attribute VB_Name = "frmDatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'IT Group Inc. 2005.09.23

Option Explicit

'Object variables
Private oNavRec As clsNavRec
Private oFormSetup As New clsFormSetup
Private oRecordset As New clsRecordset
Private oBar As New clsToolBarMenuBit

'Recordset variables
Private WithEvents rsHeader As ADODB.Recordset
Attribute rsHeader.VB_VarHelpID = -1

'Form mode enumeration
Public Enum eMaintReport
    Normal
    AddNewEdit
    Find
End Enum
Private Mode As eMaintReport

'Other declaration
Public sBit As String
Private vBM As Variant 'Recordset bookmark variable

Private connHeader As ADODB.Connection
Private oConnection As New clsConnection

'Security Acess Level variables
Public lACNew As Boolean
Public lACEdit As Boolean
Public lACDelete As Boolean
Public lACPost As Boolean
Public lACCancel As Boolean
Public lACPrint As Boolean


Public ReportManagement As Boolean

Public lCancel As Boolean

'Add new record to the recordset
Public Sub TBNew()
    
    oBar.BitEnabled Me, ITGLedgerMain.tbrMain, , , , True, True, , , True, True, , , , True
    txtcCode.BackColor = &HE0FFFF
    Mode = AddNewEdit
    
    If rsHeader.State <> adStateOpen Then
        oRecordset.OpenRecordset rsHeader, "*", "SEC_TRAN_REPORT", "WHERE 1 = 0 ", True
        SetDataSource
        SetDataField
    Else
        vBM = rsHeader.Bookmark
    End If
    
    rsHeader.AddNew
    rsHeader!cCompanyID = COID
    
    Set FrmName = Me
    oFormSetup.FormLocking False
    oFormSetup.ClrRequired &HC0&
    SSTab1.ActiveTab = 0
    SSTab1.TabEnabled(1) = False
    txtcCode.SetFocus

End Sub

'Save all changes
Public Sub TBSave()
Dim OKUpdate As Boolean
On Error GoTo ErrHandler
    
    'Audit Trail
    lBoolean = False
    If rsHeader.Status = adRecNew Then lBoolean = True
    
    If Not MandatoryOK Then Exit Sub
    
    OKUpdate = False
    cn.BeginTrans
    
    rsHeader.UpdateBatch adAffectAll
    
    cn.CommitTrans
    OKUpdate = True
    
    Set FrmName = Me
    oFormSetup.FormLocking True
    oBar.BitEnabled Me, ITGLedgerMain.tbrMain, lACNew, lACEdit, lACDelete, , , , , , , True, True, , True
    Mode = Normal
    
    MsgBox "Record/s successfully saved.", vbInformation, "ComUnion"
    sbRS.Panels(2) = ""
    
    SSTab1.TabEnabled(1) = True
    
    'Audit trail
    UpdateLogFile "Maintenance Report", Trim(txtcCode), IIf(lBoolean, "Inserted", "Updated")

ErrHandler:
    If err.Number = -2147217885 Then
        Resume Next
    ElseIf err.Number = -2147217864 Then
        OKUpdate = True
        cn.RollbackTrans
        MsgBox "Record cannot be updated. Some values may have been changed by other user/s since last read." & vbCr & _
                "Records will be automatically refreshed. All changes made to the record will be gone upon refresh.", vbInformation, "ComUnion"
        vBookMark = rsHeader.Bookmark
        oRecordset.UnbindControls
        rsHeader.Requery
        oBar.BitEnabled Me, ITGLedgerMain.tbrMain, lACNew, lACEdit, lACDelete, , , , , , , , True, True, , True
        Mode = Normal
        If rsHeader.RecordCount <> 0 Then
            Set FrmName = Me
            oFormSetup.FormLocking True
            SetDataField
            SetDataSource
            rsHeader.Bookmark = vBookMark
        Else
            RSZero
        End If
    End If
    
    If Not OKUpdate Then
        MsgBox "Transaction update failed.", vbInformation, "ComUnion"
        cn.RollbackTrans
        ErrorLog err.Number, err.Description, Me.Name 'Error log
    End If
    
End Sub

'Sets the form & recorset to add/edit mode
Public Sub TBEdit()
    Mode = AddNewEdit
    oBar.BitEnabled Me, ITGLedgerMain.tbrMain, , , , True, True, , , True, True, , , , True
    Set FrmName = Me
    oFormSetup.FormLocking False
    oFormSetup.ClrRequired &HC0&
    txtcCode.Locked = True
    SSTab1.ActiveTab = 0
    SSTab1.TabEnabled(1) = False
    
    txtcCode.SetFocus
    vBM = rsHeader.Bookmark
End Sub

'Delete record
Public Sub TBDelete()
On Error GoTo ErrorHandler
    
    If rsHeader.RecordCount = 0 Then Exit Sub
    
    If MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo, "ComUnion") = vbNo Then Exit Sub
    
    'Audit trail
    UpdateLogFile "Maintenance Report", Trim(txtcCode), "Deleted"

    rsHeader.Delete adAffectCurrent
    rsHeader.UpdateBatch adAffectAll
    
    TBPrevRec

    Mode = Normal

    If rsHeader.RecordCount = 0 Then
        RSZero
    End If

ErrorHandler:
    If err.Number = -2147217885 Then
        Resume Next
    End If

End Sub

'Search using the frmITGSearch
Public Sub TBFind()
    Mode = Normal
    txtcCode.Locked = True

End Sub

'Search using the recordset primary key
Public Sub TBFindPrimary()
Dim sTemp As String
If Mode = Find Then
    FormWaitShow "Loading data . . ."
            
    oConnection.OpenNewConnection connHeader
    
    Set rsHeader = Nothing
    Set rsHeader = New ADODB.Recordset
    
    sTemp = Trim$(GetSearchString)
    If sTemp = "ERROR" Then
        MsgBox "Only Allows A - Z, 0 - 9, '.', ',' and %(wildcard)", vbExclamation, ""
        oBar.BitEnabled Me, ITGLedgerMain.tbrMain, lACNew, lACEdit, lACDelete, , , , , , , True, , , True
        txtcCode.Locked = False
        txtcCode.SetFocus
        Exit Sub
    End If
                
    oRecordset.OpenRecordsetWithCN rsHeader, "*", "SEC_TRAN_REPORT", connHeader, IIf(sTemp = "", "", " WHERE " & sTemp), True
                        
    Set FrmName = Me
    oFormSetup.FormLocking True
            
    If rsHeader.RecordCount = 0 Then
        Unload frmWait
        MsgBox "No matching record/s found.", vbInformation, "ComUnion Search"
        RSZero
        oFormSetup.FormSearch True
        Mode = Find
        Exit Sub
    End If
    SetDataSource
    SetDataField
    txtcCode.BackColor = &HE0FFFF
    Mode = Normal
    oBar.BitEnabled Me, ITGLedgerMain.tbrMain, lACNew, lACEdit, lACDelete, , , , , , , True, True, , True

    Unload frmWait
Else
    RSZero
    oFormSetup.FormSearch True
    Mode = Find
End If

End Sub

'Reload menu buttons (do not delete this sub)
Public Sub TBBitReload()
    oBar.BitVisible ITGLedgerMain.tbrMain
    ITGLedgerMain.tbrMain.Buttons("btnFind").ButtonMenus("btnFindP").Enabled = True
    oBar.BitReload Me, ITGLedgerMain.tbrMain, sBit
    Set FrmName = Me
End Sub

'Close active window
Public Sub TBCloseWindow()
    Unload Me
End Sub

'Move first
Public Sub TBFirstRec()
    If rsHeader.State <> adStateOpen Then Exit Sub
    oNavRec.MoveFirst rsHeader
End Sub

'Move previous
Public Sub TBPrevRec()
    If rsHeader.State <> adStateOpen Then Exit Sub
    oNavRec.MovePrevious rsHeader
End Sub

'Move next
Public Sub TBNextRec()
    If rsHeader.State <> adStateOpen Then Exit Sub
    oNavRec.MoveNext rsHeader
End Sub

'Move last
Public Sub TBLastRec()
    If rsHeader.State <> adStateOpen Then Exit Sub
    oNavRec.MoveLast rsHeader
End Sub

Private Function GetSearchString() As String
On Error GoTo ErrorHandler
Dim sWhere As String

    GetSearchString = True

    If Trim$(txtcCode.Text) <> "" Then
        sWhere = sWhere & IIf(Trim$(sWhere) = "", "", " AND ") & " cCode LIKE '" & Trim$(txtcCode.Text) & "%'"
    End If
    GetSearchString = Trim$(sWhere)
    Exit Function
ErrorHandler:
    GetSearchString = "ERROR"
End Function

'Check if all mandatory fields are complete
Function MandatoryOK() As Boolean
    
    MandatoryOK = True
    
    If Trim(txtcCode) = "" Then
        MandatoryOK = False
        MsgBox "Report Code' is mandatory. Null value is not allowed.", vbInformation, "ComUnion"
        txtcCode.SetFocus
        Exit Function
    End If
    
    If rsHeader.Status = adRecNew Then
        If IDExisting(rsHeader, "cCode", "SEC_TRAN_REPORT", Trim(rsHeader!cCode), , True) Then
            MandatoryOK = False
            MsgBox "Report Code already exists.", vbInformation, "ComUnion"
            txtcCode.SetFocus
            Exit Function
        End If
    End If
    
    If Trim(txtCRHeader) = "" Then
        MandatoryOK = False
        MsgBox "Report Header' is mandatory. Null value is not allowed", vbInformation, "ComUnion"
        txtCRHeader.SetFocus
        Exit Function
    End If

    If Trim(txtSPHeader) = "" Then
        MandatoryOK = False
        MsgBox "Stored Procedure Header' is mandatory. Null value is not allowed", vbInformation, "ComUnion"
        txtSPHeader.SetFocus
        Exit Function
    End If

End Function

'Sets the form if record number is zero
Private Sub RSZero()
    sbRS.Panels(1) = "Record: 0/0"
    sbRS.Panels(2) = ""
    
    Set FrmName = Me
    oRecordset.UnbindControls
    oFormSetup.TextClearing
    oFormSetup.FormLocking True
    
    If rsHeader.State = adStateOpen Then rsHeader.Close
    
    oBar.BitEnabled Me, ITGLedgerMain.tbrMain, lACNew, lACEdit, lACDelete, , , , , , , True, , , True
    
    txtcCode.Locked = False
    txtcCode.SetFocus
    
    Mode = Find
End Sub

'Sets the data field for every bounded controls
Sub SetDataField()
    With rsHeader
        chkSP.DataField = !lIsSP
        txtNoSub.DataField = !nSubRptNo
        txtcCode.DataField = !cCode
        txtCRHeader.DataField = !cReportHeader
        txtSPHeader.DataField = !cTranHeader
        txtCRSubReport.DataField = !cReportDetail
        txtSPSubReport.DataField = !cTranDetail
        txtParameters.DataField = !cParam
    End With
End Sub

'Sets the data source of the controls
Sub SetDataSource()
    Set FrmName = Me
    oRecordset.BindControls rsHeader
    Set dtgList.DataSource = rsHeader
    'Set chkSP.DataSource = rsHeader
End Sub

Private Sub cmdEmail_Click()
    'We Will Now send the email
    If rs.State = 1 Then rs.Close
    rs.Open "select * from MailSetup where cEmailSetupFor = '" & cboReport.Text & "' and cCompanyID = '" & COID & "'", cn, adOpenStatic, adLockReadOnly
    
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

End Sub

Private Sub Form_Activate()
    TBBitReload
End Sub

Private Sub rsHeader_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo ErrorHandler
    
    If Not (rsHeader.EOF) Or Not (rsHeader.BOF) Then
        sbRS.Panels(1) = "Record: " & IIf((rsHeader.AbsolutePosition = -2), "0", rsHeader.AbsolutePosition) & "/" & rsHeader.RecordCount
        
        If rsHeader.Status <> adRecNew Then
            txtcCode.Locked = True
        Else
            txtcCode.Locked = False
        End If
        
        If Mode = AddNewEdit Then
            Select Case rsHeader.Status
                Case adRecNew
                    sbRS.Panels(2) = "New"
                Case adRecModified
                    sbRS.Panels(2) = "Modified"
                Case Else
                    sbRS.Panels(2) = ""
            End Select
        Else
            sbRS.Panels(2) = ""
        End If
    Else
        sbRS.Panels(1) = "Record: 0/0"
        sbRS.Panels(2) = ""
        txtcCode.Locked = False
    End If
    
ErrorHandler:
    'Err.Number -2147217885
    'Description - Row handle referred to a deleted row or a row marked for deletion.
    If err.Number = -2147217885 Then
        Resume Next
    End If
    
End Sub

Private Sub dtgList_HeadClick(ByVal ColIndex As Integer)
    SortGrid dtgList, ColIndex, rsHeader
End Sub

Public Sub ShowForm(sTranNo As String)
On Error GoTo ErrorHandler
    If Mode = Find Then

        oConnection.OpenNewConnection connHeader

        Set rsHeader = Nothing
        Set rsHeader = New ADODB.Recordset

        oRecordset.OpenRecordsetWithCN rsHeader, "*", "SEC_TRAN_REPORT", connHeader, "WHERE cReportHeader = '" & Trim(sTranNo) & "'"

        Set FrmName = Me
        oFormSetup.FormLocking True

        If rsHeader.RecordCount = 0 Then
            Unload frmWait
            MsgBox "No matching record/s found.", vbInformation, "ComUnion Search"
            RSZero
            oFormSetup.FormSearch True
            Mode = Find
            Exit Sub
        End If
            SetDataSource
            SetDataField
            txtcCode.BackColor = &HE0FFFF
            Mode = Normal
            oBar.BitEnabled Me, ITGLedgerMain.tbrMain, lACNew, lACEdit, lACDelete, , , , , , , True, True, , True

    End If
ErrorHandler:
    If err.Number <> 0 Then
        MsgBox err.Description, vbInformation, "ComUnion"
    End If
End Sub



Private Sub cmdOK_Click()
    Dim oRPT As New clsPrinting
    FormWaitShow "Loading Report. . ."
    With oRPT
        .pModule = cboReport.Text
        .pCOID = COID
        .sDateTran = frmDatePicker.dDate
        .RPTPath = App.path & "\Reports"
        Set .pCN = cn
        .PrintReceipt
    End With
    Set oRPT = Nothing
    Unload frmWait
    
    lCancel = False
End Sub

Private Sub cmdcancel_Click()
    lCancel = True
    Unload Me
End Sub

Private Sub Form_Load()
    If ReportManagement = False Then
        dDate.Value = Format(Now, "MM/DD/YYYY")
        LoadComboValues cboReport, "cReportHeader", "sec_tran_report"
        
        Me.Width = 3780
        Me.Height = 2385
        
        cmdEmail.Enabled = True
        cmdOK.Enabled = True
        cmdCancel.Enabled = True
        dDate.Enabled = True
        cboReport.Enabled = True
        
        SSTab1.Visible = False
        
    Else
        oBar.AcessBit Me, GetValueFrTable("AccessLevel", "SEC_ACCESSLEVEL", "RoleID = '" & SecUserRole & "' AND [Module] = 'MTPRICE'")
    
        Set rsHeader = New ADODB.Recordset
        Set oNavRec = New clsNavRec
        
        Set FrmName = Me
        oFormSetup.FormLocking True
        oFormSetup.FormSearch True
        oBar.BitEnabled Me, ITGLedgerMain.tbrMain, lACNew, lACEdit, lACDelete, , , , , , , , True, , , True
        oBar.BitVisible ITGLedgerMain.tbrMain
        ITGLedgerMain.tbrMain.Buttons("btnFind").ButtonMenus("btnFindP").Enabled = True
    
        Mode = Find
        txtcCode.Locked = False
        
        If LoadOption("COM_THEME", 4) = "1 - Blue" Then
            oFormSetup.FormTheme (1)
        Else
            oFormSetup.FormTheme (2)
        End If
        
        
        Me.Width = 5625
        Me.Height = 4665
        
        SSTab1.Visible = True
        
        cmdEmail.Enabled = False
        cmdOK.Enabled = False
        cmdCancel.Enabled = False
        dDate.Enabled = False
        cboReport.Enabled = False
        
        cmdEmail.Visible = False
        cmdOK.Visible = False
        cmdCancel.Visible = False
        dDate.Visible = False
        cboReport.Visible = False
        
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If Mode = AddNewEdit Then
        MsgBox "Unable to close. You are in Add/New/Edit mode." & vbCr & _
            " Must Save or Undo", vbCritical, Me.Caption
        Cancel = True
        Exit Sub
    End If
    
    oBar.BitEnabled Me, ITGLedgerMain.tbrMain, , , , , , , , , , True
    oBar.BitVisible ITGLedgerMain.tbrMain
    ITGLedgerMain.tbrMain.Buttons("btnFind").ButtonMenus("btnFindP").Enabled = False
    
    Set oNavRec = Nothing
    Set oFormSetup = Nothing
    Set oRecordset = Nothing
    Set oBar = Nothing
    Set rsHeader = Nothing
    
    lCloseWindow = True

End Sub

'Undo all changes to the recordset
Public Sub TBUndoAll()
On Error GoTo ErrorHandler
    
    Mode = Normal
    
    If rsHeader.Status = adRecNew Then TBUndoCurrent
    rsHeader.CancelBatch adAffectAll
    oRecordset.UnbindControls
    If rsHeader.RecordCount <> 0 Then rsHeader.Bookmark = vBM
    
    Set FrmName = Me
    oFormSetup.FormLocking True
    
    If rsHeader.RecordCount <> 0 Then
        oBar.BitEnabled Me, ITGLedgerMain.tbrMain, lACNew, lACEdit, lACDelete, , , , , , , , True, True, , True
    Else
        RSZero
        Mode = Find
    End If
    
    sbRS.Panels(2) = ""
    
    SetDataSource
    SetDataField

ErrorHandler:
    If err.Number = -2147217885 Then
        Resume Next
    ElseIf err.Number = -2147217842 Then 'Operation was cancelled. (Error returned by ITGDateBox)
        TBUndoAll
    End If

End Sub

'Undo changes on the current record
Public Sub TBUndoCurrent()
On Error GoTo ErrorHandler

    If rsHeader.Status = adRecNew Then
        rsHeader.CancelUpdate
    Else
        rsHeader.CancelBatch adAffectCurrent
    End If
    If rsHeader.RecordCount = 0 Then RSZero
    
ErrorHandler:
    If err.Number = -2147217885 Then
        Resume Next
    ElseIf err.Number = -2147217842 Then 'Operation was cancelled. (Error returned by ITGDateBox)
        TBUndoCurrent
    End If

End Sub

