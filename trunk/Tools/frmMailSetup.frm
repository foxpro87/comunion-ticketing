VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{B3FB64BF-91F9-11D7-A482-0008A14158BC}#2.22#0"; "ITGControls.ocx"
Begin VB.Form frmMailSetup 
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Email Setup"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7230
   Icon            =   "frmMailSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5535
   ScaleWidth      =   7230
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7935
      Top             =   3180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sbRS 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5280
      Width           =   7230
      _ExtentX        =   12753
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
   Begin ITGControls.ITGTab SSTab1 
      Height          =   5100
      Left            =   105
      TabIndex        =   1
      Top             =   105
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   8996
      TabCount        =   2
      TabCaption(0)   =   "              Main              "
      TabContCtrlCnt(0)=   1
      Tab(0)ContCtrlCap(1)=   "ComunionFrames2"
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
         Height          =   4650
         Left            =   -74955
         TabIndex        =   20
         Top             =   345
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   8202
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "cEmailSetupFor"
            Caption         =   "Setup Email For"
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
            DataField       =   "cFrom"
            Caption         =   "From"
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
            DataField       =   "cTo"
            Caption         =   "To"
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
               ColumnWidth     =   1349.858
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   2445.166
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2624.882
            EndProperty
         EndProperty
      End
      Begin ITGControls.ComunionFrames ComunionFrames2 
         Height          =   4620
         Left            =   90
         Top             =   360
         Width           =   6870
         _ExtentX        =   12118
         _ExtentY        =   8149
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
         Begin VB.TextBox txtDispBody 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   1725
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   4
            Top             =   2670
            Width           =   5055
         End
         Begin VB.ComboBox cboTime 
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
            ItemData        =   "frmMailSetup.frx":0CCA
            Left            =   5265
            List            =   "frmMailSetup.frx":0D58
            TabIndex        =   3
            Text            =   "cboTime"
            Top             =   1710
            Width           =   1410
         End
         Begin ITGControls.ITGXPOptionButton optTimeFixed 
            Height          =   285
            Index           =   0
            Left            =   4950
            TabIndex        =   2
            Top             =   1725
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            BackColor       =   16777215
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
         Begin ITGControls.ITGTextBox txtUserID 
            Height          =   285
            Left            =   120
            TabIndex        =   5
            Top             =   885
            Width           =   3660
            _ExtentX        =   6244
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
            Label           =   "User Name"
            BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LabelWidth      =   1580
            TextBoxWidth    =   2020
         End
         Begin ITGControls.ITGTextBox txtPassword 
            Height          =   285
            Left            =   120
            TabIndex        =   6
            Top             =   1200
            Width           =   3660
            _ExtentX        =   6244
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
            LabelWidth      =   1580
            TextBoxWidth    =   2020
         End
         Begin ITGControls.ITGTextBox txtSMTP 
            Height          =   285
            Left            =   3975
            TabIndex        =   7
            Top             =   885
            Width           =   2700
            _ExtentX        =   4551
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
            Text            =   "smtp.gmail.com"
            Mandatory       =   -1  'True
            Label           =   "SMTP"
            BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LabelWidth      =   1280
            TextBoxWidth    =   1360
         End
         Begin ITGControls.ITGTextBox txtPORT 
            Height          =   285
            Left            =   3975
            TabIndex        =   8
            Top             =   1200
            Width           =   2700
            _ExtentX        =   4551
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
            Text            =   "587"
            Mandatory       =   -1  'True
            Label           =   "Port Number"
            BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LabelWidth      =   1280
            TextBoxWidth    =   1360
         End
         Begin ITGControls.ITGTextBox txtFrom 
            Height          =   285
            Left            =   120
            TabIndex        =   9
            Top             =   1740
            Width           =   3660
            _ExtentX        =   6244
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
            Label           =   "Email From"
            BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LabelWidth      =   1580
            TextBoxWidth    =   2020
         End
         Begin ITGControls.ITGTextBox txtTo 
            Height          =   285
            Left            =   120
            TabIndex        =   10
            Top             =   2040
            Width           =   3660
            _ExtentX        =   6244
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
            Label           =   "Email To"
            BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LabelWidth      =   1580
            TextBoxWidth    =   2020
         End
         Begin ITGControls.ITGTextBox txtSubject 
            Height          =   285
            Left            =   120
            TabIndex        =   11
            Top             =   2340
            Width           =   3660
            _ExtentX        =   6244
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
            Label           =   "Subject"
            BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LabelWidth      =   1580
            TextBoxWidth    =   2020
         End
         Begin ITGControls.ITGTextBox txtBody 
            Height          =   285
            Left            =   135
            TabIndex        =   12
            Top             =   2670
            Width           =   3660
            _ExtentX        =   6244
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
            Label           =   "Body"
            BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LabelWidth      =   1580
            TextBoxWidth    =   2020
         End
         Begin ITGControls.ITGTextBox txtTranNo 
            Height          =   285
            Left            =   120
            TabIndex        =   13
            Top             =   570
            Width           =   3645
            _ExtentX        =   6218
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
            TextButton      =   -1  'True
            Mandatory       =   -1  'True
            Label           =   "Email Setup For"
            BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LabelWidth      =   1580
            TextBoxWidth    =   1720
            Hover           =   -1  'True
         End
         Begin ITGControls.ITGTextBox txtAttachment 
            Height          =   285
            Left            =   105
            TabIndex        =   14
            Top             =   4140
            Width           =   5985
            _ExtentX        =   10345
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
            Label           =   "Attachment"
            BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LabelWidth      =   1280
            TextBoxWidth    =   4645
         End
         Begin ITGControls.ComunionButton cmdAttachment 
            Height          =   285
            Left            =   6135
            TabIndex        =   15
            Top             =   4140
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   503
            BTYPE           =   3
            TX              =   "..."
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
            MICON           =   "frmMailSetup.frx":0F04
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ITGControls.ITGTextBox txtTime 
            Height          =   285
            Left            =   3960
            TabIndex        =   16
            Top             =   1710
            Width           =   2265
            _ExtentX        =   3784
            _ExtentY        =   503
            SendKeysTab     =   -1  'True
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
            Label           =   "Email Every"
            BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LabelWidth      =   1280
            TextBoxWidth    =   925
         End
         Begin ITGControls.ITGXPOptionButton optTimeFixed 
            Height          =   285
            Index           =   1
            Left            =   4950
            TabIndex        =   17
            Top             =   2055
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
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
         Begin ITGControls.ITGTextBox txtHour 
            Height          =   285
            Left            =   3960
            TabIndex        =   18
            Top             =   2055
            Width           =   2175
            _ExtentX        =   3625
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
            DecimalPlace    =   2
            Label           =   ""
            BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LabelWidth      =   1280
            TextBoxWidth    =   835
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hour/s"
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
            Left            =   6180
            TabIndex        =   19
            Top             =   2085
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frmMailSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'IT Group Inc. 2011.08.10

Option Explicit

'Object variables
Private oNavRec As clsNavRec
Private oFormSetup As New clsFormSetup
Private oRecordset As New clsRecordset
Private oBar As New clsToolBarMenuBit

'Recordset variables
Private WithEvents rsHeader As ADODB.Recordset
Attribute rsHeader.VB_VarHelpID = -1

'ADO Connection variables
Private oConnection As New clsConnection
Private connHeader As ADODB.Connection

'Form mode enumeration
Public Enum eMailMode
    Normal
    AddNewEdit
    Find
End Enum
Private Mode As eMailMode

'Other declaration
Public sBit As String
Private vBM As Variant 'Recordset bookmark variable

Private sTempPath As String ' Path of the Picture : Temporary
Private stm As New ADODB.Stream

'Security Acess Level variables
Public lACNew As Boolean
Public lACEdit As Boolean
Public lACDelete As Boolean
Public lACPost As Boolean
Public lACCancel As Boolean
Public lACPrint As Boolean
Private Function GetSearchString() As String
On Error GoTo ErrorHandler
Dim sWhere As String

    GetSearchString = True

    If Trim$(txtTranNo.Text) <> "" Then
        sWhere = sWhere & IIf(Trim$(sWhere) = "", "", " AND ") & " cEmailSetupFor LIKE '" & Trim$(txtTranNo.Text) & "%'"
    End If

    GetSearchString = Trim$(sWhere)

    Exit Function
ErrorHandler:
    GetSearchString = "ERROR"
End Function


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


'Delete record
Public Sub TBDelete()
On Error GoTo ErrorHandler
    
    If rsHeader.RecordCount = 0 Then Exit Sub
    If MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo, "ComUnion") = vbNo Then Exit Sub
    
    'Audit trail
    UpdateLogFile "Mail Setup", Trim(txtTranNo), "Deleted"

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


'Add new record to the recordset
Public Sub TBNew()
    
    oBar.BitEnabled Me, ITGLedgerMain.tbrMain, , , , , , , , True, True
    Mode = AddNewEdit
    
    If rsHeader.State <> adStateOpen Then
        oConnection.OpenNewConnection connHeader
        oRecordset.OpenRecordsetWithCN rsHeader, "*", "MAILSETUP", connHeader, "WHERE 1 = 0 ", True
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
    txtTranNo.SetFocus

    optTimeFixed(0).Enabled = True
    optTimeFixed(1).Enabled = True

End Sub


'Close active window
Public Sub TBCloseWindow()
    Unload Me
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
            oBar.BitEnabled Me, ITGLedgerMain.tbrMain, lACNew, , , , , , , , , True, , , True
            txtTranNo.Locked = False
            txtTranNo.SetFocus
            Exit Sub
        End If
        oRecordset.OpenRecordsetWithCN rsHeader, "*", "MailSetup", connHeader, IIf(sTemp = "", "", " WHERE " & sTemp)
                            
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
        Mode = Normal
        oBar.BitEnabled Me, ITGLedgerMain.tbrMain, lACNew, lACEdit, lACDelete, , , , , , , True, True, , True
        
        Unload frmWait
    Else
        
        RSZero
        oBar.BitEnabled Me, ITGLedgerMain.tbrMain, lACNew, , , , , , , , , True, , , True
        oFormSetup.FormSearch True
        Mode = Find
    End If
    
    optTimeFixed(0).Enabled = False
    optTimeFixed(1).Enabled = False

End Sub


'Sets the data source of the controls
Sub SetDataSource()
    Set FrmName = Me
    oRecordset.BindControls rsHeader
    Set dtgList.DataSource = rsHeader
End Sub
    
'Sets the data field for every bounded controls
Sub SetDataField()
    With rsHeader
        txtTranNo.DataField = !cEmailSetupFor
        txtUserID.DataField = !cUserID
        txtpassword.DataField = !cPassword
        txtSMTP.DataField = !cSMTP
        txtPORT.DataField = !cPORT
        
        txtFrom.DataField = !cFrom
        txtTo.DataField = !cTo
        txtSubject.DataField = !cSubject
        txtAttachment.DataField = !cAttachment
        
        txtBody.DataField = !cBody
        txtTime.DataField = !dTime
        
        txtHour.DataField = !dTimer
        
    End With
End Sub


'Sets the form if record number is zero
Private Sub RSZero()
    Set FrmName = Me
    oRecordset.UnbindControls
    oFormSetup.TextClearing
    oFormSetup.FormLocking True
    
    If rsHeader.State = adStateOpen Then rsHeader.Close
'    oBar.BitEnabled Me, ITGLedgerMain.tbrMain, , lACEdit, , , , , , , , , , , True
    Mode = Find
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
        oBar.BitEnabled Me, ITGLedgerMain.tbrMain, lACNew, lACEdit, lACDelete, , , , , , , True, True
    Else
        RSZero
        Mode = Find
    End If
    SetDataSource
    SetDataField
    SSTab1.ActiveTab = 0
    SSTab1.TabEnabled(1) = True
    

ErrorHandler:
    If err.Number = -2147217885 Then
        Resume Next
    ElseIf err.Number = -2147217842 Then 'Operation was cancelled. (Error returned by ITGDateBox)
        TBUndoAll
    Else
        oBar.BitEnabled Me, ITGLedgerMain.tbrMain, lACNew, lACEdit, lACDelete, , , , , , , True, True
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


'Check if all mandatory fields are complete
Function MandatoryOK() As Boolean
    MandatoryOK = True
       
    If Trim(txtUserID) = "" Then
        MandatoryOK = False
        MsgBox "Field 'User ID' is mandatory. Null value is not allowed.", vbInformation, "ComUnion"
        txtUserID.SetFocus
        Exit Function
    End If
            
    If Trim(txtpassword) = "" Then
        MandatoryOK = False
        MsgBox "Field 'Password' is mandatory. Null value is not allowed", vbInformation, "ComUnion"
        txtpassword.SetFocus
        Exit Function
    End If
    
    If Trim(txtSMTP) = "" Then
        MandatoryOK = False
        MsgBox "Field 'SMTP' is mandatory. Null value is not allowed", vbInformation, "ComUnion"
        txtSMTP.SetFocus
        Exit Function
    End If
    
    If Trim(txtPORT) = "" Then
        MandatoryOK = False
        MsgBox "Field 'PORT' is mandatory. Null value is not allowed", vbInformation, "ComUnion"
        txtPORT.SetFocus
        Exit Function
    End If
    
    If Trim(txtFrom) = "" Then
        MandatoryOK = False
        MsgBox "Field 'txtFrom' is mandatory. Null value is not allowed", vbInformation, "ComUnion"
        txtFrom.SetFocus
        Exit Function
    End If
    
    If Trim(txtTo) = "" Then
        MandatoryOK = False
        MsgBox "Field 'Email To' is mandatory. Null value is not allowed", vbInformation, "ComUnion"
        txtTo.SetFocus
        Exit Function
    End If
    
    If Trim(txtSubject) = "" Then
        MandatoryOK = False
        MsgBox "Field 'Subject is mandatory. Null value is not allowed", vbInformation, "ComUnion"
        txtSubject.SetFocus
        Exit Function
    End If
    
    If Trim(txtBody) = "" Then
        MandatoryOK = False
        MsgBox "Field 'Body' is mandatory. Null value is not allowed", vbInformation, "ComUnion"
        txtBody.SetFocus
        Exit Function
    End If

End Function


'Save all changes
Public Sub TBSave()
Dim OKUpdate As Boolean
Dim lNew As Boolean
Dim cmdDefault As ADODB.Command

On Error GoTo ErrHandler
    
    'Audit trail
    lBoolean = False
    lNew = False
    If rsHeader.Status = adRecNew Then
        lNew = True
        lBoolean = True
    End If
    
    If Not MandatoryOK Then Exit Sub
    
    OKUpdate = False
    cn.BeginTrans
    connHeader.BeginTrans
    
    rsHeader.UpdateBatch adAffectAll
       
    cn.CommitTrans
    connHeader.CommitTrans
    OKUpdate = True
    
    Set FrmName = Me
    oFormSetup.FormLocking True
    oBar.BitEnabled Me, ITGLedgerMain.tbrMain, lACNew, lACEdit, lACDelete, , , , , , , True, True, , True
    Mode = Normal
    
    MsgBox "Record/s successfully saved.", vbInformation, "ComUnion"
    SSTab1.ActiveTab = 0
    SSTab1.TabEnabled(1) = True
    
    InitEmailTime
    
    'Audit trail
    UpdateLogFile "Mail Setup", "Email", IIf(lBoolean, "Inserted", "Updated")
    
ErrHandler:
    If err.Number = -2147217885 Then
        Resume Next
    ElseIf err.Number = -2147217864 Then
        OKUpdate = True
        cn.RollbackTrans
        connHeader.RollbackTrans
        MsgBox "Record cannot be updated. Some values may have been changed by other user/s since last read." & vbCr & _
                "Records will be automatically refreshed. All changes made to the record will be gone upon refresh.", vbInformation, "ComUnion"
        vBookMark = rsHeader.Bookmark
        oRecordset.UnbindControls
        rsHeader.Requery
        oBar.BitEnabled Me, ITGLedgerMain.tbrMain, lACNew, lACEdit, lACDelete, , , , , , , True, True, , True
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
        connHeader.RollbackTrans
        ErrorLog err.Number, err.Description, Me.Name 'Error log
    End If
    
End Sub


'Sets the form & recorset to add/edit mode
Public Sub TBEdit()
    Mode = AddNewEdit
    oBar.BitEnabled Me, ITGLedgerMain.tbrMain, , , , , , , , True, True, , , , True
    Set FrmName = Me
    oFormSetup.FormLocking False
    oFormSetup.ClrRequired &HC0&
    SSTab1.ActiveTab = 0
    SSTab1.TabEnabled(1) = False
    vBM = rsHeader.Bookmark
    
    optTimeFixed(0).Enabled = True
    optTimeFixed(1).Enabled = True

    
End Sub


'Reload menu buttons (do not delete this sub)
Public Sub TBBitReload()
    oBar.BitVisible ITGLedgerMain.tbrMain
    ITGLedgerMain.tbrMain.Buttons("btnFind").ButtonMenus("btnFindP").Enabled = True
    oBar.BitReload Me, ITGLedgerMain.tbrMain, sBit
    Set FrmName = Me
End Sub

Private Sub ComunionButton1_Click()

End Sub


Private Sub cboTime_Change()
    cboTime_Click
End Sub

Private Sub cboTime_Click()
    txtTime = cboTime.Text
End Sub

Private Sub cmdAttachment_Click()
    With cd1
        .Filter = "All Files | *.*"
        .Flags = &H4 Or &H800 Or &H40000 Or &H200 Or &H80000
        .CancelError = True                             'Die if there are any errors
        .MaxFileSize = 30000                            'Just to make sure we have enough room
'        .ShowOpen
'        If cd1.FileName <> "" Then
'            txtAttachment = cd1.FileName
'        End If
        On Error Resume Next
        .ShowOpen                'Open
        Select Case err.Number
          Case cdlCancel
            'Cancel was selected
          Case Is <> 0
            'Some other error occurred
          Case 0
            'No error occured
            Dim sFile()  As String
            Dim i As Integer
            txtAttachment = ""
            sFile = Split(.FileName, ChrW$(0))               'Take apart null delimited list returned from multiselect CD
             If UBound(sFile) <> 0 Then
                For i = 0 To UBound(sFile)
                   txtAttachment = txtAttachment & sFile(i) & ";"
                Next i
                txtAttachment = Left(txtAttachment, Len(txtAttachment) - 1)
            Else    'for single file
                txtAttachment = .FileName
            End If
         End Select
    End With
    
End Sub

 
Private Sub Form_Activate()
    TBBitReload
End Sub

Private Sub Form_Load()
    oBar.AcessBit Me, GetValueFrTable("AccessLevel", "SEC_ACCESSLEVEL", "RoleID = '" & SecUserRole & "' AND [Module] = 'MAILSETUP'")
    
    Set rsHeader = New ADODB.Recordset
    Set oNavRec = New clsNavRec
    
    Set FrmName = Me
    oFormSetup.FormLocking True
    oFormSetup.FormSearch True
    oBar.BitEnabled Me, ITGLedgerMain.tbrMain, lACNew, , , , , , , , , True, , , True
    oBar.BitVisible ITGLedgerMain.tbrMain
    'ITGLedgerMain.tbrMain.Buttons("btnFind").ButtonMenus("btnFindP").Enabled = True
    If LoadOption("COM_THEME", 4) = "1 - Blue" Then
        oFormSetup.FormTheme (1)
    Else
        oFormSetup.FormTheme (2)
    End If
    
    txtTranNo.ConnectionStrings = "driver={" & sDBDriver & "};" & "server=" & sServer & ";uid=sa;pwd=" & sDBPassword & ";database=" & sDBname
    txtTranNo.SQLScript = "SELECT cReportHeader FROM sec_tran_report"
    
    Mode = Find
End Sub

Private Sub ITGTextBox1_LabelClick()

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

Private Sub optTimeFixed_Click(Index As Integer)
    If optTimeFixed(Index).Value = True Then
        If Index = 0 Then
            cboTime.Enabled = True
            txtHour.Enabled = False
            rsHeader!lTimer = False
        Else
            txtHour.Enabled = True
            cboTime.Enabled = False
            rsHeader!lTimer = True
        End If
    End If
End Sub

Private Sub txtBody_Change()
    txtDispBody.Text = txtBody
End Sub

Private Sub txtDispBody_Change()
    If Mode <> AddNewEdit Then Exit Sub
    txtBody = txtDispBody.Text
End Sub

Private Sub txtDispBody_LostFocus()
    On Error Resume Next
    rsHeader!cBody = txtDispBody.Text
End Sub

Private Sub rsHeader_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error GoTo ErrorHandler
    
    If Not (rsHeader.EOF) Or Not (rsHeader.BOF) Then
        sbRS.Panels(1) = "Record: " & IIf((rsHeader.AbsolutePosition = -2), "0", rsHeader.AbsolutePosition) & "/" & rsHeader.RecordCount
        
        If rsHeader.Status <> adRecNew Then
            txtTranNo.Locked = True
        Else
            txtTranNo.Locked = False
        End If
        
        txtDispBody.Text = txtBody
        
        If rsHeader!lTimer = True Then
            optTimeFixed(1).Value = True
            optTimeFixed(0).Value = False
        Else
            optTimeFixed(0).Value = True
            optTimeFixed(1).Value = False
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
        txtTranNo.Locked = False
    End If
    
ErrorHandler:
    'Err.Number -2147217885
    'Description - Row handle referred to a deleted row or a row marked for deletion.
    If err.Number = -2147217885 Then
        Resume Next
    End If
    
End Sub

Private Sub txtTime_Change()
    cboTime.Text = txtTime
End Sub

Private Sub txtTranNo_LabelClick()
    frmDatePicker.ReportManagement = True
    frmDatePicker.Show
    frmDatePicker.ShowForm txtTranNo
    frmDatePicker.Move (Me.Left + Me.Width), Me.Top
End Sub
