VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{B3FB64BF-91F9-11D7-A482-0008A14158BC}#2.22#0"; "itgcontrols.ocx"
Begin VB.Form frmHybridAutoNum 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction Number Setup"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10905
   Icon            =   "frmHybridAutoNum.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4860
   ScaleWidth      =   10905
   Begin VB.ComboBox cboType 
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
      Left            =   8640
      TabIndex        =   29
      Text            =   "Combo1"
      Top             =   3450
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tools"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   30
      TabIndex        =   18
      Top             =   3060
      Width           =   4710
      Begin VB.CheckBox chkApplyVisible 
         Alignment       =   1  'Right Justify
         Caption         =   "Visible All"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   190
         Left            =   2745
         TabIndex        =   22
         Top             =   630
         Width           =   1230
      End
      Begin VB.CheckBox chkApplyBold 
         Alignment       =   1  'Right Justify
         Caption         =   "Bold All"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   190
         Left            =   2745
         TabIndex        =   21
         Top             =   330
         Width           =   1230
      End
      Begin VB.CheckBox chkApplyLocked 
         Alignment       =   1  'Right Justify
         Caption         =   "Locked All"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   190
         Left            =   270
         TabIndex        =   20
         Top             =   615
         Width           =   1230
      End
      Begin VB.CheckBox chkApplyInactive 
         Alignment       =   1  'Right Justify
         Caption         =   "Inactive All"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   190
         Left            =   270
         TabIndex        =   19
         Top             =   330
         Width           =   1230
      End
      Begin ITGControls.ComunionButton cmdApplyBold 
         Height          =   285
         Left            =   4050
         TabIndex        =   23
         Top             =   300
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "Apply"
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
         MICON           =   "frmHybridAutoNum.frx":0CCA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ITGControls.ITGTextBox txtUniFormat 
         Height          =   285
         Left            =   270
         TabIndex        =   24
         Top             =   915
         Width           =   2490
         _ExtentX        =   4180
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
         Label           =   "Universal Format"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelWidth      =   1300
         TextBoxWidth    =   1130
      End
      Begin ITGControls.ComunionButton cmdApplyInactive 
         Height          =   285
         Left            =   1605
         TabIndex        =   25
         Top             =   285
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "Apply"
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
         MICON           =   "frmHybridAutoNum.frx":0CE6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ITGControls.ComunionButton cmdApplyLocked 
         Height          =   285
         Left            =   1590
         TabIndex        =   26
         Top             =   570
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "Apply"
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
         MICON           =   "frmHybridAutoNum.frx":0D02
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ITGControls.ComunionButton cmdApplyVisible 
         Height          =   285
         Left            =   4050
         TabIndex        =   27
         Top             =   585
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "Apply"
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
         MICON           =   "frmHybridAutoNum.frx":0D1E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ITGControls.ComunionButton cmdApplyFormat 
         Height          =   285
         Left            =   2820
         TabIndex        =   28
         Top             =   930
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "Apply"
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
         MICON           =   "frmHybridAutoNum.frx":0D3A
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
   Begin VB.CheckBox chkVisible 
      Height          =   190
      Index           =   0
      Left            =   4890
      TabIndex        =   15
      Top             =   4170
      Width           =   190
   End
   Begin VB.CheckBox chkBold 
      Height          =   190
      Index           =   0
      Left            =   5145
      TabIndex        =   14
      Top             =   4170
      Width           =   190
   End
   Begin VB.CheckBox chkLocked 
      Height          =   190
      Index           =   0
      Left            =   5400
      TabIndex        =   13
      Top             =   4170
      Width           =   190
   End
   Begin VB.CheckBox chkInactive 
      Height          =   190
      Index           =   0
      Left            =   5640
      TabIndex        =   12
      Top             =   4170
      Width           =   190
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filters"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   4800
      TabIndex        =   5
      Top             =   3060
      Width           =   3240
      Begin VB.ComboBox cboOtherFilter 
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
         ItemData        =   "frmHybridAutoNum.frx":0D56
         Left            =   1080
         List            =   "frmHybridAutoNum.frx":0D63
         TabIndex        =   10
         Text            =   "cboOtherFilter"
         Top             =   600
         Width           =   1260
      End
      Begin VB.ComboBox cboModule 
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
         ItemData        =   "frmHybridAutoNum.frx":0D83
         Left            =   1095
         List            =   "frmHybridAutoNum.frx":0D8D
         TabIndex        =   7
         Text            =   "cboModule"
         Top             =   240
         Width           =   1260
      End
      Begin ITGControls.ITGTextBox txtModule 
         Height          =   285
         Left            =   195
         TabIndex        =   6
         Top             =   270
         Width           =   2025
         _ExtentX        =   3360
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
         Label           =   "Module ID"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextBoxWidth    =   465
      End
      Begin ITGControls.ComunionButton cmdView 
         Height          =   285
         Left            =   2415
         TabIndex        =   8
         Top             =   285
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "View"
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
         MICON           =   "frmHybridAutoNum.frx":0D9E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ITGControls.ITGTextBox txtOtherFilter 
         Height          =   285
         Left            =   180
         TabIndex        =   9
         Top             =   615
         Width           =   2880
         _ExtentX        =   4868
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
         Label           =   "Other Filter"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelWidth      =   2200
         TextBoxWidth    =   620
      End
   End
   Begin VB.Frame Frame3 
      Height          =   990
      Left            =   75
      TabIndex        =   1
      Top             =   0
      Width           =   9390
      Begin VB.Label lblDetails 
         Caption         =   "Format"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   255
         TabIndex        =   3
         Top             =   600
         Width           =   6030
      End
      Begin VB.Label lblHeader 
         Caption         =   "ModuleID - Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   9045
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   7485
      Top             =   4170
   End
   Begin MSComctlLib.StatusBar sbRS 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4605
      Width           =   10905
      _ExtentX        =   19235
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
   Begin MSDataGridLib.DataGrid dtgAutoNum 
      Height          =   1560
      Left            =   105
      TabIndex        =   11
      Top             =   1035
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   2752
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "lInactive"
         Caption         =   "Inactive"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   4
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   8
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "cCompanyID"
         Caption         =   "Company ID"
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
         DataField       =   "cModuleID"
         Caption         =   "Module ID"
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
         DataField       =   "cDesc"
         Caption         =   "Desc"
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
      BeginProperty Column04 
         DataField       =   "cTableName"
         Caption         =   "Table Name"
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
      BeginProperty Column05 
         DataField       =   "cType"
         Caption         =   "Type"
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
      BeginProperty Column06 
         DataField       =   "cFormat"
         Caption         =   "Format"
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
      BeginProperty Column07 
         DataField       =   "cPrefix"
         Caption         =   "Prefix"
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
      BeginProperty Column08 
         DataField       =   "cSuffix"
         Caption         =   "Suffix"
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
      BeginProperty Column09 
         DataField       =   "nValue"
         Caption         =   "Value"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "lLocked"
         Caption         =   "Locked"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   4
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   8
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "lBold"
         Caption         =   "Bold"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   4
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   8
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "lVisible"
         Caption         =   "Visible"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   4
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   8
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   734.74
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column05 
            Button          =   -1  'True
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   585.071
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   524.976
         EndProperty
         BeginProperty Column10 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column11 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   510.236
         EndProperty
         BeginProperty Column12 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   615.118
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Legend: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   15
      TabIndex        =   17
      Top             =   2745
      Width           =   705
   End
   Begin VB.Label lblLegend 
      Caption         =   "Format"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   825
      TabIndex        =   16
      Top             =   2730
      Width           =   10455
   End
   Begin VB.Image imgLogo 
      BorderStyle     =   1  'Fixed Single
      Height          =   930
      Left            =   9555
      Stretch         =   -1  'True
      Top             =   30
      Width           =   1125
   End
   Begin VB.Label Label6 
      Caption         =   "Logo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9915
      TabIndex        =   4
      Top             =   390
      Width           =   480
   End
End
Attribute VB_Name = "frmHybridAutoNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'IT Group Inc. 2011.02.10

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
Public Enum eHybridAutoNumber
    Normal
    AddNewEdit
    Find
End Enum
Private Mode As eHybridAutoNumber

'Other declaration
Public sBit As String
Private vBM As Variant 'Recordset bookmark variable

'Security Acess Level variables
Public lACNew As Boolean
Public lACEdit As Boolean
Public lACDelete As Boolean
Public lACPost As Boolean
Public lACCancel As Boolean
Public lACPrint As Boolean

'AutoNumbering Variables
Dim sPrefix As String
Dim sDay As String
Dim sMonth As String
Dim sYear As String
Dim sSeparator As String
Dim sNumeric As String
Dim sSuffix As String

'Filtering View Variables
Dim sCond As String
Public dtgName As String

'Flag Variables
Private lCheckingMandatory As Boolean
Private lApplyAll As Boolean

'Dynamic Checkbox Creation
Private bInSetCheckboxes As Boolean
Private bSizingEnabled As Boolean


'Add new record to the recordset
Public Sub TBNew()
    'module not available
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
        oBar.BitEnabled Me, ITGLedgerMain.tbrMain, , lACEdit, , , , , , , , , True, , True
    Else
        RSZero
        Mode = Find
    End If
    
    sbRS.Panels(2) = ""
    
    SetDataSource
    SetDataField
    
    

ErrorHandler:
    If Err.Number = -2147217885 Then
        Resume Next
    ElseIf Err.Number = -2147217842 Then 'Operation was cancelled. (Error returned by ITGDateBox)
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
    If Err.Number = -2147217885 Then
        Resume Next
    ElseIf Err.Number = -2147217842 Then 'Operation was cancelled. (Error returned by ITGDateBox)
        TBUndoCurrent
    End If

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
    oBar.BitEnabled Me, ITGLedgerMain.tbrMain, , lACEdit, , , , , , , , , True, , True
    Mode = Normal
    
    MsgBox "Record/s successfully saved.", vbInformation, "ComUnion"
    RellocateChk
    sbRS.Panels(2) = ""

    'Audit trail
    UpdateLogFile "Sec - Auto Number", Trim(rsHeader!cModuleID) & "-" & Trim(rsHeader!cType), IIf(lBoolean, "Inserted", "Updated")

ErrHandler:
    If Err.Number = -2147217885 Then
        Resume Next
    ElseIf Err.Number = -2147217864 Then
        OKUpdate = True
        cn.RollbackTrans
        MsgBox "Record cannot be updated. Some values may have been changed by other user/s since last read." & vbCr & _
                "Records will be automatically refreshed. All changes made to the record will be gone upon refresh.", vbInformation, "ComUnion"
        vBookMark = rsHeader.Bookmark
        oRecordset.UnbindControls
        rsHeader.Requery
        oBar.BitEnabled Me, ITGLedgerMain.tbrMain, , lACEdit, , , , , , , , , True, , True
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
        ErrorLog Err.Number, Err.Description, Me.Name 'Error log
    End If
    
End Sub

'Sets the form & recorset to add/edit mode
Public Sub TBEdit()
    On Error Resume Next
    Mode = AddNewEdit
    oBar.BitEnabled Me, ITGLedgerMain.tbrMain, , , , True, True, , , True, True, , , , True
    Set FrmName = Me
    oFormSetup.FormLocking False
    vBM = rsHeader.Bookmark
End Sub

'Delete record
Public Sub TBDelete()
    'module not available
End Sub

'Search using the frmITGSearch
Public Sub TBFind()
'    Mode = Normal
'    txtModule.Locked = True
'    frmITGSearch.Show
End Sub

'Search using the recordset primary key
Public Sub TBFindPrimary()
    If Mode = Find Then
            
            FormWaitShow "Loading data . . ."
            
            Set rsHeader = Nothing
            Set rsHeader = New ADODB.Recordset
            
            oRecordset.OpenRecordset rsHeader, "*", "AutoNum", sCond, True
            
            Set FrmName = Me
            oFormSetup.FormLocking True
            
            cboModule.Enabled = True
            cboOtherFilter.Enabled = True
            
'            If rsHeader.RecordCount = 0 Then
'                Unload frmWait
'                MsgBox "No matching record/s found.", vbInformation, "ComUnion Search"
'                rsHeader.Close
'                oBar.BitEnabled Me, ITGLedgerMain.tbrMain, , , , , , , , , , , , , True
'                Mode = Find
'                Exit Sub
'            End If
            
            SetDataSource
            SetDataField
            
            Mode = Normal
            
            oBar.BitEnabled Me, ITGLedgerMain.tbrMain, , lACEdit, , , , , , , , , True, , True
        
            Unload frmWait
    Else
        RSZero
        RellocateChk
        oFormSetup.FormSearch True
        Mode = Find
    End If
    
    cboModule.Enabled = True
    cboOtherFilter.Enabled = True
    txtOtherFilter.Locked = False

End Sub

'Reload menu buttons (do not delete this sub)
Public Sub TBBitReload()
    oBar.BitVisible ITGLedgerMain.tbrMain, True
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
    RellocateChk
End Sub

'Move previous
Public Sub TBPrevRec()
    If rsHeader.State <> adStateOpen Then Exit Sub
    oNavRec.MovePrevious rsHeader
    RellocateChk
End Sub

'Move next
Public Sub TBNextRec()
    If rsHeader.State <> adStateOpen Then Exit Sub
    oNavRec.MoveNext rsHeader
    RellocateChk
End Sub

'Move last
Public Sub TBLastRec()
    If rsHeader.State <> adStateOpen Then Exit Sub
    oNavRec.MoveLast rsHeader
    RellocateChk
End Sub

'Add new line to the detail recordset
Public Sub TBNewLine()
    With rsHeader
        .AddNew
        !cCompanyID = COID
        !cModuleID = "SALESTICK"
        !cType = "ALL"
        !lInactive = False
        !nValue = 1
        !lLocked = True
        !lBold = False
        !lVisible = True
    End With
    
    RellocateChk
    rsHeader.MoveLast
    
    Set FrmName = Me
    oFormSetup.FormLocking False
End Sub

'Delete line in the detail recordset
Public Sub TBDeleteLine()
    On Error GoTo ErrorHandler
    
    If rsHeader.RecordCount = 0 Then Exit Sub
    
    If MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo, "ComUnion") = vbNo Then Exit Sub
        
    rsHeader.Delete adAffectCurrent
    
    RellocateChk
    TBPrevRec

    Mode = Normal

    If rsHeader.RecordCount = 0 Then
        RSZero
    End If

ErrorHandler:
    If Err.Number = -2147217885 Then
        Resume Next
    End If

End Sub

'Undo all
Public Sub TBUndoLineAll()
    'Not available for this module
End Sub

'Undo current line
Public Sub TBUndoLineCurrent()
    'Not available for this module
End Sub

'Post current record
Public Sub TBPostRecord()
    'Not available for this module
End Sub

'Cancel current record
Public Sub TBCancelRecord()
    'Not available for this module
End Sub

'Print
Public Sub TBPrintRecord()
    'MsgBox "Unavailable on " & Me.Name
End Sub

'Sets the data source of the controls
Sub SetDataSource()
    Set FrmName = Me
    oRecordset.BindControls rsHeader
    Set dtgAutoNum.DataSource = rsHeader
End Sub
    
'Sets the data field for every bounded controls
Sub SetDataField()
    RellocateChk
End Sub

Private Sub cboType_Click()
    If Mode <> AddNewEdit Then Exit Sub
    'dtgAutoNum.Columns(5).Text = cboType.Text
    rsHeader!cType = cboType.Text
    cboType.Visible = False
End Sub

Private Sub chkBold_Click(Index As Integer)
    If lCheckingMandatory = False Then UpdateCHKField Index, 11, rsHeader, chkBold
End Sub

Private Sub chkLocked_Click(Index As Integer)
    If lCheckingMandatory = False Then UpdateCHKField Index, 10, rsHeader, chkLocked
End Sub

Private Sub chkVisible_Click(Index As Integer)
    If lCheckingMandatory = False Then UpdateCHKField Index, 12, rsHeader, chkVisible
End Sub

Private Sub chkInactive_Click(Index As Integer)
    If lCheckingMandatory = False Then UpdateCHKField Index, 0, rsHeader, chkInactive
End Sub

Public Sub UpdateCHKField(Index As Integer, colnum As Integer, tmpRS As Recordset, chk As Object)
    If Mode <> AddNewEdit Then Exit Sub
    If tmpRS.State = adStateOpen Then
        With dtgAutoNum
            On Error Resume Next
            tmpRS.Bookmark = .RowBookmark(Index)
            tmpRS.Fields(colnum).Value = chk(Index).Value * -1
        End With
    End If
End Sub

Private Sub cmdApplyBold_Click()
    Apply_All "lBold", chkApplyBold
End Sub

Private Sub cmdApplyFormat_Click()
    Dim I As Integer
    lApplyAll = True
    If Mode = AddNewEdit Then
        rsHeader.MoveFirst
        For I = 0 To rsHeader.RecordCount - 1
            rsHeader!cFormat = txtUniFormat
            rsHeader.MoveNext
        Next I
        rsHeader.MoveFirst
    End If
    RellocateChk
    lApplyAll = False
End Sub

Private Sub cmdApplyInactive_Click()
    Apply_All "lInactive", chkApplyInactive
End Sub

Public Sub Apply_All(sField As String, chk As Object)
    Dim I As Integer
    lApplyAll = True
    If Mode = AddNewEdit Then
        rsHeader.MoveFirst
        For I = 0 To rsHeader.RecordCount - 1
            rsHeader.Fields(sField) = chk.Value * -1
            rsHeader.MoveNext
        Next I
        rsHeader.MoveFirst
    End If
    RellocateChk
    lApplyAll = False
End Sub

Private Sub cmdApplyLocked_Click()
    Apply_All "lLocked", chkApplyLocked
End Sub

Private Sub cmdApplyVisible_Click()
    Apply_All "lVisible", chkApplyVisible
End Sub

Private Sub cmdView_Click()
    sCond = ""
    If (cboModule.Text <> "") Then
        sCond = " where cModuleID='" & cboModule.Text & "' "
    Else
        sCond = ""
    End If
    
    If txtOtherFilter <> "" Then
        If sCond = "" Then
            sCond = " where " & cboOtherFilter.Text & "='" & txtOtherFilter & "'"
        Else
            sCond = sCond & " AND " & cboOtherFilter.Text & "='" & txtOtherFilter & "'"
        End If
    End If
    Mode = Find
    TBFindPrimary
End Sub

Private Sub dtgAutoNum_AfterColUpdate(ByVal ColIndex As Integer)
    On Error GoTo ErrorHandler
    If Mode <> AddNewEdit Then Exit Sub
    Select Case ColIndex
        Case 4  'Table Name
            rsHeader!cTableName = UCase(rsHeader!cTableName)
        Case 5  'Type
            rsHeader!cType = UCase(rsHeader!cType)
        Case 6  'Format
            rsHeader!cFormat = UCase(rsHeader!cFormat)
        Case 7  'Prefix
            rsHeader!cPrefix = UCase(rsHeader!cPrefix)
        Case 8  'Suffix
            rsHeader!cSuffix = UCase(rsHeader!cSuffix)
    End Select
ErrorHandler:
End Sub

Private Sub dtgAutoNum_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
    If lApplyAll = False Then RellocateChk
End Sub

Private Sub dtgAutoNum_HeadClick(ByVal ColIndex As Integer)
    SortGrid dtgAutoNum, ColIndex, rsHeader
End Sub

Private Sub dtgAutoNum_ButtonClick(ByVal ColIndex As Integer)
    If Mode = AddNewEdit Then 'Exit Sub
        Select Case ColIndex
            Case 5
                MoveCombo cboType, dtgAutoNum, dtgAutoNum.Columns(ColIndex)
        End Select
    End If
End Sub

Private Sub dtgAutoNum_KeyPress(KeyAscii As Integer)
    If Mode = AddNewEdit Then 'Exit Sub
        Select Case dtgAutoNum.Col
            Case 6
                If UCase(Chr(KeyAscii)) = UCase(sPrefix) Then
                ElseIf UCase(Chr(KeyAscii)) = UCase(sDay) Then
                ElseIf UCase(Chr(KeyAscii)) = UCase(sMonth) Then
                ElseIf UCase(Chr(KeyAscii)) = UCase(sYear) Then
                ElseIf UCase(Chr(KeyAscii)) = UCase(sSeparator) Then
                ElseIf UCase(Chr(KeyAscii)) = UCase(sNumeric) Then
                ElseIf UCase(Chr(KeyAscii)) = UCase(sSuffix) Then
                ElseIf KeyAscii = 8 Then
                Else
                    KeyAscii = 0
                End If
        End Select
    End If
End Sub

Private Sub dtgAutoNum_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    If Mode <> AddNewEdit Then Exit Sub
'    If Button = 2 Then
'        dtgName = dtgAutoNum.Name
'        PopupMenu ITGLedgerMain.mnuDetail
'    End If
End Sub


Private Sub dtgAutoNum_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If lApplyAll = False Then RellocateChk
End Sub

Private Sub dtgAutoNum_RowResize(Cancel As Integer)
    If lApplyAll = False Then RellocateChk
End Sub

Private Sub dtgAutoNum_Scroll(Cancel As Integer)
    If lApplyAll = False Then RellocateChk
End Sub

'Private Sub Form_Deactivate()
'    WheelUnHook
'End Sub

'Set Your Object
Private Sub Form_Load()
    
    oBar.AcessBit Me, GetValueFrTable("AccessLevel", "SEC_ACCESSLEVEL", "RoleID = '" & SecUserRole & "' AND [Module] = 'AUTONUM'")
    
    Set rsHeader = New ADODB.Recordset
    Set oNavRec = New clsNavRec
    
    Set FrmName = Me
    oFormSetup.FormLocking True
    oBar.BitEnabled Me, ITGLedgerMain.tbrMain, , , , , , , , , , , , , True
    oBar.BitVisible ITGLedgerMain.tbrMain, True
    ITGLedgerMain.tbrMain.Buttons("btnFind").ButtonMenus("btnFindP").Enabled = True

    Mode = Find
    
    Me.Icon = ITGLedgerMain.Icon
    imgLogo.Picture = ITGLedgerMain.Picture1.Picture
    
    TBFindPrimary
    DoEvents
    RellocateChk
    
    GetConst_AutoNum_Param sPrefix, sDay, sMonth, sYear, sSeparator, sNumeric, sSuffix
    
    lblLegend.Caption = "Prefix: " & sPrefix & " ;  Day: " & sDay & " ;  Month: " & sMonth & _
            " ;  Year: " & sYear & " ;  Separator: " & sSeparator & " ;  Numeric: " & sNumeric & " ;  Suffix: " & sSuffix
    
    LoadComboValues cboModule, "cModuleID", "AutoNum", "Where cCompanyID='" & COID & "'"
    LoadComboValues cboType, "cTypeID", "Ticket_Type", "Where cCompanyID='" & COID & "'"
        
End Sub

Public Sub RellocateChk()
'    MoveCHK rsHeader, dtgAutoNum, 0, chkInactive
'    MoveCHK rsHeader, dtgAutoNum, 10, chkLocked
'    MoveCHK rsHeader, dtgAutoNum, 11, chkBold
'    MoveCHK rsHeader, dtgAutoNum, 12, chkVisible
    SetCheckboxes 0, chkInactive
    SetCheckboxes 10, chkLocked
    SetCheckboxes 11, chkBold
    SetCheckboxes 12, chkVisible
End Sub


' Must be called from:
'   ColResize
'   RowResize
'   RowColChange
'   Scroll
'   [parent form].Load (after Rebinding)
'   [parent form].Resize

Private Sub SetCheckboxes(ColNdx As Long, ByRef ChkboxArray As Object)
    bInSetCheckboxes = True
On Error GoTo ErrorExit
    Dim I
    Dim obj As Object
    Set obj = dtgAutoNum
        
    Dim OffsetX As Long, OffsetY As Long
    If Not ChkboxArray(0).Container Is dtgAutoNum.Container Then
        CalcContainerOffset obj, OffsetX, OffsetY
    End If
    
    On Error Resume Next
    
    With dtgAutoNum
        If (ChkboxArray.UBound <> .VisibleRows) Then
            For I = ChkboxArray.UBound + 1 To .VisibleRows - 1
                Load ChkboxArray(I)
                ChkboxArray(I).Width = 190
                ChkboxArray(I).Height = 190
            Next
            For I = .VisibleRows To ChkboxArray.UBound
                Unload ChkboxArray(I)
            Next
        End If
    
        OffsetX = OffsetX + (.Columns(ColNdx).Width - ChkboxArray(0).Width) / 2
        OffsetY = OffsetY + 10 ''(.RowHeight - ChkboxArray(0).Height) / 2

        .Columns(ColNdx).Alignment = dbgCenter
        .Columns(ColNdx).Locked = True
        
        '
        If .LeftCol <= ColNdx Then
            For I = 0 To .VisibleRows - 1
                ChkboxArray(I).Value = Abs(.Columns(ColNdx).CellValue(.RowBookmark(I)))
                ChkboxArray(I).Top = .Top + .RowTop(I) + OffsetY + 20
                ChkboxArray(I).Left = .Left + .Columns(ColNdx).Left + OffsetX
                ChkboxArray(I).Visible = True
                ChkboxArray(I).ZOrder
            Next
        Else
            I = 0
        End If
        
        For I = I To ChkboxArray.UBound
            ChkboxArray(I).Visible = False
        Next
        
    End With
   
ExitPoint:
    bInSetCheckboxes = False
    Exit Sub

ErrorExit:
    Resume ExitPoint
End Sub

Public Function CalcContainerOffset( _
        obj As Object, _
        ByRef OffsetX As Long, _
        ByRef OffsetY As Long _
    )
    
    
    Do While Not (obj.Container Is obj.Parent)
        Set obj = obj.Container
        If Not (obj Is Nothing) Then
            OffsetX = OffsetX + obj.Left
            OffsetY = OffsetY + obj.Top
            
            '' The offsets for borders below are not exact for frames,
            '' this positioning algorithm works perfectly at any depth
            '' of nested pictureboxes, with any combination of borders
            '' and/or 3D at any levels.
            ''
            '' Using a frame with borders and/or 3d throws a visible skew
            '' on the positions but this should be fixable with some trial
            '' and error -- the skew is no more than 2 px. per frame.
            ''
            '' Other containers may be subject to different metrics.
            ''
            If obj.BorderStyle = 1 Then '' fixed single
                If obj.Appearance = 1 Then '' 3d
                    OffsetX = OffsetX + 30
                    OffsetY = OffsetY + 30
                Else
                    OffsetX = OffsetX + 15
                    OffsetY = OffsetY + 15
                End If
            End If
        End If
        If (TypeOf obj Is Form) Or (TypeOf obj Is MDIForm) Then Exit Do
    Loop


End Function

'Public Sub MoveCHK(tmpRS As Recordset, dtg As DataGrid, colnum As Integer, chk As Object)
'    Dim i As Integer
'
'    'Initialized the objects (checkboxes)
'    For i = 1 To chk.ubound
'        Unload chk(i)
'    Next i
'
'    On Error Resume Next
'    For i = 0 To dtg.VisibleRows - 1 'tmpRS.RecordCount - 1
'        Load chk(i)
'        chk(i).Width = 190
'        chk(i).Height = 190
'        chk(i).Visible = True
'        chk(i).Move dtg.Left + dtg.Columns(colnum).Left + (dtg.Columns(colnum).Width / 2.5), _
'            dtg.Top + dtg.RowTop(i) + 50
'        chk(i).Value = Abs(dtg.Columns(i).Text) 'Abs(tmpRS.Fields(dtg.Columns(colnum).DataField))
'        chk(i).ZOrder
'        'tmpRS.MoveNext
'    Next i
'    'tmpRS.MoveFirst
'End Sub

'Activate your Toolbar Mode
Private Sub Form_Activate()
    TBBitReload
'    WheelHook dtgAutoNum
End Sub

'Release your Object
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
    
    Set frmHybridAutoNum = Nothing
    
    lCloseWindow = True
End Sub

Private Sub rsHeader_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo ErrorHandler
    
    If Not (rsHeader.EOF) Or Not (rsHeader.BOF) Then
        sbRS.Panels(1) = "Record: " & IIf((rsHeader.AbsolutePosition = -2), "0", rsHeader.AbsolutePosition) & "/" & rsHeader.RecordCount
        lblHeader.Caption = rsHeader!cModuleID & " - " & rsHeader!cDesc & " (" & rsHeader!cType & ")"
        lblDetails.Caption = rsHeader!cFormat
        If Mode = AddNewEdit Then
            Select Case rsHeader.Status
                Case adRecNew
                    sbRS.Panels(2) = "New"
                Case adRecModified
                    sbRS.Panels(2) = "Modified"
                Case Else
                    sbRS.Panels(2) = ""
                    'RellocateChk
            End Select
        Else
            sbRS.Panels(2) = ""
        End If
    Else
        sbRS.Panels(1) = "Record: 0/0"
        sbRS.Panels(2) = ""
        txtModule.Locked = False
    End If
    
ErrorHandler:
    'Err.Number -2147217885
    'Description - Row handle referred to a deleted row or a row marked for deletion.
    If Err.Number = -2147217885 Then
        Resume Next
    End If
    
End Sub

'Check if all mandatory fields are complete
Function MandatoryOK() As Boolean
    lCheckingMandatory = True
    MandatoryOK = True
    
    If rsHeader.RecordCount <> 0 Then rsHeader.MoveFirst
    Do Until rsHeader.EOF
        If rsHeader!cTableName = "" Then
            MandatoryOK = False
            MsgBox "Table Name is mandatory. Null value is not allowed.", vbInformation, "ComUnion"
            Exit Function
        ElseIf rsHeader!cType = "" Then
            MandatoryOK = False
            MsgBox "Module Type is mandatory. Null value is not allowed.", vbInformation, "ComUnion"
            Exit Function
        ElseIf rsHeader!cFormat = "" Then
            MandatoryOK = False
            MsgBox "Format  is mandatory. Null value is not allowed.", vbInformation, "ComUnion"
            Exit Function
        End If
        
        rsHeader.MoveNext
    Loop
    rsHeader.MoveFirst
    lCheckingMandatory = False
End Function

'Sets the form if record number is zero
Private Sub RSZero()
    sbRS.Panels(1) = "Record: 0/0"
    sbRS.Panels(2) = ""
    
    Set FrmName = Me
    oRecordset.UnbindControls
    oFormSetup.TextClearing
    oFormSetup.FormLocking True
    
    Set dtgAutoNum.DataSource = Nothing
    dtgAutoNum.Refresh
    
    If rsHeader.State = adStateOpen Then rsHeader.Close
    
    oBar.BitEnabled Me, ITGLedgerMain.tbrMain, lACNew, , , , , , , , , True, , , True
    
    cboModule.Enabled = True
    cboOtherFilter.Enabled = True
    
    Mode = Find
End Sub

Private Sub txtUniFormat_KeyPress(KeyAscii As Integer)
    If UCase(Chr(KeyAscii)) = UCase(sPrefix) Then
    ElseIf UCase(Chr(KeyAscii)) = UCase(sDay) Then
    ElseIf UCase(Chr(KeyAscii)) = UCase(sMonth) Then
    ElseIf UCase(Chr(KeyAscii)) = UCase(sYear) Then
    ElseIf UCase(Chr(KeyAscii)) = UCase(sSeparator) Then
    ElseIf UCase(Chr(KeyAscii)) = UCase(sNumeric) Then
    ElseIf UCase(Chr(KeyAscii)) = UCase(sSuffix) Then
    ElseIf KeyAscii = 8 Then
    Else
        KeyAscii = 0
    End If
End Sub

