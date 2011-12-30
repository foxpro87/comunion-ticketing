VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B3FB64BF-91F9-11D7-A482-0008A14158BC}#2.22#0"; "ITGControls.ocx"
Begin VB.Form frmToolAuditTrail 
   Caption         =   "Audit Trail"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmToolAuditTrail.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8040
   ScaleWidth      =   11775
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   315
      Left            =   10140
      TabIndex        =   10
      Top             =   615
      Width           =   1380
   End
   Begin VB.CommandButton cmdShowList 
      Caption         =   "&Show List"
      Height          =   315
      Left            =   10140
      TabIndex        =   9
      Top             =   270
      Width           =   1380
   End
   Begin VB.Frame fraFiltering 
      Height          =   1200
      Left            =   1500
      TabIndex        =   4
      Top             =   90
      Width           =   10170
      Begin VB.ComboBox cboMachine 
         Height          =   315
         Left            =   4740
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   525
         Width           =   1965
      End
      Begin ITGControls.ITGLabel lblAudit 
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   14
         Top             =   195
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
         Caption         =   "Module"
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
      Begin VB.ComboBox cboEvent 
         Height          =   315
         ItemData        =   "frmToolAuditTrail.frx":0CCA
         Left            =   1320
         List            =   "frmToolAuditTrail.frx":0CEC
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   495
         Width           =   1950
      End
      Begin VB.ComboBox cboUser 
         Height          =   315
         Left            =   4740
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   180
         Width           =   1965
      End
      Begin VB.ComboBox cboModule 
         Height          =   315
         ItemData        =   "frmToolAuditTrail.frx":0D62
         Left            =   1320
         List            =   "frmToolAuditTrail.frx":0E3E
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   165
         Width           =   1950
      End
      Begin ITGControls.ITGDateBox dtbFrom 
         Height          =   285
         Left            =   7350
         TabIndex        =   5
         Tag             =   "Order Date"
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
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
         Text            =   ""
         SendKeysTab     =   -1  'True
         Mandatory       =   -1  'True
      End
      Begin ITGControls.ITGLabel ITGLabel4 
         Height          =   285
         Left            =   6810
         TabIndex        =   6
         Top             =   180
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   503
         Caption         =   "From :"
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
      Begin ITGControls.ITGDateBox dtbTo 
         Height          =   285
         Left            =   7350
         TabIndex        =   7
         Tag             =   "Order Date"
         Top             =   495
         Width           =   1275
         _ExtentX        =   2249
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
         Text            =   ""
         SendKeysTab     =   -1  'True
         Mandatory       =   -1  'True
      End
      Begin ITGControls.ITGLabel ITGLabel1 
         Height          =   285
         Left            =   6825
         TabIndex        =   8
         Top             =   495
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         Caption         =   "To :"
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
      Begin ITGControls.ITGLabel lblAudit 
         Height          =   285
         Index           =   1
         Left            =   3450
         TabIndex        =   15
         Top             =   210
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   503
         Caption         =   "User / Logon ID"
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
      Begin ITGControls.ITGLabel lblAudit 
         Height          =   285
         Index           =   2
         Left            =   105
         TabIndex        =   16
         Top             =   510
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   503
         Caption         =   "Event"
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
      Begin ITGControls.ITGLabel lblAudit 
         Height          =   285
         Index           =   3
         Left            =   3435
         TabIndex        =   18
         Top             =   525
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         Caption         =   "Computer Name"
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
      Begin ITGControls.ITGTextBox txtTranNo 
         Height          =   285
         Left            =   90
         TabIndex        =   19
         Top             =   840
         Width           =   3150
         _ExtentX        =   5345
         _ExtentY        =   503
         SendKeysTab     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Label           =   "Transaction No"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelWidth      =   1230
         TextBoxWidth    =   1860
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "List Filter"
      Height          =   1200
      Left            =   30
      TabIndex        =   1
      Top             =   90
      Width           =   1440
      Begin VB.OptionButton optAll 
         Caption         =   "All"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   330
         Width           =   735
      End
      Begin VB.OptionButton optFilter 
         Caption         =   "Filter by "
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   630
         Width           =   1035
      End
   End
   Begin MSComctlLib.ListView lvwAudit 
      Height          =   6360
      Left            =   15
      TabIndex        =   0
      Top             =   1350
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   11218
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmToolAuditTrail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'IT Group Inc. 2005.09.23
Option Explicit

'Object variables
Private oNavRec As clsNavRec
Private oBar As New clsToolBarMenuBit

Public sBit As String

'Security Acess Level variables
Public lACNew As Boolean
Public lACEdit As Boolean
Public lACDelete As Boolean
Public lACPost As Boolean
Public lACCancel As Boolean
Public lACPrint As Boolean

Enum eAuditList
    AllName
    username
    ModuleName
    DateRange
End Enum

Public mAuditList As eAuditList

Private lLoaded As Boolean

Private oConnection As New clsConnection
Private connList As ADODB.Connection



Private Sub cmdRefresh_Click()
    cmdShowList_Click
End Sub

Private Sub cmdShowList_Click()
Dim sWhere As String
    If optAll.Value = False Then
        If Trim(cboModule) <> "" Then sWhere = sWhere & IIf(Trim$(sWhere) = "", "", " AND ") & " cModule = '" & Trim$(cboModule) & "'"
        If Trim(cboUser) <> "" Then sWhere = sWhere & IIf(Trim$(sWhere) = "", "", " AND ") & " cUser ='" & Trim$(cboUser) & "'"
        If Trim(txtTranNo) <> "" Then sWhere = sWhere & IIf(Trim$(sWhere) = "", "", " AND ") & " cTranNo ='" & Trim$(txtTranNo) & "'"
        If Trim(cboEvent) <> "" Then sWhere = sWhere & IIf(Trim$(sWhere) = "", "", " AND ") & " cEvent = '" & Trim$(cboEvent) & "'"
        If Trim(cboMachine) <> "" Then sWhere = sWhere & IIf(Trim$(sWhere) = "", "", " AND ") & " cMachine = '" & Trim$(cboMachine) & "'"
        If IsDate(dtbFrom.Text) And (IsNull(dtbTo.Text) Or dtbTo.Text = "__/__/____") Then
            sWhere = sWhere & IIf(Trim$(sWhere) = "", "", " AND ") & " dDate >= '" & Trim$(dtbFrom.Text) & "'"
        ElseIf IsDate(dtbTo.Text) And (IsNull(dtbFrom.Text) Or dtbFrom.Text = "__/__/____") Then
            sWhere = sWhere & IIf(Trim$(sWhere) = "", "", " AND ") & " dDate <= '" & Trim$(dtbTo.Text) & "'"
        ElseIf IsDate(dtbFrom.Text) And IsDate(dtbTo.Text) Then
            sWhere = sWhere & IIf(Trim$(sWhere) = "", "", " AND ") & " dDate BETWEEN '" & Trim(dtbFrom.Text) & "' AND '" & Trim(dtbTo.Text) & "'"
        End If
        LoadList sWhere
    Else: LoadList
    End If
End Sub

'Activate your Toolbar Mode
Private Sub Form_Activate()
    TBBitReload
End Sub

'Release your Object
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    oBar.BitEnabled Me, ITGLedgerMain.tbrMain, , , , , , , , , , True
    oBar.BitVisible ITGLedgerMain.tbrMain

    Set oNavRec = Nothing
    Set oBar = Nothing
    
    Set frmToolAuditTrail = Nothing
    
    lCloseWindow = True
End Sub

'Reload menu buttons (do not delete this sub)
Public Sub TBBitReload()
    oBar.BitVisible ITGLedgerMain.tbrMain
    oBar.BitReload Me, ITGLedgerMain.tbrMain, sBit
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
    Set oNavRec = New clsNavRec
    Top = 0
    Left = 2675
    Height = 6600
    Width = 11895
    
    optAll.Value = True
    lLoaded = True
    
    Me.dtbTo.Text = Format(Date, "MM/dd/yyyy")
    Me.dtbFrom.Text = Format((DateAdd("m", -1, Date) + 1), "MM/dd/yyyy")

    Set FrmName = Me
    oBar.BitEnabled Me, ITGLedgerMain.tbrMain, , , , , , , , , , True, , , True
    oBar.BitVisible ITGLedgerMain.tbrMain, True
    ITGLedgerMain.tbrMain.Buttons("btnFind").ButtonMenus("btnFindP").Enabled = True
    
    LoadComboDetails
    LoadList
End Sub

Private Sub Form_Resize()
    If Not lLoaded Then Exit Sub
    If Me.WindowState <> vbMinimized Then
        If Me.Height < 1600 Then Me.Height = 1600
        lvwAudit.Height = frmToolAuditTrail.Height - 1900
        lvwAudit.Width = frmToolAuditTrail.Width - 220
    End If
End Sub

'Load transaction list for approval
Sub LoadList(Optional Condition As String)
    lvwAudit.ColumnHeaders.Clear
    lvwAudit.ListItems.Clear
    If rs.State = adStateOpen Then rs.Close
    oConnection.OpenNewConnection connList

    Select Case mAuditList
        Case AllName

            Caption = "All Transactions"

            Set itmX = lvwAudit.ColumnHeaders.Add(, , "Module Name")
            Set itmX = lvwAudit.ColumnHeaders.Add(, , "Date")
            Set itmX = lvwAudit.ColumnHeaders.Add(, , "Transaction No")
            Set itmX = lvwAudit.ColumnHeaders.Add(, , "Event")
            Set itmX = lvwAudit.ColumnHeaders.Add(, , "User Name")
            Set itmX = lvwAudit.ColumnHeaders.Add(, , "Machine Name")
            lvwAudit.ColumnHeaders(1).Width = "1750"
            lvwAudit.ColumnHeaders(2).Width = "2000"
            lvwAudit.ColumnHeaders(3).Width = "1750"
            lvwAudit.ColumnHeaders(4).Width = "2400"
            lvwAudit.ColumnHeaders(5).Width = "1750"
            lvwAudit.ColumnHeaders(6).Width = "1750"
                
             
            sSQL = "SELECT cModule, dDate, cTranNo, cEvent, cUser, cMachine FROM LogFile " & _
                    "WHERE cCompanyID = '" & COID & "' " & IIf(Condition = "", "", " And " & Condition) & "AND convert(datetime,convert(varchar(20),dDate,101)) BETWEEN '" & Trim(dtbFrom.Text) & "' AND '" & Trim(dtbTo.Text) & "' ORDER BY dDate DESC"
            oConnection.OpenNewConnection connList
            rs.Open sSQL, connList, adOpenKeyset

            Do Until rs.EOF
                Set itmX = lvwAudit.ListItems.Add(, , Trim(rs!cModule))
                    itmX.SubItems(1) = Format(rs!dDate, "MM/dd/yyyy   H:MM:SS")
                    itmX.SubItems(2) = Trim(rs!cTranNo)
                    itmX.SubItems(3) = Trim(rs!cEvent)
                    itmX.SubItems(4) = Trim(rs!cUser)
                    itmX.SubItems(5) = Trim(rs!cMachine)
                rs.MoveNext
            Loop

            Set rs = Nothing
            Set connList = Nothing

        Case ModuleName
        Case username
    End Select
End Sub

Private Sub lvwAudit_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwAudit.Sorted = True
    If lvwAudit.SortOrder = lvwAscending Then
        lvwAudit.SortOrder = lvwDescending
    Else: lvwAudit.SortOrder = lvwAscending
    End If
    lvwAudit.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub optAll_Click()
    cboModule.Enabled = False
    cboUser.Enabled = False
    cboEvent.Enabled = False
    cboMachine.Enabled = False
    dtbFrom.Enabled = False
    dtbTo.Enabled = False
    cmdShowList.Enabled = True
    fraFiltering.Enabled = False
End Sub

Private Sub optFilter_Click()
    lvwAudit.ListItems.Clear
    cboModule.Enabled = True
    cboUser.Enabled = True 'False
    cboEvent.Enabled = True
    cboMachine.Enabled = True 'False
    dtbFrom.Enabled = True
    dtbTo.Enabled = True
    cmdShowList.Enabled = True
    fraFiltering.Enabled = True
End Sub

Sub LoadComboDetails()
    Call LoadComboValues(cboModule, "cModule", "LOGFILE", "where cCompanyID = '" & Trim(COID) & "'", "cModule")
    Call LoadComboValues(cboUser, "cUser", "LOGFILE", "where cCompanyID = '" & Trim(COID) & "'", "cUser")
    Call LoadComboValues(cboMachine, "cMachine", "LOGFILE", "where cCompanyID = '" & Trim(COID) & "'", "cMachine")
    Call LoadComboValues(cboEvent, "cEvent", "LOGFILE", "where cCompanyID = '" & Trim(COID) & "'", "cEvent")
    
    cboModule.AddItem "", 0
    cboUser.AddItem "", 0
    cboEvent.AddItem "", 0
    cboMachine.AddItem "", 0
End Sub
