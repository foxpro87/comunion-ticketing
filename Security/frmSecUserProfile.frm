VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{B3FB64BF-91F9-11D7-A482-0008A14158BC}#2.22#0"; "ITGControls.ocx"
Begin VB.Form frmSecUserProfile 
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Security - User Profile"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSecUserProfile.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   7440
   Begin MSComctlLib.StatusBar sbRS 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2940
      Width           =   7440
      _ExtentX        =   13123
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
      Height          =   2850
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   5027
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
         Height          =   2535
         Left            =   -74940
         TabIndex        =   8
         Top             =   360
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   4471
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
            DataField       =   "UserID"
            Caption         =   "User ID"
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
            DataField       =   "LastName"
            Caption         =   "Last Name"
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
            DataField       =   "FirstName"
            Caption         =   "First Name"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "MI"
            Caption         =   "MI"
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
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   2399.811
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   2429.858
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   615.118
            EndProperty
         EndProperty
      End
      Begin ITGControls.ComunionFrames ComunionFrames2 
         Height          =   2430
         Left            =   75
         Top             =   345
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   4286
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
         Begin VB.CheckBox chkAdmin 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Administrator"
            Height          =   195
            Left            =   5340
            TabIndex        =   2
            Top             =   585
            Width           =   1725
         End
         Begin VB.Timer Timer1 
            Interval        =   300
            Left            =   3045
            Top             =   465
         End
         Begin ITGControls.ITGTextBox txtUserID 
            Height          =   285
            Left            =   210
            TabIndex        =   3
            Top             =   555
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
            AllCaps         =   -1  'True
            Mandatory       =   -1  'True
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
            LabelWidth      =   1280
            TextBoxWidth    =   1360
         End
         Begin ITGControls.ITGTextBox txtFName 
            Height          =   285
            Left            =   225
            TabIndex        =   4
            Top             =   1665
            Width           =   4080
            _ExtentX        =   6985
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
            AllCaps         =   -1  'True
            Label           =   "First Name"
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
            TextBoxWidth    =   2740
         End
         Begin ITGControls.ITGTextBox txtMI 
            Height          =   285
            Left            =   225
            TabIndex        =   5
            Top             =   2025
            Width           =   1800
            _ExtentX        =   2963
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
            AllCaps         =   -1  'True
            Label           =   "Middle Initial"
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
            TextBoxWidth    =   460
         End
         Begin ITGControls.ITGTextBox txtLName 
            Height          =   285
            Left            =   225
            TabIndex        =   6
            Top             =   1305
            Width           =   4080
            _ExtentX        =   6985
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
            AllCaps         =   -1  'True
            Label           =   "Last Name"
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
            TextBoxWidth    =   2740
         End
         Begin VB.Line Line1 
            X1              =   1185
            X2              =   7125
            Y1              =   1065
            Y2              =   1065
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Information"
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
            Height          =   195
            Left            =   105
            TabIndex        =   7
            Top             =   945
            Width           =   1020
         End
      End
   End
End
Attribute VB_Name = "frmSecUserProfile"
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
Public Enum eSecUPMode
    Normal
    AddNewEdit
    Find
End Enum
Private Mode As eSecUPMode

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
 
 
Private Sub dtgList_HeadClick(ByVal ColIndex As Integer)
    SortGrid dtgList, ColIndex, rsHeader
End Sub

'Set Your Object
Private Sub Form_Load()
    
    oBar.AcessBit Me, GetValueFrTable("AccessLevel", "SEC_ACCESSLEVEL", "RoleID = '" & SecUserRole & "' AND [Module] = 'SUSER'")
    
    Set rsHeader = New ADODB.Recordset
    Set oNavRec = New clsNavRec
    
    Set FrmName = Me
    oFormSetup.FormLocking True
    oBar.BitEnabled Me, ITGLedgerMain.tbrMain, lACNew, , , , , , , , , True, , , True
    oBar.BitVisible ITGLedgerMain.tbrMain
    ITGLedgerMain.tbrMain.Buttons("btnFind").ButtonMenus("btnFindP").Enabled = True

    Mode = Find
    txtUserID.Locked = False
    
    
    If LoadOption("COM_THEME", 4) = "1 - Blue" Then
        oFormSetup.FormTheme (1)
    Else
        oFormSetup.FormTheme (2)
    End If
    
End Sub

'Activate your Toolbar Mode
Private Sub Form_Activate()
    TBBitReload
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
    
    Set frmSecUserProfile = Nothing

    lCloseWindow = True
End Sub

'Add new record to the recordset
Public Sub TBNew()
    
    oBar.BitEnabled Me, ITGLedgerMain.tbrMain, , , , True, True, , , True, True, , , , True
    txtUserID.BackColor = &HE0FFFF
    Mode = AddNewEdit
    
    If rsHeader.State <> adStateOpen Then
        oRecordset.OpenRecordset rsHeader, "*", "SEC_USER", "WHERE 1 = 0 ", True
        SetDataSource
        SetDataField
    Else
        vBM = rsHeader.Bookmark
    End If
    
    rsHeader.AddNew
    'rsHeader!RoleID = "SUPERUSER"
    
    
    Set FrmName = Me
    oFormSetup.FormLocking False
    oFormSetup.ClrRequired &HC0&
    SSTab1.ActiveTab = 0
    SSTab1.TabEnabled(1) = False
    txtUserID.SetFocus

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
        oBar.BitEnabled Me, ITGLedgerMain.tbrMain, lACNew, lACEdit, lACDelete, , , , , , , True, True, , True
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

'Save all changes
Public Sub TBSave()
Dim OKUpdate As Boolean
On Error GoTo ErrHandler
    
    'Audit Trail
    lBoolean = False
    If rsHeader.Status = adRecNew Then lBoolean = True
    
    If Not MandatoryOK Then Exit Sub
    
    If chkAdmin.Value = 1 Then
        rsHeader!RoleID = "SUPERUSER"
    Else
        rsHeader!RoleID = "EMPLOYEE"
    End If
        
    
    OKUpdate = False
    cn.BeginTrans
    
    rsHeader!Password = Encrypt("PASSWORD")
    
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
    
    
    'Company Access
    Dim sCmpNme As String
    sCmpNme = GetValueFrTable("cCompanyName", "COMPANY", "cCompanyID = '" & COID & "'", True)
    cn.Execute ("insert into SEC_COMPANYACCESS values ('" & txtUserID & "','" & COID & "',NULL,'" & sCmpNme & "')")

    'Audit trail
    UpdateLogFile "Sec - User Profile", Trim(txtUserID), IIf(lBoolean, "Inserted", "Updated")

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
    txtUserID.Locked = True
    SSTab1.ActiveTab = 0
    SSTab1.TabEnabled(1) = False
    txtLName.SetFocus
    vBM = rsHeader.Bookmark
End Sub

'Delete record
Public Sub TBDelete()
On Error GoTo ErrorHandler
    
    If rsHeader.RecordCount = 0 Then Exit Sub
    
    If UCase(Trim(txtUserID)) = "SA" Then
        MsgBox "User cannot be deleted. [Administrator ID]", vbCritical, "ComUnion"
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo, "ComUnion") = vbNo Then Exit Sub
    
    'Audit trail
    UpdateLogFile "Sec - User Profile", Trim(txtUserID), "Deleted"
    
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
    txtUserID.Locked = True
'    frmITGSearch.Show
End Sub

'Search using the recordset primary key
Public Sub TBFindPrimary()
    RSZero
    Mode = Find
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

'Add new line to the detail recordset
Public Sub TBNewLine()
    'Not available for this module
End Sub

'Delete line in the detail recordset
Public Sub TBDeleteLine()
    'Not avilable for this module
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
    MsgBox "Unavailable on " & Me.Name
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
        txtUserID.DataField = !UserID
        txtLName.DataField = !LastName
        txtFName.DataField = !FirstName
        txtMI.DataField = !MI
'        txtRoleID.DataField = !RoleID
'        txtDeptID.DataField = !DeptID
'        txtUnitDiv.DataField = !cDivision
    End With
End Sub

Private Sub ITGTab1_AfterCompleteInit()

End Sub

Private Sub rsHeader_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo ErrorHandler
    
    If Not (rsHeader.EOF) Or Not (rsHeader.BOF) Then
        sbRS.Panels(1) = "Record: " & IIf((rsHeader.AbsolutePosition = -2), "0", rsHeader.AbsolutePosition) & "/" & rsHeader.RecordCount
        
        If rsHeader.Status <> adRecNew Then
            txtUserID.Locked = True
        Else
            txtUserID.Locked = False
        End If
        
        If rsHeader!RoleID = "SUPERUSER" Then
            chkAdmin.Value = 1
        Else
            chkAdmin.Value = 0
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
        txtUserID.Locked = False
    End If
    
ErrorHandler:
    'Err.Number -2147217885
    'Description - Row handle referred to a deleted row or a row marked for deletion.
    If err.Number = -2147217885 Then
        Resume Next
    End If
    
End Sub

Private Sub Timer1_Timer()
    If Mode = AddNewEdit Then
        SSTab1.TabEnabled(1) = False
    Else
        SSTab1.TabEnabled(1) = True
    End If
    
    If Mode <> Find Then Exit Sub
    If txtUserID.BackColor = &HE0FFFF Then
        txtUserID.BackColor = &HE0E0E0
        Exit Sub
    End If
    If txtUserID.BackColor <> &HE0FFFF Then
        txtUserID.BackColor = &HE0FFFF
        Exit Sub
    End If

End Sub
 
Private Sub txtUserID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        'Find
        If Mode = Find Then
            
            FormWaitShow "Loading data . . ."
            
            Set rsHeader = Nothing
            Set rsHeader = New ADODB.Recordset
            
            If Trim(txtUserID) = "" Then
                oRecordset.OpenRecordset rsHeader, "*", "SEC_USER", , True
            Else
                oRecordset.OpenRecordset rsHeader, "*", "SEC_USER", "WHERE UserID LIKE '" & Trim(txtUserID) & "%'", True
            End If
            
            Set FrmName = Me
            oFormSetup.FormLocking True
            
            If rsHeader.RecordCount = 0 Then
                Unload frmWait
                MsgBox "No matching record/s found.", vbInformation, "ComUnion Search"
                rsHeader.Close
                oBar.BitEnabled Me, ITGLedgerMain.tbrMain, lACNew, , , , , , , , , True, , , True
                txtUserID.Locked = False
                txtUserID.SetFocus
                Exit Sub
            End If
            
            SetDataSource
            SetDataField
            
            If rsHeader!RoleID = "SUPERUSER" Then
                chkAdmin.Value = 1
            Else
                chkAdmin.Value = 0
            End If
            
            txtUserID.BackColor = &HE0FFFF
            Mode = Normal
            
            oBar.BitEnabled Me, ITGLedgerMain.tbrMain, lACNew, lACEdit, lACDelete, , , , , , , True, True, , True
        
            Unload frmWait
        End If
    End If
End Sub

Private Sub txtUserID_LostFocus()
    'Does Code Already exists
    If Mode = AddNewEdit Then
        If Trim(txtUserID) = "" Then
            MsgBox "Empty primary input.", vbInformation, "ComUnion"
            txtUserID.SetFocus
        Else
            If rsHeader.Status <> adRecNew Then Exit Sub
            txtUserID = Trim(txtUserID)
            If IDExisting(rsHeader, "UserID", "SEC_USER", Trim(rsHeader!UserID), , True) Then
                MsgBox "User ID already exists.", vbInformation, "ComUnion"
                txtUserID.SetFocus
            End If
        End If
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
    
    If rsHeader.Status = adRecNew Then
        If IDExisting(rsHeader, "UserID", "SEC_USER", Trim(rsHeader!UserID), , True) Then
            MandatoryOK = False
            MsgBox "User ID already exists.", vbInformation, "ComUnion"
            txtUserID.SetFocus
            Exit Function
        End If
    End If
    
    If Trim(txtLName) = "" Then
        MandatoryOK = False
        MsgBox "Field 'Last Name' is mandatory. Null value is not allowed", vbInformation, "ComUnion"
        txtLName.SetFocus
        Exit Function
    ElseIf Trim(txtFName) = "" Then
        MandatoryOK = False
        MsgBox "Field 'First Name' is mandatory. Null value is not allowed", vbInformation, "ComUnion"
        txtFName.SetFocus
        Exit Function
    ElseIf Trim(txtMI) = "" Then
        MandatoryOK = False
        MsgBox "Field 'Middle Initial' is mandatory. Null value is not allowed", vbInformation, "ComUnion"
        txtMI.SetFocus
        Exit Function
'    ElseIf Trim(txtRoleID) = "" Then
'        MandatoryOK = False
'        MsgBox "Field 'User Role' is mandatory. Null value is not allowed", vbInformation, "ComUnion"
'        txtRoleID.SetFocus
'        Exit Function
'    ElseIf Trim(txtDeptID) = "" Then
'        MandatoryOK = False
'        MsgBox "Field 'Department' is mandatory. Null value is not allowed", vbInformation, "ComUnion"
'        txtDeptID.SetFocus
'        Exit Function
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
    
    oBar.BitEnabled Me, ITGLedgerMain.tbrMain, lACNew, , , , , , , , , True, , , True
    
    txtUserID.Locked = False
    txtUserID.SetFocus
    
    Mode = Find
End Sub







