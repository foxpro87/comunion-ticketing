VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{B3FB64BF-91F9-11D7-A482-0008A14158BC}#2.22#0"; "ITGControls.ocx"
Begin VB.Form frmSecCompanyAccess 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Security - User Company Access"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7770
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSecCompanyAccess.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3810
   ScaleWidth      =   7770
   Begin TabDlg.SSTab SSTab1 
      Height          =   3435
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   6059
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabMaxWidth     =   3528
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Main"
      TabPicture(0)   =   "frmSecCompanyAccess.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "List"
      TabPicture(1)   =   "frmSecCompanyAccess.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dtgList"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   2940
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   7380
         Begin VB.ComboBox cboCA 
            Height          =   315
            ItemData        =   "frmSecCompanyAccess.frx":0D02
            Left            =   240
            List            =   "frmSecCompanyAccess.frx":0D0C
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   2400
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Timer Timer1 
            Interval        =   300
            Left            =   960
            Top             =   660
         End
         Begin MSDataGridLib.DataGrid dtgCA 
            Height          =   1575
            Left            =   240
            TabIndex        =   1
            Top             =   1140
            Width           =   6855
            _ExtentX        =   12091
            _ExtentY        =   2778
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
            Appearance      =   0
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
            ColumnCount     =   2
            BeginProperty Column00 
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
            BeginProperty Column01 
               DataField       =   "cCompanyName"
               Caption         =   "Company Name"
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
                  Button          =   -1  'True
                  Locked          =   -1  'True
                  ColumnWidth     =   1470.047
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  ColumnWidth     =   5054.74
               EndProperty
            EndProperty
         End
         Begin ITGControls.ITGTextBox txtName 
            Height          =   285
            Left            =   2880
            TabIndex        =   6
            Top             =   300
            Width           =   4200
            _ExtentX        =   7303
            _ExtentY        =   503
            SendKeysTab     =   -1  'True
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Label           =   "ITGtext"
            BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LabelWidth      =   0
            TextBoxWidth    =   4140
            Enabled         =   0   'False
         End
         Begin ITGControls.ITGTextBox txtUserID 
            Height          =   285
            Left            =   240
            TabIndex        =   0
            Top             =   300
            Width           =   2565
            _ExtentX        =   4313
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
            LabelWidth      =   1250
            TextBoxWidth    =   1255
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000000&
            X1              =   780
            X2              =   7200
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Details"
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
            Index           =   2
            Left            =   120
            TabIndex        =   5
            Top             =   720
            Width           =   585
         End
      End
      Begin MSDataGridLib.DataGrid dtgList 
         Height          =   2865
         Left            =   -74880
         TabIndex        =   7
         Top             =   435
         Width           =   7380
         _ExtentX        =   13018
         _ExtentY        =   5054
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         Appearance      =   0
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
               ColumnWidth     =   1365.165
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
   End
   Begin MSComctlLib.StatusBar sbRS 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   3555
      Width           =   7770
      _ExtentX        =   13705
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
End
Attribute VB_Name = "frmSecCompanyAccess"
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
Private rsDetail As ADODB.Recordset

'Form mode enumeration
Enum eSecCAMode
    Normal
    AddNewEdit
    Find
End Enum
Public Mode As eSecCAMode

'Other declarations
Public dtgName As String
Public sBit As String
Private vBM As Variant 'Recordset bookmark variable

'Security Acess Level variables
Public lACNew As Boolean
Public lACEdit As Boolean
Public lACDelete As Boolean
Public lACPost As Boolean
Public lACCancel As Boolean
Public lACPrint As Boolean

Private Sub cboCA_Click()
    rsDetail!cCompanyName = GetValueFrTable("cCompanyName", "COMPANY", "cCompanyID = '" & Trim(cboCA) & "'", True)
    rsDetail!cCompanyID = cboCA
End Sub

Private Sub cboCA_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If dtgCA.Col = 0 Then
            rsDetail!cCompanyID = cboCA
            cboCA.Visible = False
            TBNewLine
        End If
    ElseIf KeyCode = vbKeyEscape Then
        cboCA.Visible = False
    End If
End Sub

Private Sub cboCA_LostFocus()
    cboCA.Visible = False
End Sub

Private Sub dtgCA_ButtonClick(ByVal ColIndex As Integer)
On Error Resume Next
    If Mode <> AddNewEdit Then Exit Sub
    If rsDetail.RecordCount < 1 Then Exit Sub
    Select Case ColIndex
        Case 0
            Call LoadComboValues(cboCA, "cCompanyID", "COMPANY")
            MoveCombo cboCA, dtgCA, dtgCA.Columns(0)
            ComboLoadValue cboCA, Trim(dtgCA.Columns(0).Text)
            cboCA.SetFocus
    End Select
End Sub

'Set the datagrid as active control
Private Sub dtgCA_Click()
    If Mode = AddNewEdit Then dtgName = dtgCA.Name
End Sub

Private Sub dtgCA_Error(ByVal DataError As Integer, Response As Integer)
    If DataError = 7007 Then
        MsgBox "Type mismatch", vbExclamation, "ComUnion"
    ElseIf DataError = 13 Then
        MsgBox "Type mismatch", vbExclamation, "ComUnion"
    End If
    Response = 0
End Sub

Private Sub dtgCA_GotFocus()
    dtgName = dtgCA.Name
End Sub

Private Sub dtgCA_KeyDown(KeyCode As Integer, Shift As Integer)
    If Mode <> AddNewEdit Then Exit Sub
    If (Shift = vbCtrlMask And KeyCode = 45) Then
        TBNewLine
    ElseIf (Shift = vbCtrlMask And KeyCode = 46) Then
        TBDeleteLine
    ElseIf (Shift = vbCtrlMask And KeyCode = 83) Then
        TBSave
    End If
End Sub

Private Sub dtgCA_KeyPress(KeyAscii As Integer)
    If Mode <> AddNewEdit Then Exit Sub
    If rsDetail.RecordCount = 0 Then Exit Sub
    
    If KeyAscii = 39 Then KeyAscii = 0 'Apostrophe {'}
    
    If KeyAscii = 13 Then
        Select Case dtgCA.Col
            Case 0
                If Not cboCA.Visible Then dtgCA_ButtonClick 0
        End Select
    End If
End Sub

'Right click menu popup
Private Sub dtgCA_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Mode <> AddNewEdit Then Exit Sub
    If Button = 2 Then
        dtgName = dtgCA.Name
        'PopupMenu ITGLedgerMain.mnuDetail
    End If
End Sub

Private Sub dtgList_HeadClick(ByVal ColIndex As Integer)
    SortGrid dtgList, ColIndex, rsHeader
End Sub

'Set Your Object
Private Sub Form_Load()

    oBar.AcessBit Me, GetValueFrTable("AccessLevel", "SEC_ACCESSLEVEL", "RoleID = '" & SecUserRole & "' AND [Module] = 'SCACCESS'")
    
    Set rsHeader = New ADODB.Recordset
    Set rsDetail = New ADODB.Recordset
    Set oNavRec = New clsNavRec

    Set FrmName = Me
    oFormSetup.FormLocking True
    oBar.BitEnabled Me, ITGLedgerMain.tbrMain, , , , , , , , , , True, , , True
    oBar.BitVisible ITGLedgerMain.tbrMain, True
    ITGLedgerMain.tbrMain.Buttons("btnFind").ButtonMenus("btnFindP").Enabled = True

    Mode = Find
    txtUserID.Locked = False
    
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
    Set rsDetail = Nothing
    
    Set frmSecCompanyAccess = Nothing

    lCloseWindow = True
End Sub

'Add new record to the recordset
Public Sub TBNew()

    '**********

End Sub

'Undo all changes to the recordset
Public Sub TBUndoAll()
On Error GoTo ErrorHandler

    Mode = Normal
    
    If rsHeader.Status = adRecNew Then TBUndoCurrent
    
    rsHeader.CancelBatch adAffectAll
    rsDetail.CancelBatch adAffectAll
    
    oRecordset.UnbindControls
    
    If rsHeader.RecordCount <> 0 Then rsHeader.Bookmark = vBM
    
    Set FrmName = Me
    oFormSetup.FormLocking True
    
    If rsHeader.RecordCount <> 0 Then
        oBar.BitEnabled Me, ITGLedgerMain.tbrMain, , lACEdit, , , , , , , , True, True, , True
    Else
        RSZero
        Mode = Find
    End If

    sbRS.Panels(2) = ""
    cboCA.Visible = False
    
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

    GetChild
    If rsHeader.Status = adRecNew Then
        rsDetail.CancelBatch adAffectAll
        rsHeader.CancelUpdate
    Else
        rsHeader.CancelBatch adAffectCurrent
        rsDetail.CancelBatch adAffectAll
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
    rsDetail.UpdateBatch adAffectAll
    
    cn.CommitTrans
    OKUpdate = True
    
    cboCA.Visible = False
    Set FrmName = Me
    oFormSetup.FormLocking True
    oBar.BitEnabled Me, ITGLedgerMain.tbrMain, , lACEdit, , , , , , , , True, True, , True
    Mode = Normal
    
    MsgBox "Record/s successfully saved.", vbInformation, "ComUnion"
    sbRS.Panels(2) = ""
    
    'Audit trail
    UpdateLogFile "Sec - Company Access", Trim(txtUserID), IIf(lBoolean, "Inserted", "Updated")

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
        rsDetail.Requery
        oBar.BitEnabled Me, ITGLedgerMain.tbrMain, , lACEdit, , , , , , , , True, True, , True
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
    Mode = AddNewEdit
    oBar.BitEnabled Me, ITGLedgerMain.tbrMain, , , , True, True, , , True, True, , , , True
    Set FrmName = Me
    oFormSetup.FormLocking False
    oFormSetup.ClrRequired &HC0&
    txtUserID.Locked = True
    dtgCA.SetFocus
    vBM = rsHeader.Bookmark
End Sub

'Delete record
Public Sub TBDelete()
On Error GoTo ErrorHandler

ErrorHandler:
    If Err.Number = -2147217885 Then
        Resume Next
    End If

End Sub

'Search using the frmITGSearch
Public Sub TBFind()
    Mode = Normal
    txtUserID.Locked = True

End Sub

'Search using the recordset primary key
Public Sub TBFindPrimary()
    RSZero
    Mode = Find
End Sub

'Reload menu buttons (do not delete this sub)
Public Sub TBBitReload()
    oBar.BitVisible ITGLedgerMain.tbrMain, True
    ITGLedgerMain.tbrMain.Buttons("btnFind").ButtonMenus("btnFindP").Enabled = True
    oBar.BitReload Me, ITGLedgerMain.tbrMain, sBit
    Set FrmName = Me
    dtgName = dtgCA.Name
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

'Move previuos
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
    
    If dtgName = "dtgCA" Then
        rsDetail.AddNew
        rsDetail!UserID = Trim(rsHeader!UserID)
        dtgCA.Col = 0
        dtgCA.Columns(0).Value = ""
        GetChild
        If rsDetail.RecordCount <> 0 Then rsDetail.MoveLast
        dtgCA.SetFocus
    End If
    
    oBar.BitEnabled Me, ITGLedgerMain.tbrMain, , , , True, True, , , True, True, , , , True
End Sub

'Delete line in the detail recordset
Public Sub TBDeleteLine()
On Error GoTo ErrorHandler

    If dtgName = dtgCA.Name Then
        If rsDetail.RecordCount = 0 Then Exit Sub
        vBookMark = dtgCA.Bookmark
        GetChild
        dtgCA.Bookmark = vBookMark
        rsDetail.Delete adAffectCurrent
    End If
    GetChild

    oBar.BitEnabled Me, ITGLedgerMain.tbrMain, , , , True, True, , , True, True, , , , True

ErrorHandler:
    If Err.Number = -2147217885 Then
        Resume Next
    End If

End Sub

'Undo All
Public Sub TBUndoLineAll()
On Error GoTo ErrorHandler

    MsgBox "Unavailable on " & Me.Name

ErrorHandler:
    If Err.Number = -2147217885 Then
        Resume Next
    End If

End Sub

'Undo current line
Public Sub TBUndoLineCurrent()
    MsgBox "Unavailable on " & Me.Name
End Sub

'Post current record
Public Sub TBPostRecord()
    MsgBox "Unavailable on " & Me.Name
End Sub

'Cancel current record
Public Sub TBCancelRecord()
    MsgBox "Unavailable on " & Me.Name
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
    Set dtgCA.DataSource = rsDetail
End Sub

'Sets the data field for every bounded controls
Sub SetDataField()
    With rsHeader
        txtUserID.DataField = !UserID
    End With
End Sub

Private Sub rsHeader_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo ErrorHandler

    If Not (rsHeader.EOF) Or Not (rsHeader.BOF) Then
        'Status bar setup
        sbRS.Panels(1) = "Record: " & IIf((rsHeader.AbsolutePosition = -2), "0", rsHeader.AbsolutePosition) & "/" & rsHeader.RecordCount

        If rsHeader.Status <> adRecNew Then
            txtUserID.Locked = True
        Else
            txtUserID.Locked = False
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
        
        GetChild
    Else
        sbRS.Panels(1) = "Record: 0/0"
        sbRS.Panels(2) = ""
        txtUserID.Locked = False
    End If

    If Mode = AddNewEdit Then
        dtgCA.Refresh
    End If

ErrorHandler:
    'Err.Number -2147217885
    'Description - Row handle referred to a deleted row or a row marked for deletion.
    If Err.Number = -2147217885 Then
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

Private Sub txtUserID_Change()
    txtName = ""
    sSQL = "SELECT (LastName + ', ' + LastName + ' ' + MI + '.') AS Name FROM SEC_USER WHERE UserID = '" & Trim(txtUserID) & "'"
    Set rs = New Recordset
    rs.Open sSQL, cn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount <> 0 Then
        txtName = Trim(rs!Name)
    Else
        txtName = ""
    End If
    Set rs = Nothing
End Sub

Private Sub txtUserID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        'Find
        If Mode = Find Then
            
            FormWaitShow "Loading data . . ."
            
            Set rsHeader = Nothing
            Set rsDetail = Nothing
            Set rsHeader = New ADODB.Recordset
            Set rsDetail = New ADODB.Recordset

            If Trim(txtUserID) = "" Then
                oRecordset.OpenRecordset rsDetail, "*", "SEC_COMPANYACCESS", , True
                oRecordset.OpenRecordset rsHeader, "*", "SEC_USER", , True
            Else
                oRecordset.OpenRecordset rsDetail, "*", "SEC_COMPANYACCESS", "WHERE UserID LIKE '" & Trim(txtUserID) & "%'", True
                oRecordset.OpenRecordset rsHeader, "*", "SEC_USER", "WHERE UserID LIKE '" & Trim(txtUserID) & "%'", True
            End If

            Set FrmName = Me
            oFormSetup.FormLocking True

            If rsHeader.RecordCount = 0 Then
                Unload frmWait
                MsgBox "No matching record/s found.", vbInformation, "ComUnion Search"
                rsHeader.Close
                rsDetail.Close
                oBar.BitEnabled Me, ITGLedgerMain.tbrMain, , , , , , , , , , True, , , True
                txtUserID.Locked = False
                txtUserID.SetFocus
                Exit Sub
            End If

            SetDataSource
            SetDataField

            txtUserID.BackColor = &HE0FFFF
            Mode = Normal

            oBar.BitEnabled Me, ITGLedgerMain.tbrMain, , lACEdit, , , , , , , , True, True, , True

            Unload frmWait
        End If
    End If
End Sub

'Check if all mandatory fields are complete
Function MandatoryOK() As Boolean

    MandatoryOK = True

    GetChild
    
    If rsDetail.RecordCount <> 0 Then rsDetail.MoveFirst
    Do Until rsDetail.EOF
        If rsDetail.Status = (adRecNew) Or rsDetail.Status = (adRecModified) Then
            If IsNull(rsDetail!cCompanyID) Then
                rsDetail.Delete
            ElseIf Trim(rsDetail!cCompanyID) = "" Then
                rsDetail.Delete
            Else
                rsDetail.MoveNext
            End If
        Else
            rsDetail.MoveNext
        End If
    Loop

    GetChild

End Function

'Filter detail recordset to header's primary
Private Sub GetChild()
    rsDetail.Filter = "UserID = '" & Trim(rsHeader!UserID) & "'"
End Sub

'Sets the form if record number is zero
Private Sub RSZero()
    sbRS.Panels(1) = "Record: 0/0"
    sbRS.Panels(2) = ""
    
    Set dtgCA.DataSource = Nothing
    dtgCA.Refresh
    
    Set FrmName = Me
    oRecordset.UnbindControls
    oFormSetup.TextClearing
    oFormSetup.FormLocking True
    
    If rsHeader.State = adStateOpen Then rsHeader.Close
    If rsDetail.State = adStateOpen Then rsDetail.Close
    
    oBar.BitEnabled Me, ITGLedgerMain.tbrMain, , , , , , , , , , True, , , True
    
    txtUserID.Locked = False
    txtUserID.SetFocus
    
    Mode = Find
    
End Sub
