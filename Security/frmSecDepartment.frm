VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{B3FB64BF-91F9-11D7-A482-0008A14158BC}#2.22#0"; "ITGControls.ocx"
Begin VB.Form frmSecDepartment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Security - Department"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSecDepartment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2085
   ScaleWidth      =   7695
   Begin TabDlg.SSTab SSTab1 
      Height          =   1695
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   2990
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
      TabPicture(0)   =   "frmSecDepartment.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "List"
      TabPicture(1)   =   "frmSecDepartment.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dtgList"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   1200
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   7320
         Begin VB.Timer Timer1 
            Interval        =   300
            Left            =   3135
            Top             =   240
         End
         Begin ITGControls.ITGTextBox txtDeptID 
            Height          =   285
            Left            =   240
            TabIndex        =   0
            Top             =   315
            Width           =   2880
            _ExtentX        =   4868
            _ExtentY        =   503
            SendKeysTab     =   -1  'True
            BackColor       =   14745599
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
            Label           =   "Department ID"
            BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LabelWidth      =   1410
            TextBoxWidth    =   1410
         End
         Begin ITGControls.ITGTextBox txtDescription 
            Height          =   285
            Left            =   240
            TabIndex        =   1
            Top             =   675
            Width           =   6840
            _ExtentX        =   11853
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
            Label           =   "Description"
            BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            LabelWidth      =   1410
            TextBoxWidth    =   5370
         End
      End
      Begin MSDataGridLib.DataGrid dtgList 
         Height          =   1155
         Left            =   -74880
         TabIndex        =   4
         Top             =   420
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   2037
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
            DataField       =   "DeptID"
            Caption         =   "Department ID"
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
            DataField       =   "Description"
            Caption         =   "Description"
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
               ColumnWidth     =   2174.74
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   4619.906
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar sbRS 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1830
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
End
Attribute VB_Name = "frmSecDepartment"
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
Public Enum eSecDeptMode
    Normal
    AddNewEdit
    Find
End Enum
Private Mode As eSecDeptMode

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

Private Function GetSearchString() As String
On Error GoTo ErrorHandler
Dim sWhere As String

    GetSearchString = True

    If Trim$(txtDeptID.Text) <> "" Then
        sWhere = sWhere & IIf(Trim$(sWhere) = "", "", " AND ") & " DeptID LIKE '" & Trim$(txtDeptID.Text) & "%'"
    End If

    If Trim$(txtDescription.Text) <> "" Then
        sWhere = sWhere & IIf(Trim$(sWhere) = "", "", " AND ") & " Description LIKE '" & Trim$(txtDescription.Text) & "%'"
    End If
    
    GetSearchString = Trim$(sWhere)

    Exit Function
ErrorHandler:
    GetSearchString = "ERROR"
End Function

Private Sub dtgList_HeadClick(ByVal ColIndex As Integer)
    SortGrid dtgList, ColIndex, rsHeader
End Sub

'Set Your Object
Private Sub Form_Load()
    
    oBar.AcessBit Me, GetValueFrTable("AccessLevel", "SEC_ACCESSLEVEL", "RoleID = '" & SecUserRole & "' AND [Module] = 'SDEPT'")
    
    Set rsHeader = New ADODB.Recordset
    Set oNavRec = New clsNavRec
    
    Set FrmName = Me
    oFormSetup.FormLocking True
    oFormSetup.FormSearch True
    oBar.BitEnabled Me, ITGLedgerMain.tbrMain, lACNew, , , , , , , , , True, , , True
    oBar.BitVisible ITGLedgerMain.tbrMain
    ITGLedgerMain.tbrMain.Buttons("btnFind").ButtonMenus("btnFindP").Enabled = True

    Mode = Find
    txtDeptID.Locked = False
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
    
    Set frmSecDepartment = Nothing

    lCloseWindow = True
End Sub

'Add new record to the recordset
Public Sub TBNew()
    
    oBar.BitEnabled Me, ITGLedgerMain.tbrMain, , , , True, True, , , True, True, , , , True
    txtDeptID.BackColor = &HE0FFFF
    Mode = AddNewEdit
    
    If rsHeader.State <> adStateOpen Then
        oRecordset.OpenRecordset rsHeader, "*", "SEC_DEPARTMENT", "WHERE 1 = 0 ", True
        SetDataSource
        SetDataField
    Else
        vBM = rsHeader.Bookmark
    End If
    
    rsHeader.AddNew
    
    Set FrmName = Me
    oFormSetup.FormLocking False
    oFormSetup.ClrRequired &HC0&
    SSTab1.Tab = 0
    txtDeptID.SetFocus

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
    oBar.BitEnabled Me, ITGLedgerMain.tbrMain, lACNew, lACEdit, lACDelete, , , , , , , True, True, , True
    Mode = Normal
    
    MsgBox "Record/s successfully saved.", vbInformation, "ComUnion"
    sbRS.Panels(2) = ""
    
    'Audit trail
    UpdateLogFile "Sec - Department", Trim(txtDeptID), IIf(lBoolean, "Inserted", "Updated")

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
    txtDeptID.Locked = True
    SSTab1.Tab = 0
    txtDescription.SetFocus
    vBM = rsHeader.Bookmark
End Sub

'Delete record
Public Sub TBDelete()
On Error GoTo ErrorHandler
    
    If rsHeader.RecordCount = 0 Then Exit Sub
    
    If MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo, "ComUnion") = vbNo Then Exit Sub
    
    'Audit trail
    UpdateLogFile "Sec - Department", Trim(txtDeptID), "Deleted"

    rsHeader.Delete adAffectCurrent
    rsHeader.UpdateBatch adAffectAll
    
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

'Search using the frmITGSearch
Public Sub TBFind()
    Mode = Normal
    txtDeptID.Locked = True

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
        txtDeptID.Locked = False
        txtDeptID.SetFocus
        Exit Sub
    End If
                
    oRecordset.OpenRecordsetWithCN rsHeader, "*", "SEC_DEPARTMENT", connHeader, IIf(sTemp = "", "", " WHERE " & sTemp), True
                        
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
    txtDeptID.BackColor = &HE0FFFF
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
        txtDeptID.DataField = !DeptID
        txtDescription.DataField = !Description
    End With
End Sub

Private Sub rsHeader_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo ErrorHandler
    
    If Not (rsHeader.EOF) Or Not (rsHeader.BOF) Then
        sbRS.Panels(1) = "Record: " & IIf((rsHeader.AbsolutePosition = -2), "0", rsHeader.AbsolutePosition) & "/" & rsHeader.RecordCount
        
        If rsHeader.Status <> adRecNew Then
            txtDeptID.Locked = True
        Else
            txtDeptID.Locked = False
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
        txtDeptID.Locked = False
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
End Sub

Private Sub txtDeptID_LostFocus()
    'Does Code Already exists
    If Mode = AddNewEdit Then
        If Trim(txtDeptID) = "" Then
            MsgBox "Empty primary input.", vbInformation, "ComUnion"
            txtDeptID.SetFocus
        Else
            If rsHeader.Status <> adRecNew Then Exit Sub
            txtDeptID = Trim(txtDeptID)
            If IDExisting(rsHeader, "DeptID", "SEC_DEPARTMENT", Trim(rsHeader!DeptID), , True) Then
                MsgBox "Department ID already exists.", vbInformation, "ComUnion"
                txtDeptID.SetFocus
            End If
        End If
    End If
End Sub

'Check if all mandatory fields are complete
Function MandatoryOK() As Boolean
    
    MandatoryOK = True
    
    If Trim(txtDeptID) = "" Then
        MandatoryOK = False
        MsgBox "Field 'Department ID' is mandatory. Null value is not allowed.", vbInformation, "ComUnion"
        txtDeptID.SetFocus
        Exit Function
    End If
    
    If rsHeader.Status = adRecNew Then
        If IDExisting(rsHeader, "DeptID", "SEC_DEPARTMENT", Trim(rsHeader!DeptID), , True) Then
            MandatoryOK = False
            MsgBox "Department ID already exists.", vbInformation, "ComUnion"
            txtDeptID.SetFocus
            Exit Function
        End If
    End If
    
    If Trim(txtDescription) = "" Then
        MandatoryOK = False
        MsgBox "Field 'Description' is mandatory. Null value is not allowed", vbInformation, "ComUnion"
        txtDescription.SetFocus
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
    
    oBar.BitEnabled Me, ITGLedgerMain.tbrMain, lACNew, , , , , , , , , True, , , True
    
    txtDeptID.Locked = False
    txtDeptID.SetFocus
    
    Mode = Find
End Sub







