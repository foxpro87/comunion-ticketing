VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{B3FB64BF-91F9-11D7-A482-0008A14158BC}#2.22#0"; "ITGControls.ocx"
Begin VB.Form frmMaintTicketType 
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ticket Type"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   7515
   Begin MSComctlLib.StatusBar sbRS 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2400
      Width           =   7515
      _ExtentX        =   13256
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
      Height          =   2310
      Left            =   30
      TabIndex        =   1
      Top             =   45
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   4075
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
         Height          =   1920
         Left            =   -74940
         TabIndex        =   6
         Top             =   360
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   3387
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
            DataField       =   "cTypeID"
            Caption         =   "Type ID"
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
            DataField       =   "cDescription"
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
         BeginProperty Column02 
            DataField       =   "cKeyStroke"
            Caption         =   "Key Stroke"
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
               ColumnWidth     =   1560.189
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   3899.906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1560.189
            EndProperty
         EndProperty
      End
      Begin ITGControls.ComunionFrames ComunionFrames2 
         Height          =   1860
         Left            =   75
         Top             =   360
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   3281
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
         Begin VB.Timer Timer1 
            Interval        =   300
            Left            =   3105
            Top             =   510
         End
         Begin ITGControls.ITGTextBox txtTypeID 
            Height          =   285
            Left            =   210
            TabIndex        =   2
            Top             =   585
            Width           =   2880
            _ExtentX        =   4868
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
            Label           =   "Type ID"
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
            Left            =   225
            TabIndex        =   3
            Top             =   930
            Width           =   6840
            _ExtentX        =   11853
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
         Begin ITGControls.ITGTextBox txtDiscount 
            Height          =   285
            Left            =   210
            TabIndex        =   4
            Top             =   1290
            Width           =   2940
            _ExtentX        =   4974
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
            Label           =   "Discount"
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
            TextBoxWidth    =   1470
         End
         Begin ITGControls.ITGTextBox txtStroke 
            Height          =   285
            Left            =   4125
            TabIndex        =   5
            Top             =   1275
            Width           =   2940
            _ExtentX        =   4974
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
            Label           =   "Key Stroke"
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
            TextBoxWidth    =   1470
         End
      End
   End
End
Attribute VB_Name = "frmMaintTicketType"
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
Public Enum eMaintTicketType
    Normal
    AddNewEdit
    Find
End Enum
Private Mode As eMaintTicketType

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

    If Trim$(txtTypeID.Text) <> "" Then
        sWhere = sWhere & IIf(Trim$(sWhere) = "", "", " AND ") & " cTypeID LIKE '" & Trim$(txtTypeID.Text) & "%'"
    End If

    If Trim$(txtDescription.Text) <> "" Then
        sWhere = sWhere & IIf(Trim$(sWhere) = "", "", " AND ") & " cDescription LIKE '" & Trim$(txtDescription.Text) & "%'"
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
    Me.Icon = ITGLedgerMain.Icon
    oBar.AcessBit Me, GetValueFrTable("AccessLevel", "SEC_ACCESSLEVEL", "RoleID = '" & SecUserRole & "' AND [Module] = 'MTTYPE'")
    
    Set rsHeader = New ADODB.Recordset
    Set oNavRec = New clsNavRec
    
    Set FrmName = Me
    oFormSetup.FormLocking True
    oFormSetup.FormSearch True
    oBar.BitEnabled Me, ITGLedgerMain.tbrMain, lACNew, , , , , , , , , True, , , True
    oBar.BitVisible ITGLedgerMain.tbrMain
    ITGLedgerMain.tbrMain.Buttons("btnFind").ButtonMenus("btnFindP").Enabled = True

    Mode = Find
    txtTypeID.Locked = False
    
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
    
'    Set frmSecDepartment = Nothing

    lCloseWindow = True
End Sub

'Add new record to the recordset
Public Sub TBNew()
    
    oBar.BitEnabled Me, ITGLedgerMain.tbrMain, , , , True, True, , , True, True, , , , True
    txtTypeID.BackColor = &HE0FFFF
    Mode = AddNewEdit
    
    If rsHeader.State <> adStateOpen Then
        oRecordset.OpenRecordset rsHeader, "*", "TICKET_TYPE", "WHERE 1 = 0 ", True
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
    txtTypeID.SetFocus

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
    UpdateLogFile "Sec - Department", Trim(txtTypeID), IIf(lBoolean, "Inserted", "Updated")

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
    txtTypeID.Locked = True
    SSTab1.ActiveTab = 0
    SSTab1.TabEnabled(1) = False
    txtDescription.SetFocus
    vBM = rsHeader.Bookmark
End Sub

'Delete record
Public Sub TBDelete()
On Error GoTo ErrorHandler
    
    If rsHeader.RecordCount = 0 Then Exit Sub
    
    If MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo, "ComUnion") = vbNo Then Exit Sub
    
    'Audit trail
    UpdateLogFile "Maintenance - Ticket Type", Trim(txtTypeID), "Deleted"

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
    txtTypeID.Locked = True

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
        txtTypeID.Locked = False
        txtTypeID.SetFocus
        Exit Sub
    End If
                
    oRecordset.OpenRecordsetWithCN rsHeader, "*", "TICKET_TYPE", connHeader, IIf(sTemp = "", "", " WHERE " & sTemp), True
                        
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
    txtTypeID.BackColor = &HE0FFFF
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
        txtTypeID.DataField = !cTypeID
        txtDescription.DataField = !cDescription
        txtDiscount.DataField = !cDiscount
        txtStroke.DataField = !cKeyStroke
    End With
End Sub

Private Sub rsHeader_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo ErrorHandler
    
    If Not (rsHeader.EOF) Or Not (rsHeader.BOF) Then
        sbRS.Panels(1) = "Record: " & IIf((rsHeader.AbsolutePosition = -2), "0", rsHeader.AbsolutePosition) & "/" & rsHeader.RecordCount
        
        If rsHeader.Status <> adRecNew Then
            txtTypeID.Locked = True
        Else
            txtTypeID.Locked = False
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
        txtTypeID.Locked = False
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
End Sub

Private Sub txtDiscount_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
    Case 45
    Case 47
    Case 48 To 57
    Case Else
        KeyAscii = 0
End Select
End Sub

Private Sub txtStroke_KeyDown(KeyCode As Integer, Shift As Integer)
    txtStroke = ""
    SendKeys "{TAB}"
    txtStroke = KeyCode
End Sub

Private Sub txtStroke_KeyPress(KeyAscii As Integer)
'    txtStroke = Asc(txtStroke)
End Sub

Private Sub txtStroke_KeyUp(KeyCode As Integer, Shift As Integer)
'    txtStroke = KeyCode
End Sub

Private Sub txtStroke_LostFocus()
    On Error GoTo ErrHandler
    txtStroke = Right(txtStroke, Len(txtStroke) - 1)
ErrHandler:
End Sub

Private Sub txtTypeID_LostFocus()
    'Does Code Already exists
    If Mode = AddNewEdit Then
        If Trim(txtTypeID) = "" Then
            MsgBox "Empty primary input.", vbInformation, "ComUnion"
            txtTypeID.SetFocus
        Else
            If rsHeader.Status <> adRecNew Then Exit Sub
            txtTypeID = Trim(txtTypeID)
            If IDExisting(rsHeader, "cTypeID", "TICKET_TYPE", Trim(rsHeader!cTypeID), , True) Then
                MsgBox "Type ID already exists.", vbInformation, "ComUnion"
                txtTypeID.SetFocus
            End If
        End If
    End If
End Sub

'Check if all mandatory fields are complete
Function MandatoryOK() As Boolean
    
    MandatoryOK = True
    
    If Trim(txtTypeID) = "" Then
        MandatoryOK = False
        MsgBox "Field 'Department ID' is mandatory. Null value is not allowed.", vbInformation, "ComUnion"
        txtTypeID.SetFocus
        Exit Function
    End If
    
    If rsHeader.Status = adRecNew Then
        If IDExisting(rsHeader, "cTypeID", "TICKET_TYPE", Trim(rsHeader!cTypeID), , True) Then
            MandatoryOK = False
            MsgBox "Type ID already exists.", vbInformation, "ComUnion"
            txtTypeID.SetFocus
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
    
    txtTypeID.Locked = False
    txtTypeID.SetFocus
    
    Mode = Find
End Sub









