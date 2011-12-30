VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B3FB64BF-91F9-11D7-A482-0008A14158BC}#2.22#0"; "ITGControls.ocx"
Begin VB.Form frmSysStructure 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auto Numbering Setup"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10005
   Icon            =   "frmSysStructure.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   10005
   Begin VB.Frame Frame2 
      Height          =   5310
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   9900
      Begin VB.Frame Frame1 
         Caption         =   "Format Legend"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   5430
         TabIndex        =   4
         Top             =   3780
         Width           =   3435
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#/0 - Numerical"
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
            Left            =   240
            TabIndex        =   8
            Top             =   300
            Width           =   1110
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mm - Month"
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
            Left            =   240
            TabIndex        =   7
            Top             =   570
            Width           =   840
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "dd - Day"
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
            Left            =   1920
            TabIndex        =   6
            Top             =   285
            Width           =   615
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "yy/yyyy - Year"
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
            Left            =   1905
            TabIndex        =   5
            Top             =   570
            Width           =   1080
         End
      End
      Begin VB.ComboBox cbommyear 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6945
         TabIndex        =   2
         Top             =   2010
         Width           =   1200
      End
      Begin VB.ComboBox cboResetCtr 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6960
         TabIndex        =   1
         Top             =   3120
         Width           =   1170
      End
      Begin MSComctlLib.ListView lv 
         Height          =   4980
         Left            =   90
         TabIndex        =   3
         Top             =   210
         Width           =   4920
         _ExtentX        =   8678
         _ExtentY        =   8784
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   6174
         EndProperty
      End
      Begin ITGControls.ITGXPCheckBox cboTrans 
         Height          =   210
         Left            =   5430
         TabIndex        =   9
         Top             =   930
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   370
         BorderColor     =   16243138
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Transactional Module"
      End
      Begin ITGControls.ITGXPCheckBox cboAutoNo 
         Height          =   210
         Left            =   5430
         TabIndex        =   10
         Top             =   1290
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   370
         BorderColor     =   16243138
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Auto Numbering"
      End
      Begin ITGControls.ComunionButton cmdOK 
         Height          =   345
         Left            =   5430
         TabIndex        =   11
         Top             =   4815
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         BTYPE           =   14
         TX              =   "O&k"
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
         BCOL            =   15724527
         BCOLO           =   15724527
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmSysStructure.frx":0CCA
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
         Height          =   345
         Left            =   6540
         TabIndex        =   12
         Top             =   4815
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         BTYPE           =   14
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
         BCOL            =   15724527
         BCOLO           =   15724527
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmSysStructure.frx":0CE6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ITGControls.ComunionButton cmdApply 
         Height          =   345
         Left            =   7650
         TabIndex        =   13
         Top             =   4815
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         BTYPE           =   14
         TX              =   "&Apply"
         ENAB            =   0   'False
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
         BCOL            =   15724527
         BCOLO           =   15724527
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmSysStructure.frx":0D02
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ITGControls.ITGTextBox txtAlphaVal 
         Height          =   285
         Left            =   5415
         TabIndex        =   14
         Top             =   1650
         Width           =   2715
         _ExtentX        =   4577
         _ExtentY        =   503
         BorderColor     =   16243138
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Label           =   "Alphabet Value"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextBoxWidth    =   1155
      End
      Begin ITGControls.ITGTextBox txtCounter 
         Height          =   285
         Left            =   8160
         TabIndex        =   15
         Top             =   3150
         Width           =   1575
         _ExtentX        =   2566
         _ExtentY        =   503
         BorderColor     =   16243138
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
         Label           =   "Counter"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelWidth      =   700
         TextBoxWidth    =   815
      End
      Begin ITGControls.ITGXPCheckBox chkline 
         Height          =   210
         Left            =   8250
         TabIndex        =   16
         Top             =   1695
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   370
         BorderColor     =   16243138
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Line"
      End
      Begin ITGControls.ITGTextBox txtnumeric 
         Height          =   285
         Left            =   5430
         TabIndex        =   17
         Top             =   2385
         Width           =   2715
         _ExtentX        =   4577
         _ExtentY        =   503
         BorderColor     =   16243138
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Label           =   "Numerical"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextBoxWidth    =   1155
      End
      Begin ITGControls.ITGTextBox txtformat 
         Height          =   285
         Left            =   5445
         TabIndex        =   18
         Top             =   2745
         Width           =   4290
         _ExtentX        =   7355
         _ExtentY        =   503
         BorderColor     =   16243138
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Label           =   "Numbering Format"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextBoxWidth    =   2730
      End
      Begin ITGControls.ITGXPCheckBox chkLock 
         Height          =   210
         Left            =   5460
         TabIndex        =   23
         Top             =   3480
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   370
         BorderColor     =   16243138
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Lock Transaction Number"
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reset Counter By"
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
         Left            =   5460
         TabIndex        =   22
         Top             =   3135
         Width           =   1275
      End
      Begin VB.Label lblAutoNumber 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5400
         TabIndex        =   21
         Top             =   255
         Width           =   3390
      End
      Begin VB.Label lblvalue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5400
         TabIndex        =   20
         Top             =   645
         Width           =   3390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month and Year"
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
         Left            =   5445
         TabIndex        =   19
         Top             =   2085
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmSysStructure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

'ADO Connection variables
Private oConnection As New clsConnection
Private connHeader As ADODB.Connection
Private connDetail As ADODB.Connection


Public rs As ADODB.Recordset
Public img As ImageList
Dim sID() As String
Dim nOrderID() As Integer
Dim nTransactional() As Integer
Dim nAutoNo() As Integer
Dim sFormat() As String
Dim sResetCtr() As String
Dim sAlphaVal() As String
Dim nCtr() As Integer
Dim sLine As String
Dim MonYear As String
Dim nIndex As Integer
Public sBit As String

'Public Sub LoadList()
'
'
'
'Dim i As Integer
'Dim sSpace As String
'Dim n As Integer
'Set rs = New ADODB.Recordset
'rs.CursorLocation = adUseClient
'rs.Open "Select * from SYS_STRUCTURE", cn, adOpenKeyset
'
'
'
'    If rs.RecordCount > 0 Then
'
'        ReDim sID(0 To rs.RecordCount) As String
'        ReDim nOrderID(0 To rs.RecordCount) As Integer
'        ReDim nTransactional(0 To rs.RecordCount) As Integer
'        ReDim nAutoNo(0 To rs.RecordCount) As Integer
'        ReDim sFormat(0 To rs.RecordCount) As String
'        ReDim sResetCtr(0 To rs.RecordCount) As String
'        ReDim sAlphaVal(0 To rs.RecordCount) As String
'        ReDim nCtr(0 To rs.RecordCount) As Integer
'
'        i = 0
'        rs.MoveFirst
'        lst.Clear
'
'        Do Until rs.EOF
'            sID(i) = rs!cCode
'            nOrderID(i) = IIf(IsNull(rs!nOrderID), i, rs!nOrderID)
'            nTransactional(i) = IIf(IsNull(rs!lTransactional), 0, rs!lTransactional)
'            nAutoNo(i) = IIf(IsNull(rs!lAutoNo), 0, rs!lAutoNo)
'            sFormat(i) = IIf(IsNull(rs!cNumberFormat), "", rs!cNumberFormat)
'            sResetCtr(i) = IIf(IsNull(rs!cResetCtr), "", rs!cResetCtr)
'            sAlphaVal(i) = IIf(IsNull(rs!cAlphaVal), "", rs!cAlphaVal)
'            nCtr(i) = IIf(IsNull(rs!nCtr), "", rs!nCtr)
'            sSpace = ""
'            For n = 1 To rs!nLevel
'                sSpace = sSpace & "     "
'            Next n
'
'            lst.AddItem sSpace & rs!cCaption
'            i = i + 1
'            rs.MoveNext
'        Loop
'    End If
'End Sub

Private Sub cboAutoNo_Click()
'On Error GoTo ErrHandler
'    cmdApply.Enabled = True
'    nAutoNo(lst.ListIndex) = IIf(cboAutoNo.Value = Checked, 1, 0)
'ErrHandler:
'    If err.Number <> 0 Then MsgBox err.Number & ":" & err.Description
End Sub

Private Sub cbommyear_Change()
    FormatAutoNumber
End Sub

Private Sub cbommyear_Click()
    FormatAutoNumber
    txtnumeric.SetFocus
End Sub

Private Sub cbommyear_GotFocus()
    FormatAutoNumber
End Sub

Private Sub cbommyear_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cbommyear_LostFocus()
    FormatAutoNumber
End Sub

Private Sub cboResetCtr_Change()
On Error GoTo ErrHandler
    cmdApply.Enabled = True
    'sResetCtr(lst.ListIndex) = cboResetCtr.Text
ErrHandler:
    If Err.Number <> 0 Then MsgBox Err.Number & ":" & Err.Description
End Sub

Private Sub cboResetCtr_Click()
On Error GoTo ErrHandler
    cmdApply.Enabled = True
    'sResetCtr(lst.ListIndex) = cboResetCtr.Text
ErrHandler:
    If Err.Number <> 0 Then MsgBox Err.Number & ":" & Err.Description
End Sub

Private Sub cboResetCtr_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cboTrans_Click()
'On Error GoTo ErrHandler
'    cmdApply.Enabled = True
'    nTransactional(lst.ListIndex) = IIf(cboTrans.Value = Checked, 1, 0)
'ErrHandler:
'If err.Number <> 0 Then MsgBox err.Number & ":" & err.Description
End Sub

Private Sub chkline_Click()
If Trim(txtAlphaVal.Text) = "" Then
    If chkline.Value = 1 Then
        chkline.Value = 1
    Else: chkline.Value = 0
    End If
End If
FormatAutoNumber
End Sub



'Private Sub cmdAdd_Click()
'    If sID(lst.ListIndex) <> "" Then
'        frmAddModule.sMode = "Add"
'        frmAddModule.sParent = sID(lst.ListIndex)
'        frmAddModule.Show vbModal, Me
'    End If
'End Sub

Private Sub cmdApply_Click()
On Error GoTo ErrHandler
If cboAutoNo.Value = 1 Then
    If Trim(cbommyear.Text) = "" Then
        MsgBox "Field 'Month and Year' is mandatory. Null value is not allowed.", vbInformation, "ComUnion"
        cbommyear.SetFocus
        Exit Sub
    End If
    If Trim(txtnumeric) = "" Then
        MsgBox "Field 'Numeric' is mandatory. Null value is not allowed.", vbInformation, "ComUnion"
        txtnumeric.SetFocus
        Exit Sub
    End If
    If Trim(cboResetCtr.Text) = "" Then
        MsgBox "Field 'Reset Counter By' is mandatory. Null value is not allowed.", vbInformation, "ComUnion"
        cboResetCtr.SetFocus
        Exit Sub
    End If
'    If Len(Trim(txtnumeric)) < 4 Then
'        MsgBox "The Minimum length of numerical digit is six(6).", vbInformation, "ComUnion"
'        txtnumeric.SetFocus
'        SendKeys "{HOME} + {END}"
'        Exit Sub
'    End If
    If Len(Trim(txtformat)) > 20 Then
        MsgBox "The Maximum length of Numbering Format is twenty(20).", vbInformation, "ComUnion"
        Exit Sub
    End If
End If


    UpdateChange
    'Audit trail
    UpdateLogFile "System Auto Number Setup", Trim(lblAutoNumber.Caption), "Updated"

'Dim cIndex As Integer
'    cIndex = lst.ListIndex
'    If UpdateChanges(cIndex) = True Then
'        lst.ListIndex = cIndex
'        cmdApply.Enabled = False
'    Else
'        MsgBox "Unable to update."
'    End If
ErrHandler:
If Err.Number <> 0 Then MsgBox Err.Number & ":" & Err.Description
End Sub

Private Sub UpdateChange()
Dim i As Integer
Dim sql As String
On Error GoTo ErrHandler
sql = "UPDATE system_autonumber SET lTransactional=" & cboTrans.Value & _
        ", lAutoNo =" & cboAutoNo.Value & ", cNumberFormat= '" & txtformat.Text & "',cResetCtr ='" & _
        cboResetCtr.Text & "', cAlphaVal ='" & txtAlphaVal.Text & "', nCtr = " & txtCounter.Text & _
        " , cAutoNumberLabel = '" & lblAutoNumber.Caption & "',  cValue = '" & lblvalue.Caption & _
        "' , cMMYYYY = '" & cbommyear.Text & "', lLine = " & chkline.Value & " , lLock = " & chkLock.Value & " , cNumeric = '" & txtnumeric.Text & _
        "' WHERE cCode = '" & lv.SelectedItem & "' AND cCompanyID = '" & COID & "' "
        cn.Execute sql
        MsgBox "Successfully Updated."
ErrHandler:
    If Err.Number <> 0 Then MsgBox Err.Description
End Sub

'Private Function UpdateChanges(curIndex As Integer) As Boolean
'Dim i As Integer
''Dim oCN As New clsDatabase
'On Error GoTo ErrHandler
'    For i = 0 To lst.ListCount - 1
'        cn.Execute "UPDATE SYS_STRUCTURE SET lTransactional=" & nTransactional(i) & _
'        ", lAutoNo =" & nAutoNo(i) & ", cNumberFormat='" & sFormat(i) & "',cResetCtr ='" & _
'        sResetCtr(i) & "', nOrderID=" & i & ", cAlphaVal ='" & sAlphaVal(i) & "', nCtr = " & nCtr(i) & " WHERE cCode='" & sID(i) & "'"
'    Next i
'    UpdateChanges = True
'    Set cn = Nothing
'ErrHandler:
'    If err.Number <> 0 Then UpdateChanges = False
'End Function

Private Sub cmdcancel_Click()
    Unload Me
End Sub

'Private Sub cmdDown_Click()
'Dim sText As String
'Dim iIndex As Integer, Ctr As Integer
'Dim Order As Integer, ID As String, Trans As Integer, Auto As Integer, sForm As String, sReset As String, sAlpha As String
'cmdApply.Enabled = True
'If lst.SelCount = 1 Then
'    If lst.ListCount - 1 = lst.ListIndex Then Exit Sub
'    sText = lst.List(lst.ListIndex)
'    iIndex = lst.ListIndex
'    lst.RemoveItem lst.ListIndex
'    lst.AddItem sText, iIndex + 1
'
'    Order = nOrderID(iIndex + 1)
'    nOrderID(iIndex + 1) = nOrderID(iIndex)
'    nOrderID(iIndex) = Order
'
'    ID = sID(iIndex + 1)
'    sID(iIndex + 1) = sID(iIndex)
'    sID(iIndex) = ID
'
'    Trans = nTransactional(iIndex + 1)
'    nTransactional(iIndex + 1) = nTransactional(iIndex)
'    nTransactional(iIndex) = Trans
'
'    Auto = nAutoNo(iIndex + 1)
'    nAutoNo(iIndex + 1) = nAutoNo(iIndex)
'    nAutoNo(iIndex) = Auto
'
'    sForm = sFormat(iIndex + 1)
'    sFormat(iIndex + 1) = sFormat(iIndex)
'    sFormat(iIndex) = sForm
'
'    sReset = sResetCtr(iIndex + 1)
'    sResetCtr(iIndex + 1) = sResetCtr(iIndex)
'    sResetCtr(iIndex) = sReset
'
'    sAlpha = sAlphaVal(iIndex + 1)
'    sAlphaVal(iIndex + 1) = sAlphaVal(iIndex)
'    sAlphaVal(iIndex) = sAlpha
'
'    Ctr = nCtr(iIndex + 1)
'    nCtr(iIndex + 1) = nCtr(iIndex)
'    nCtr(iIndex) = Ctr
'
'    lst.Selected(iIndex + 1) = True
'End If
'End Sub

'Private Sub cmdEdit_Click()
'    If sID(lst.ListIndex) <> "" Then
'        frmAddModule.sMode = "Edit"
'        frmAddModule.sName = sID(lst.ListIndex)
'        frmAddModule.Show vbModal, Me
'    End If
'End Sub

Private Sub cmdOK_Click()
    If cmdApply.Enabled = True Then Call cmdApply_Click
    Call cmdcancel_Click
End Sub


'Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 27 Then
'        Call cmdCancel_Click
'    End If
'End Sub

Private Sub Form_Unload(Cancel As Integer)
'    oBar.BitEnabled Me, ITGLedgerMain.tbrMain, , , , , , , , , , True
'    oBar.BitVisible ITGLedgerMain.tbrMain
'    oBar.BitEnabled ITGLedgerMain, Me, ITGLedgerMain.tbrMain, , , , , , , , , , True

    Set frmSysStructure = Nothing
    
'    lCloseWindow = True
'    CloseMenuTab ITGLedgerMain
    
'    'oForm(Me.Tag).Mode = 1
'    'oForm(Me.Tag).Tag = Me.Tag
End Sub


Private Sub Form_Load()
    Set FrmName = Me
'    FormSetup
'    oBar.BitEnabled ITGLedgerMain, Me, ITGLedgerMain.tbrMain, , , , , , , , , , , , , True
    
    cboResetCtr.AddItem "NONE"
    cboResetCtr.AddItem "DAILY"
    cboResetCtr.AddItem "MONTHLY"
    cboResetCtr.AddItem "ANNUALY"
    
    cbommyear.AddItem "NONE"
    cbommyear.AddItem "mmddyy"
    cbommyear.AddItem "mmyy"
    cbommyear.AddItem "mmyyyy"
    cbommyear.AddItem "yy"
    cbommyear.AddItem "yyyy"
    txtCounter.Locked = False
    FillList
End Sub

Public Sub TBCloseWindow()
    Unload Me
End Sub

Public Sub TBBitReload()
    Set FrmName = Me
    oBar.BitVisible ITGLedgerMain.tbrMain, True, True
    sBit = "0000000000000000"
    oBar.BitReload Me, ITGLedgerMain.tbrMain, sBit
End Sub


Sub FillList()
Set rs = Nothing
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open "SELECT cCode, cAutonumberLabel FROM system_autonumber where cCompanyID = '" & COID & "' ORDER BY cAutonumberLabel ASC", cn, adOpenKeyset
If Not rs.EOF Then
    While Not rs.EOF
        lv.ListItems.Add , , rs!cCode
        lv.ListItems(lv.ListItems.Count).ListSubItems.Add , , rs!cAutoNumberLabel
        lv.Refresh
    rs.MoveNext
    Wend
End If
End Sub


Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
Set rs = Nothing
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open "SELECT * FROM system_autonumber where cCompanyID = '" & COID & "' AND cCode = '" & lv.SelectedItem & "'", cn, adOpenKeyset, adLockPessimistic
    If Not rs.EOF Then
        lblAutoNumber.Caption = IIf(IsNull(rs!cAutoNumberLabel), "", rs!cAutoNumberLabel)
        lblvalue.Caption = IIf(IsNull(rs!cValue), "", rs!cValue)
        If (IsNull(rs!lTransactional) = True) Then
            cboTrans.Value = 0
        Else
            If rs!lTransactional = True Then
                cboTrans.Value = 1
            Else: cboTrans.Value = 0
            End If
        End If
        
        If (IsNull(rs!lAutoNo) = True) Then
            cboAutoNo.Value = 0
        Else
            If rs!lAutoNo = True Then
                cboAutoNo.Value = 1
            Else: cboAutoNo.Value = 0
            End If
        End If
        If (IsNull(rs!lLine) = True) Then
            chkline.Value = 0
        Else
            If rs!lLine = True Then
                chkline.Value = 1
            Else: chkline.Value = 0
            End If
        End If
        
        If (IsNull(rs!lLock) = True) Then
            chkLock.Value = 0
        Else
            If rs!lLock = True Then
                chkLock.Value = 1
            Else: chkLock.Value = 0
            End If
        End If
        txtAlphaVal.Text = IIf(IsNull(rs!cAlphaVal), "", rs!cAlphaVal)
        cbommyear.Text = IIf(IsNull(rs!cMMYYYY), "", rs!cMMYYYY)
        txtnumeric.Text = IIf(IsNull(rs!cNumeric), "", rs!cNumeric)
        cboResetCtr.Text = IIf(IsNull(rs!cResetCtr), "", rs!cResetCtr)
        txtCounter.Text = IIf(IsNull(rs!nCtr), "", rs!nCtr)
        txtformat.Text = IIf(IsNull(rs!cNumberFormat), "", rs!cNumberFormat)
    End If
End Sub

Private Sub txtAlphaVal_Change()
    cmdApply.Enabled = True
    'sAlphaVal(lst.ListIndex) = txtAlphaVal.Text
    FormatAutoNumber
End Sub
Private Sub txtCounter_Change()
    cmdApply.Enabled = True
    'nCtr(lst.ListIndex) = txtCounter.Text
End Sub
Private Sub txtFormatNo_Change()
    cmdApply.Enabled = True
    'sFormat(lst.ListIndex) = txtFormatNo.Text
End Sub

Sub FormatAutoNumber()
    If chkline.Value = 1 Then
        If Trim(txtAlphaVal.Text) <> "" Then
            sLine = "-"
        End If
    Else: sLine = ""
    End If
    
    If cbommyear.Text = "NONE" Then
        MonYear = ""
    Else: MonYear = cbommyear.Text
    End If
     txtformat.Text = txtAlphaVal.Text & sLine & MonYear & txtnumeric.Text
End Sub

Private Sub txtnumeric_Change()
FormatAutoNumber
End Sub


Private Sub txtnumeric_KeyPress(KeyAscii As Integer)
Dim str As String
str = "#0"
Dim x As Integer
x = InStr(1, str, Chr$(KeyAscii))
If x = 0 And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub
