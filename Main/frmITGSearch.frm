VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B3FB64BF-91F9-11D7-A482-0008A14158BC}#2.16#0"; "ITGControls.ocx"
Begin VB.Form frmITGSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9090
   Icon            =   "frmITGSearch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   1980
      Left            =   2850
      ScaleHeight     =   1920
      ScaleWidth      =   6090
      TabIndex        =   23
      Top             =   150
      Width           =   6150
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7980
      Picture         =   "frmITGSearch.frx":0CCA
      TabIndex        =   16
      Top             =   2400
      Width           =   795
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      Caption         =   "Refresh"
      Height          =   375
      Left            =   7200
      Picture         =   "frmITGSearch.frx":1994
      TabIndex        =   14
      Top             =   2400
      Width           =   795
   End
   Begin VB.CommandButton cmdSearch 
      Appearance      =   0  'Flat
      Caption         =   "Search"
      Height          =   375
      Left            =   6420
      Picture         =   "frmITGSearch.frx":265E
      TabIndex        =   15
      Top             =   2400
      Width           =   795
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   5805
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Records: 0"
            TextSave        =   "Records: 0"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2955
      Left            =   60
      TabIndex        =   9
      Top             =   2805
      Width           =   8970
      Begin MSComctlLib.ListView ListView1 
         Height          =   2475
         Left            =   60
         TabIndex        =   18
         Top             =   420
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   4366
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Module"
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Trans #"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Description"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Amount"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "2. Search Result"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   180
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   2685
      Begin ITGControls.ITGDateBox ITGDateBox1 
         Height          =   285
         Left            =   780
         TabIndex        =   4
         Top             =   1290
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
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   2100
         Picture         =   "frmITGSearch.frx":5010
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1635
         Width           =   375
      End
      Begin VB.ComboBox cboModule 
         Height          =   315
         ItemData        =   "frmITGSearch.frx":5CDA
         Left            =   780
         List            =   "frmITGSearch.frx":5CE7
         TabIndex        =   0
         Top             =   510
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   2100
         Picture         =   "frmITGSearch.frx":5D0B
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1290
         Width           =   375
      End
      Begin VB.ComboBox cboRange 
         Height          =   315
         ItemData        =   "frmITGSearch.frx":69D5
         Left            =   780
         List            =   "frmITGSearch.frx":69D7
         TabIndex        =   1
         Top             =   885
         Width           =   1695
      End
      Begin ITGControls.ITGDateBox ITGDateBox2 
         Height          =   285
         Left            =   780
         TabIndex        =   2
         Top             =   1635
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
      End
      Begin ITGControls.ITGTextBox txtFrom 
         Height          =   285
         Left            =   780
         TabIndex        =   20
         Top             =   1290
         Width           =   1290
         _ExtentX        =   2064
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
         Mandatory       =   -1  'True
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
         LabelWidth      =   0
         TextBoxWidth    =   1230
      End
      Begin ITGControls.ITGTextBox txtTo 
         Height          =   285
         Left            =   780
         TabIndex        =   5
         Top             =   1635
         Width           =   1290
         _ExtentX        =   2064
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
         Mandatory       =   -1  'True
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
         LabelWidth      =   0
         TextBoxWidth    =   1230
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "1. Filter Range"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Module"
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
         Left            =   150
         TabIndex        =   10
         Top             =   570
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "To"
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
         Left            =   360
         TabIndex        =   8
         Top             =   1695
         Width           =   180
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "From"
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
         Left            =   195
         TabIndex        =   7
         Top             =   1305
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Range"
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
         TabIndex        =   6
         Top             =   930
         Width           =   465
      End
   End
   Begin ITGControls.ITGTextBox txtFind 
      Height          =   285
      Left            =   75
      TabIndex        =   21
      Top             =   2460
      Width           =   5730
      _ExtentX        =   9895
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
      TextBoxWidth    =   5670
   End
   Begin VB.Label Label7 
      Caption         =   "Containing text:"
      Height          =   255
      Left            =   105
      TabIndex        =   22
      Top             =   2205
      Width           =   1155
   End
End
Attribute VB_Name = "frmITGSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'IT Group Inc. 2005.09.23

Option Explicit

Private oConnection As New clsConnection
Private connList As ADODB.Connection

'Object variables
Private oBar As New clsToolBarMenuBit

'Other declaration
Public sBit As String

Private Sub cboModule_Click()
    Select Case cboModule
        Case Is = "All"
            cboRange.Clear
            cboRange.AddItem Trim("Date")
        Case Is = "Sales Order"
            cboRange.Clear
            cboRange.AddItem Trim("Date")
            cboRange.AddItem Trim("Customer")
            cboRange.AddItem Trim("Salesman")
        Case Is = "Product File"
            cboRange.Clear
            cboRange.AddItem Trim("ItemNo")
            cboRange.AddItem Trim("Unit")
            cboRange.AddItem Trim("Description")
    End Select
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdSearch_Click()
    Dim cmdFind As ADODB.Command
    Dim rsFind As ADODB.Recordset
    
    Set cmdFind = New ADODB.Command
    Set rsFind = New ADODB.Recordset
    With cmdFind
        .ActiveConnection = cn
        .CommandTimeout = 1000
        .CommandText = "USP_ITGSearch"
        .CommandType = adCmdStoredProc
        .Parameters("@cCompanyID") = COID
        .Parameters("@cSearchItem") = Trim(txtFind.Text)
    End With

    Set rsFind = cmdFind.Execute()
    ListView1.ListItems.Clear
    If rsFind.RecordCount <> 0 Then
        While Not rsFind.EOF
            Set itmX = ListView1.ListItems.Add(, , rsFind!cModule & "")
            itmX.SubItems(1) = rsFind!cTranNo & ""
            itmX.SubItems(2) = Format(rsFind!dDate, "mm/dd/yyyy")
            itmX.SubItems(3) = rsFind!cDesc & ""
            itmX.SubItems(4) = Format(rsFind!nAmount, "#,##0.#0")
            rsFind.MoveNext
        Wend
    End If

    rsFind.Close
    
    Set rsFind = Nothing
    Set cmdFind = Nothing
End Sub

Private Sub Form_Load()
    cboModule = "All"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    oBar.BitEnabled Me, ITGLedgerMain.tbrMain, , , , , , , , , , True
    oBar.BitVisible ITGLedgerMain.tbrMain
    ITGLedgerMain.tbrMain.Buttons("btnFind").ButtonMenus("btnFindP").Enabled = False

    Me.Tag = "Close"
    Set frmITGSearch = Nothing

    lCloseWindow = True
End Sub

Public Sub TBCloseWindow()
    Unload Me
End Sub

'Reload menu buttons (do not delete this sub)
Public Sub TBBitReload()
    oBar.BitVisible ITGLedgerMain.tbrMain, True, True
    ITGLedgerMain.tbrMain.Buttons("btnFind").ButtonMenus("btnFindP").Enabled = False
    oBar.BitReload Me, ITGLedgerMain.tbrMain, sBit
    Set FrmName = Me
End Sub

Private Sub Form_Activate()
    oBar.BitVisible ITGLedgerMain.tbrMain
    oBar.BitReload Me, ITGLedgerMain.tbrMain, sBit
    ITGLedgerMain.tbrMain.Buttons("btnFind").ButtonMenus("btnFindP").Enabled = False
    Set FrmName = Me
End Sub

Private Sub ListView1_DblClick()
Dim sModule As String
If ListView1.ListItems.Count = 0 Then Exit Sub

    Select Case Me.ListView1.SelectedItem
        Case "Product"
            Set FrmName = frmMaintProduct
            sModule = "Product File"
        Case "Customer"
            Set FrmName = frmMaintCustomer
            sModule = "Customer File"
        Case "Salesman"
            Set FrmName = frmMaintSalesman
            sModule = "Salesman File"
        Case "Supplier"
            Set FrmName = frmMaintCustomerSupplier
            sModule = "Supplier File"
    End Select
    If (sModule & "") <> "" Then
        With FrmName
            If .Mode = 1 Then
                MsgBox "Sytem detected that '" & Trim(sModule) & "' is in Add/Edit mode." & _
                        vbCr & "Please save/undo the current transaction.", vbCritical + vbOKOnly, "Comunion"
                FrmName.SetFocus
                Exit Sub
            Else
                .Show
                .ZOrder
                .ShowForm Trim(ListView1.SelectedItem.SubItems(1))
            End If
        End With
    End If
End Sub
