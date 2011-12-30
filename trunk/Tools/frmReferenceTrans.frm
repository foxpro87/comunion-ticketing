VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{58BF25E9-1B7C-4F57-B713-7324CE826B02}#1.1#0"; "ITGControls.ocx"
Begin VB.Form frmReferenceTrans 
   Caption         =   "Reference Transactions"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7725
   Icon            =   "frmReferenceTrans.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   7725
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Transaction Details"
      Height          =   3270
      Left            =   60
      TabIndex        =   6
      Top             =   1275
      Width           =   7590
      Begin MSComctlLib.ListView ListView1 
         Height          =   2340
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   4128
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6390
         TabIndex        =   7
         Top             =   2745
         Width           =   960
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1110
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   7620
      Begin ITGControls.ITGTextBox txtModule 
         Height          =   285
         Left            =   345
         TabIndex        =   1
         Top             =   255
         Width           =   2730
         _ExtentX        =   4604
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
         Locked          =   -1  'True
         Label           =   "Module"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelWidth      =   1320
         TextBoxWidth    =   1350
      End
      Begin ITGControls.ITGTextBox txtTranNo 
         Height          =   285
         Left            =   345
         TabIndex        =   2
         Top             =   615
         Width           =   2730
         _ExtentX        =   4604
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
         Locked          =   -1  'True
         Label           =   "Trans. No."
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelWidth      =   1320
         TextBoxWidth    =   1350
      End
      Begin ITGControls.ITGTextBox txtType 
         Height          =   285
         Left            =   4365
         TabIndex        =   3
         Top             =   225
         Width           =   2640
         _ExtentX        =   4445
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
         Locked          =   -1  'True
         Label           =   "Type"
         BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelWidth      =   1320
         TextBoxWidth    =   1260
      End
      Begin ITGControls.ITGDateBox dtbDate 
         Height          =   285
         Left            =   5715
         TabIndex        =   4
         Tag             =   "Delivery Date"
         Top             =   600
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         Modal           =   -1  'True
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
         Left            =   4395
         TabIndex        =   5
         Top             =   615
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         Caption         =   "Date"
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
   End
End
Attribute VB_Name = "frmReferenceTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'IT Group Inc. 2005.09.23

Option Explicit

'Reference type
Enum eTransType
    JWRR
    JSO
End Enum

Public mReferType As eTransType
Public mRefPK As String

'Other declarations
Public dtgName As String

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
Dim rsGet As New Recordset
Dim cmdGet As ADODB.Command

    Set itmX = ListView1.ColumnHeaders.Add(, , "Module")
    Set itmX = ListView1.ColumnHeaders.Add(, , "Trans. No.")
    Set itmX = ListView1.ColumnHeaders.Add(, , "Date")
    Set itmX = ListView1.ColumnHeaders.Add(, , "Amount")
    Set itmX = ListView1.ColumnHeaders.Add(, , "Remarks")
    ListView1.ColumnHeaders(1).Width = "1500"
    ListView1.ColumnHeaders(2).Width = "1200"
    ListView1.ColumnHeaders(3).Width = "1000"
    ListView1.ColumnHeaders(4).Width = "1000"
    ListView1.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView1.ColumnHeaders(5).Width = "2600"
    
    txtModule = RepName
    txtTranNo = mRefPK

    Set cmdGet = New ADODB.Command
    Set rsGet = New ADODB.Recordset
    With cmdGet
        .ActiveConnection = cn
        .CommandTimeout = 1000
        .CommandText = "SP_FindReference"
        .CommandType = adCmdStoredProc
        .Parameters("@cCompanyID") = COID
        .Parameters("@cModule") = Trim(txtModule)
        .Parameters("@cTranNo") = Trim(txtTranNo)
    End With

    Set rsGet = cmdGet.Execute()
    
    ListView1.ListItems.Clear
    
    If rsGet.RecordCount <> 0 Then
        While Not rsGet.EOF
            Set itmX = ListView1.ListItems.Add(, , rsGet!cModule & "")
            itmX.SubItems(1) = rsGet!cTranNo & ""
            itmX.SubItems(2) = Format(rsGet!dDate, "mm-dd-yyyy")
            itmX.SubItems(3) = Format(rsGet!nAmount, "#,###.00")
            itmX.SubItems(4) = rsGet!cRemarks & ""
            rsGet.MoveNext
        Wend
    End If

    rsGet.Close
    Set rsGet = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RepName = ""
    cString = ""
    Set frmReferenceTrans = Nothing
End Sub

