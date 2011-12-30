VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B3FB64BF-91F9-11D7-A482-0008A14158BC}#2.22#0"; "ITGControls.ocx"
Begin VB.Form frmITGPickList 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3840
   ClientLeft      =   15
   ClientTop       =   210
   ClientWidth     =   7425
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   7425
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   105
      TabIndex        =   7
      Top             =   2775
      Width           =   4515
   End
   Begin VB.Frame fraSearch 
      Caption         =   "Primary Search Column"
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   3225
      Width           =   2595
      Begin VB.OptionButton optName 
         Caption         =   "Name/Desc"
         Height          =   255
         Left            =   1320
         TabIndex        =   6
         Top             =   180
         Width           =   1155
      End
      Begin VB.OptionButton optCode 
         Caption         =   "ID/Code"
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   180
         Width           =   915
      End
   End
   Begin VB.CommandButton cmbAll 
      Caption         =   "Select All"
      Height          =   345
      Left            =   105
      TabIndex        =   3
      Top             =   3345
      Visible         =   0   'False
      Width           =   1200
   End
   Begin ITGControls.ITGCommandButton cmdCancel 
      Height          =   345
      Left            =   6120
      TabIndex        =   2
      Top             =   3345
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Cancel"
   End
   Begin ITGControls.ITGCommandButton cmdOK 
      Height          =   345
      Left            =   4860
      TabIndex        =   1
      Top             =   3345
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&OK"
   End
   Begin MSComctlLib.ListView lvwList 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
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
Attribute VB_Name = "frmITGPickList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'IT Group Inc. 2005.09.23

Option Explicit

Enum eType
    Accounts
    Accounts_General 'List of general accounts
    APCM    'AP CM list
    APDM    'AP DM list
    APDRCR_DR  'DR list for AP Adjustment
    ARDRCR_WRR  'WRR list for AR Adjustment
    ARCM    'AR CM list
    ARDM    'AR DM list
    Asset
    AssetAccounts
    RFP_Voucher ' List of RFP for voucher
    Bank
    BANKCOMP
    BankAccount
    CheckBook
    CheckBookNo 'For Report
    CHECKS
    Checks_Deposit  'Checks for deposit
    Checks_Deposit_Ref  ' - Add
    
    Client_All  'All clients
    Client  'Both Customer & Supplier
    Client_Customer
    Client_Supplier
    Collection_Invoice  'For Collection Report by Invoice
    Collection_OR   'For Collection Report by OR
    Company
    CMEVATEWT
    Department  'For department field
    DeliveryReceipt
    DepositJournal  'For Deposit Journal Report
    EmployeeList    'Employee List
    Groupings  'For item groupings
    IssuedCheck
    Item
    Item_Inv
    MarketSegment
    Product
    ProductPM
    ProductDivision
    ProductionInv
    ProfitCenter  'List of Profit Centers
    Project  'List of Project Setup
    PurchaseInvoice
    PurchaseOrder_Details
    PurchaseOrder_Header
    PRForCanvass     'Purchase Request for Canvass
    PRItemForCanvass     'Purchase Request Items for Canvass
    Role
    RJOList
    Sales_Order
    SalesInvoice
    Salesman ' Project leader/manager
    SecDepartment  'Security department list
    SecRole  'Security role list
    SI_Journal  'Sales Journal by Invoice report list
    Supply
    TimeDeposit
    Loan
    User
    Voucher
    Warehouse
    RVoucher
    RPurchaseRequest
End Enum

Public OR_Num As String
Public Check_Num As String
Public OR_Str As String
Public Check_Str As String

Public mType As eType
Public mParam As String
Public mRefNo As String
Public mCode As String
Public mName As String
Public mQty As Double
Public mUnit As String
Public mDisc As String
Public mDate As String
Private itmX As Variant
Private rsPickList As New ADODB.Recordset
Public sColumnVariable As String
Private oConnection As New clsConnection
Private connList As ADODB.Connection
Private mCriteria As String

Private Sub cmbAll_Click()
Dim i As Integer
    If cmbAll.Caption = "Select All" Then
        cmbAll.Caption = "Deselect All"
        For i = 1 To lvwList.ListItems.Count
            lvwList.ListItems(i).Checked = True
        Next i
    Else
        cmbAll.Caption = "Select All"
        For i = 1 To lvwList.ListItems.Count
            lvwList.ListItems(i).Checked = False
        Next i
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
    Set frmITGPickList = Nothing
End Sub

Private Sub cmdOK_Click()
    SelectOK
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
        Set frmITGPickList = Nothing
    ElseIf KeyAscii = 8 Then
        If txtFind.Text <> "" Then
            txtFind = Mid(txtFind.Text, 1, Len(txtFind.Text) - 1)
        End If
    Else
        txtFind = txtFind.Text + UCase(Chr(KeyAscii))
    End If
    'sFilterString = txtFind.Text
    'ShowForm
End Sub

Private Sub Form_Load()
    lPickListActive = True
    Select Case mType
        Case Product, ProductPM
            optCode.Value = True
        Case PRItemForCanvass
            optCode.Value = True
        Case Client_All
            optName.Value = True
        Case Client
            optName.Value = True
        Case Client_Customer
            optName.Value = True
        Case Client_Supplier
            optName.Value = True
        Case EmployeeList
            optName.Value = True
        Case Salesman
            optName.Value = True
        Case MarketSegment
            optName.Value = True
        Case ProfitCenter
            optName.Value = True
        Case ProductDivision
            optName.Value = True
        Case RVoucher, RPurchaseRequest
            optName.Value = True
        Case Else
            fraSearch.Visible = False
            ShowForm
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RepName = Empty
    sFilterString = Empty
    lPickListActive = False
End Sub

Private Sub lvwList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwList.Sorted = True
    lvwList.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub lvwList_DblClick()
    If lvwList.ListItems.Count = 0 Then Exit Sub
    If lvwList.SelectedItem.Selected Then
        Call SelectOK
    End If
End Sub

Private Sub lvwList_KeyPress(KeyAscii As Integer)
    If lvwList.ListItems.Count = 0 Then Exit Sub
    If KeyAscii = 13 Then
        Call SelectOK
    End If
End Sub

Private Sub SelectOK()
    Dim j As Integer
    mCode = ""
    mName = ""
    If lvwList.ListItems.Count = 0 Then Exit Sub
    If lvwList.SelectedItem.Selected = False Then Exit Sub
    
    Select Case mType
        Case Sales_Order
            mRefNo = Trim(lvwList.SelectedItem.Text)
            mCode = Trim(lvwList.SelectedItem.SubItems(1))
            mName = Trim(lvwList.SelectedItem.SubItems(2))
        Case PurchaseOrder_Details
            mRefNo = Trim(lvwList.SelectedItem.Text)
            mCode = Trim(lvwList.SelectedItem.SubItems(1))
            mName = Trim(lvwList.SelectedItem.SubItems(2))
            mUnit = Trim(lvwList.SelectedItem.SubItems(3))
            mQty = Trim(lvwList.SelectedItem.SubItems(4))
        Case PRItemForCanvass
            If sColumnVariable = "Code" Then
                mName = Trim(lvwList.SelectedItem.SubItems(1))
                mCode = Trim(lvwList.SelectedItem.Text)
            ElseIf sColumnVariable = "Desc" Then
                mCode = Trim(lvwList.SelectedItem.SubItems(1))
                mName = Trim(lvwList.SelectedItem.Text)
            Else
                mCode = Trim(lvwList.SelectedItem.SubItems(1))
                mName = Trim(lvwList.SelectedItem.Text)
            End If
            mQty = Trim(lvwList.SelectedItem.SubItems(2))
            mRefNo = Trim(lvwList.SelectedItem.SubItems(3))
        Case CHECKS
            mCode = Trim(lvwList.SelectedItem.Text) 'Check number
            mDate = Trim(lvwList.SelectedItem.SubItems(1)) 'Check date
            mRefNo = Trim(lvwList.SelectedItem.SubItems(2)) 'Bank ID
            mQty = Trim(lvwList.SelectedItem.SubItems(3)) 'Check amount
        Case Accounts
            If Trim(lvwList.SelectedItem.SubItems(2)) <> "General" Then
                mCode = Trim(lvwList.SelectedItem.Text)
                mName = Trim(lvwList.SelectedItem.SubItems(1))
                mRefNo = Trim(lvwList.SelectedItem.SubItems(2))
                mQty = Trim(lvwList.SelectedItem.SubItems(3))
            End If
        Case Accounts_General
            mCode = Trim(lvwList.SelectedItem.Text)
            mName = Trim(lvwList.SelectedItem.SubItems(1))
            mRefNo = Trim(lvwList.SelectedItem.SubItems(2))
            mQty = Trim(lvwList.SelectedItem.SubItems(3))
        Case CMEVATEWT
            mRefNo = Trim(lvwList.SelectedItem.Text) 'Invoice No.
            mCode = Trim(lvwList.SelectedItem.SubItems(1)) 'EWT/EVAT No.
            mName = Trim(lvwList.SelectedItem.SubItems(2)) 'Type
            mQty = Trim(lvwList.SelectedItem.SubItems(3)) 'Amount
            mUnit = Trim(lvwList.SelectedItem.SubItems(4)) 'AR Tran No.
        Case BANKCOMP
            mCode = Trim(lvwList.SelectedItem.Text)
            mName = Trim(lvwList.SelectedItem.SubItems(1))
            mParam = Trim(lvwList.SelectedItem.SubItems(2)) ' Checkbook Number
            mRefNo = Trim(lvwList.SelectedItem.SubItems(3)) ' Check Number
            mUnit = Trim(lvwList.SelectedItem.SubItems(4))  ' Bank Account No.
        Case RVoucher

        For j = 1 To lvwList.ListItems.Count
                If lvwList.ListItems(j).Checked = True Then
                    If sColumnVariable = "Desc" Then
                        mCode = mCode & "'" & Trim(lvwList.ListItems(j).SubItems(1)) & "',"
                    Else
                        mCode = mCode & "'" & Trim(lvwList.ListItems(j)) & "',"
                    End If
                End If
            Next j
            If Trim(mCode) <> "" Then
                mCode = Left(mCode, Len(mCode) - 1)
            End If
            
        Case RPurchaseRequest

        For j = 1 To lvwList.ListItems.Count
                If lvwList.ListItems(j).Checked = True Then
                    If sColumnVariable = "Desc" Then
                        mCode = mCode & "''" & Trim(lvwList.ListItems(j).SubItems(1)) & "'',"
                    Else
                        mCode = mCode & "''" & Trim(lvwList.ListItems(j)) & "'',"
                    End If
                End If
            Next j
            If Trim(mCode) <> "" Then
                mCode = Left(mCode, Len(mCode) - 1)
            End If
            

        Case DeliveryReceipt
            mCode = Trim(lvwList.SelectedItem.Text) 'DR No
            mName = Trim(lvwList.SelectedItem.SubItems(1))  'SO No
            mRefNo = Trim(lvwList.SelectedItem.SubItems(2)) 'Customer
            mUnit = Trim(lvwList.SelectedItem.SubItems(3))  'Address
            mParam = Trim(lvwList.SelectedItem.SubItems(4)) 'Terms
        Case APDRCR_DR
            mCode = Trim(lvwList.SelectedItem.Text) 'DR No
        Case ARDRCR_WRR, RFP_Voucher
            mCode = Trim(lvwList.SelectedItem.Text) 'WRR No

        Case CheckBookNo
            mRefNo = Trim(lvwList.SelectedItem.Text)
            mCode = Trim(lvwList.SelectedItem.SubItems(1))
            mName = Trim(lvwList.SelectedItem.SubItems(2))
        Case APCM
            mRefNo = Trim(lvwList.SelectedItem.Text)
            mCode = Trim(lvwList.SelectedItem.SubItems(1))
        Case APDM
            mRefNo = Trim(lvwList.SelectedItem.Text)
            mCode = Trim(lvwList.SelectedItem.SubItems(1))
        Case ARCM
            mRefNo = Trim(lvwList.SelectedItem.Text)
            mCode = Trim(lvwList.SelectedItem.SubItems(1))
        Case ARDM
            mRefNo = Trim(lvwList.SelectedItem.Text)
            mCode = Trim(lvwList.SelectedItem.SubItems(1))
        Case RJOList
            mRefNo = Trim(lvwList.SelectedItem.Text)
        Case Collection_Invoice
            mRefNo = Trim(lvwList.SelectedItem.Text)
        Case Collection_OR
            mRefNo = Trim(lvwList.SelectedItem.Text)
        Case SI_Journal
            mRefNo = Trim(lvwList.SelectedItem.Text)
        Case ProductionInv
            mRefNo = Trim(lvwList.SelectedItem.Text)
        Case ProductPM
'            frmMaintPriceMatrix.AddDetailsFromList
        Case DepositJournal
            mRefNo = Trim(lvwList.SelectedItem.Text)
        Case TimeDeposit
            mCode = Trim(lvwList.SelectedItem.Text)
            mName = lvwList.SelectedItem.SubItems(1)
            mQty = lvwList.SelectedItem.SubItems(2)
            mDate = lvwList.SelectedItem.SubItems(3)
            mUnit = lvwList.SelectedItem.SubItems(4)
        Case Loan
            mCode = Trim(lvwList.SelectedItem.Text)
            mName = lvwList.SelectedItem.SubItems(1)
            mQty = lvwList.SelectedItem.SubItems(2)
            mDate = lvwList.SelectedItem.SubItems(3)
            mUnit = lvwList.SelectedItem.SubItems(4)
        Case PRForCanvass
            mName = Trim(lvwList.SelectedItem.Text)
            mCode = Trim(lvwList.SelectedItem.Text)
            
     
           
        Case Else
            If sColumnVariable = "Code" Then
                mName = Trim(lvwList.SelectedItem.SubItems(1))
                mCode = Trim(lvwList.SelectedItem.Text)
            ElseIf sColumnVariable = "Desc" Then
                mCode = Trim(lvwList.SelectedItem.SubItems(1))
                mName = Trim(lvwList.SelectedItem.Text)
            Else
                mCode = Trim(lvwList.SelectedItem.SubItems(1))
                mName = Trim(lvwList.SelectedItem.Text)
            End If
            If RepName = "CommissionCriteria" Then
'                frmMaintCommissionCriteria.AddDetailsFromList
            End If
    
        
    End Select
    
    If mType = Accounts Then
        If Trim(lvwList.SelectedItem.SubItems(2)) <> "General" Then
            Unload Me
        End If
    Else
        Unload Me
    End If
    
    
    
End Sub

Public Sub ShowForm()
On Error GoTo TheSource
    lvwList.ColumnHeaders.Clear
    lvwList.ListItems.Clear
    If rsPickList.State = adStateOpen Then rsPickList.Close
    DoEvents
    oConnection.OpenNewConnection connList
    MousePointer = vbHourglass
    lvwList.Visible = False
    If Not lModal Then FormWaitShow "Loading list . . ."
    Select Case mType
        Case Salesman
            Caption = "Project Manager List"
            sSQL = "SELECT cCode, cName FROM Salesman WHERE cCompanyID = '" & Trim(COID) & "' ORDER BY cCode"
            If RepName = "CommissionCriteria" Then
                sSQL = "SELECT cCode,cName FROM SALESMAN WHERE cCompanyID = '" & Trim(COID) & "' " & _
                        "AND cCode NOT IN (SELECT A.cID FROM COMMISSION_APPLYTO A LEFT OUTER JOIN COMMISSION B ON " & _
                        "A.cCompanyID = B.cCompanyID AND A.cComID = B.cComID WHERE B.cType = 'Quota')" & _
                        "ORDER BY cCode "
            End If
            
            rsPickList.Open sSQL, connList, adOpenForwardOnly, adLockReadOnly, adCmdText
            If sFilterString <> "" Then rsPickList.Filter = "cCode LIKE '" & sFilterString & "%'"
            
            If sColumnVariable = "Code" Then
                Set itmX = lvwList.ColumnHeaders.Add(, , "PJM ID")
                Set itmX = lvwList.ColumnHeaders.Add(, , "PJM Name")
                lvwList.ColumnHeaders(1).Width = "2500"
                lvwList.ColumnHeaders(2).Width = "5440"
                    
                Do Until rsPickList.EOF
                    Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cCode) & "")
                    itmX.SubItems(1) = Trim(rsPickList!cName) & ""
                    rsPickList.MoveNext
                Loop
            ElseIf sColumnVariable = "Desc" Then
                Set itmX = lvwList.ColumnHeaders.Add(, , "PJM Name")
                Set itmX = lvwList.ColumnHeaders.Add(, , "PJM ID")
                lvwList.ColumnHeaders(1).Width = "5440"
                lvwList.ColumnHeaders(2).Width = "2500"
                    
                Do Until rsPickList.EOF
                    Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cName) & "")
                    itmX.SubItems(1) = Trim(rsPickList!cCode) & ""
                    rsPickList.MoveNext
                Loop
            End If
        Case MarketSegment
            Caption = "Market Segment List"
            sSQL = "SELECT cClassCode, cDescription FROM Classification WHERE cType = 'MS' AND cCompanyID = '" & Trim(COID) & "' ORDER BY cDescription"
            rsPickList.Open sSQL, connList, adOpenForwardOnly, adLockReadOnly, adCmdText
            If sFilterString <> "" Then rsPickList.Filter = "cClassCode LIKE '" & sFilterString & "%'"
            If sColumnVariable = "Code" Then
                Set itmX = lvwList.ColumnHeaders.Add(, , "Market Segment ID")
                Set itmX = lvwList.ColumnHeaders.Add(, , "Market Segment")
                lvwList.ColumnHeaders(1).Width = "2500"
                lvwList.ColumnHeaders(2).Width = "5440"
                    
                Do Until rsPickList.EOF
                    Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cClassCode) & "")
                    itmX.SubItems(1) = Trim(rsPickList!cDescription) & ""
                    rsPickList.MoveNext
                Loop
            ElseIf sColumnVariable = "Desc" Then
                Set itmX = lvwList.ColumnHeaders.Add(, , "Market Segment")
                Set itmX = lvwList.ColumnHeaders.Add(, , "Market Segment ID")
                lvwList.ColumnHeaders(1).Width = "5440"
                lvwList.ColumnHeaders(2).Width = "2500"
                    
                Do Until rsPickList.EOF
                    Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cDescription) & "")
                    itmX.SubItems(1) = Trim(rsPickList!cClassCode) & ""
                    rsPickList.MoveNext
                Loop
            End If
        Case ProfitCenter
            Caption = "Profit Center List"
            sSQL = "SELECT cPCCode, cDescription FROM PROFITCENTER WHERE cCompanyID = '" & Trim(COID) & "' ORDER BY cDescription"
            rsPickList.Open sSQL, connList, adOpenForwardOnly, adLockReadOnly, adCmdText
            If sColumnVariable = "Code" Then
                Set itmX = lvwList.ColumnHeaders.Add(, , "Profit Center ID")
                Set itmX = lvwList.ColumnHeaders.Add(, , "Profit Center Description")
                lvwList.ColumnHeaders(1).Width = "2500"
                lvwList.ColumnHeaders(2).Width = "5440"
                Do Until rsPickList.EOF
                    Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cPCCode) & "")
                    itmX.SubItems(1) = Trim(rsPickList!cDescription) & ""
                    rsPickList.MoveNext
                Loop
            ElseIf sColumnVariable = "Desc" Then
                Set itmX = lvwList.ColumnHeaders.Add(, , "Profit Center Description")
                Set itmX = lvwList.ColumnHeaders.Add(, , "Profit Center ID")
                lvwList.ColumnHeaders(1).Width = "5440"
                lvwList.ColumnHeaders(2).Width = "2500"
                    
                Do Until rsPickList.EOF
                    Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cDescription) & "")
                    itmX.SubItems(1) = Trim(rsPickList!cPCCode) & ""
                    rsPickList.MoveNext
                Loop
            End If
        Case EmployeeList
            Caption = "Employee List"
            sSQL = "SELECT cEmpCode, cEmpName FROM EMPLOYEE WHERE cCompanyID = '" & Trim(COID) & "' ORDER BY cEmpName"
            rsPickList.Open sSQL, connList, adOpenForwardOnly, adLockReadOnly, adCmdText
            If sColumnVariable = "Code" Then
                Set itmX = lvwList.ColumnHeaders.Add(, , "Employee Code")
                Set itmX = lvwList.ColumnHeaders.Add(, , "Employee Name")
                lvwList.ColumnHeaders(1).Width = "2500"
                lvwList.ColumnHeaders(2).Width = "5440"
                Do Until rsPickList.EOF
                    Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cEmpCode) & "")
                    itmX.SubItems(1) = Trim(rsPickList!cEmpName) & ""
                    rsPickList.MoveNext
                Loop
            ElseIf sColumnVariable = "Desc" Then
                Set itmX = lvwList.ColumnHeaders.Add(, , "Employee Name")
                Set itmX = lvwList.ColumnHeaders.Add(, , "Employee Code")
                lvwList.ColumnHeaders(1).Width = "5440"
                lvwList.ColumnHeaders(2).Width = "2500"
                    
                Do Until rsPickList.EOF
                    Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cEmpName) & "")
                    itmX.SubItems(1) = Trim(rsPickList!cEmpCode) & ""
                    rsPickList.MoveNext
                Loop
            End If
        Case ProductDivision
            Caption = "Product Division List"
            sSQL = "SELECT cClassCode, cDescription FROM Classification WHERE cType = 'PD' AND cCompanyID = '" & Trim(COID) & "' ORDER BY cDescription"
            rsPickList.Open sSQL, connList, adOpenForwardOnly, adLockReadOnly, adCmdText
            If sColumnVariable = "Code" Then
                Set itmX = lvwList.ColumnHeaders.Add(, , "Product Division ID")
                Set itmX = lvwList.ColumnHeaders.Add(, , "Product Division")
                lvwList.ColumnHeaders(1).Width = "2500"
                lvwList.ColumnHeaders(2).Width = "5440"
                    
                Do Until rsPickList.EOF
                    Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cClassCode) & "")
                    itmX.SubItems(1) = Trim(rsPickList!cDescription) & ""
                    rsPickList.MoveNext
                Loop
            ElseIf sColumnVariable = "Desc" Then
                Set itmX = lvwList.ColumnHeaders.Add(, , "Product Division")
                Set itmX = lvwList.ColumnHeaders.Add(, , "Product Division ID")
                lvwList.ColumnHeaders(1).Width = "5440"
                lvwList.ColumnHeaders(2).Width = "2500"
                    
                Do Until rsPickList.EOF
                    Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cDescription) & "")
                    itmX.SubItems(1) = Trim(rsPickList!cClassCode) & ""
                    rsPickList.MoveNext
                Loop
            End If
        Case PRItemForCanvass
            List_PRItemForCanvass
        Case Product, ProductPM
            Caption = "Product List"
            sSQL = "SELECT cItemNo, cDesc FROM ITEM WHERE cCompanyID = '" & Trim(COID) & "' ORDER BY cDesc"
            
            If RepName = "CommissionCriteria" Then
                sSQL = "SELECT cItemNo, cDesc FROM ITEM WHERE cCompanyID = '" & Trim(COID) & "' " & _
                        "AND cItemNo NOT IN (SELECT A.cID FROM COMMISSION_APPLYTO A LEFT OUTER JOIN COMMISSION B ON " & _
                        "A.cCompanyID = B.cCompanyID AND A.cComID = B.cComID WHERE B.cType = 'Profit Margin')" & _
                        "ORDER BY cDesc "
            End If
            rsPickList.Open sSQL, connList, adOpenForwardOnly, adLockReadOnly, adCmdText
            If sFilterString <> "" Then rsPickList.Filter = "cItemNo LIKE '" & sFilterString & "%'"
            If sColumnVariable = "Code" Then
                Set itmX = lvwList.ColumnHeaders.Add(, , "Product ID")
                Set itmX = lvwList.ColumnHeaders.Add(, , "Product Description")
                lvwList.ColumnHeaders(1).Width = "2500"
                lvwList.ColumnHeaders(2).Width = "5440"
                Do Until rsPickList.EOF
                    Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cItemNo) & "")
                    itmX.SubItems(1) = Trim(rsPickList!cDesc) & ""
                    rsPickList.MoveNext
                Loop
            ElseIf sColumnVariable = "Desc" Then
                Set itmX = lvwList.ColumnHeaders.Add(, , "Product Description")
                Set itmX = lvwList.ColumnHeaders.Add(, , "Product ID")
                lvwList.ColumnHeaders(1).Width = "5440"
                lvwList.ColumnHeaders(2).Width = "2500"
                    
                Do Until rsPickList.EOF
                    Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cDesc) & "")
                    itmX.SubItems(1) = Trim(rsPickList!cItemNo) & ""
                    rsPickList.MoveNext
                Loop
            End If
        Case Groupings
            Caption = "Product Grouping List"
            Set itmX = lvwList.ColumnHeaders.Add(, , "Description")
            Set itmX = lvwList.ColumnHeaders.Add(, , "ID")
            lvwList.ColumnHeaders(1).Width = "5440"
            lvwList.ColumnHeaders(2).Width = "2500"
            sSQL = "SELECT cID, cDescription FROM GROUPINGS WHERE cCompanyID = '" & Trim(COID) & "' " & _
                    "AND cGroupNo = '" & RepName & "' ORDER BY cDescription"
            rsPickList.Open sSQL, connList, adOpenForwardOnly, adLockReadOnly, adCmdText
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cDescription) & "")
                itmX.SubItems(1) = Trim(rsPickList!cID) & ""

                rsPickList.MoveNext
            Loop
        Case Warehouse
            Caption = "Warehouse List"
            Set itmX = lvwList.ColumnHeaders.Add(, , "Warehouse Name")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Warehouse ID")
            lvwList.ColumnHeaders(1).Width = "5440"
            lvwList.ColumnHeaders(2).Width = "2500"
                
            sSQL = "SELECT cWH, cName FROM WHSE WHERE cCompanyID = '" & Trim(COID) & "' ORDER BY cName"
            rsPickList.Open sSQL, connList, adOpenForwardOnly, adLockReadOnly, adCmdText
                
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cName) & "")
                itmX.SubItems(1) = Trim(rsPickList!cWH) & ""

                rsPickList.MoveNext
            Loop
            
        Case Sales_Order
            Caption = "Sales Order List"
            Set itmX = lvwList.ColumnHeaders.Add(, , "Sales Order ID")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Supplier ID")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Supllier Name")
            lvwList.ColumnHeaders(1).Width = "1500"
            lvwList.ColumnHeaders(2).Width = "1500"
            lvwList.ColumnHeaders(3).Width = "6000"
            
            sSQL = "SELECT A.cSONo, A.cCode, B.cName " & _
                    "FROM SO A Left Outer Join Supplier B On A.cCode = B.cCode " & _
                    "WHERE A.cCompanyID = '" & Trim(COID) & "' ORDER BY cName "

            rsPickList.Open sSQL, connList, adOpenForwardOnly, adLockReadOnly, adCmdText
                
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!csono) & "")
                itmX.SubItems(1) = IIf(IsNull(rsPickList!cCode), "", Trim(rsPickList!cCode))
                itmX.SubItems(2) = IIf(IsNull(rsPickList!cName), "", Trim(rsPickList!cName))

                rsPickList.MoveNext
            Loop
        
        Case PurchaseOrder_Header
            Caption = "Suppliers List"
            Set itmX = lvwList.ColumnHeaders.Add(, , "Supplier Name")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Supplier ID")
            lvwList.ColumnHeaders(1).Width = "5440"
            lvwList.ColumnHeaders(2).Width = "2500"

            sSQL = "SELECT cCode,cName, cAddress, cContact FROM Supplier WHERE cCompanyID = '" & Trim(COID) & "' ORDER BY cCode "
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cName) & "")
                itmX.SubItems(1) = Trim(rsPickList!cCode) & ""
                
                rsPickList.MoveNext
            Loop
        
        Case PurchaseOrder_Details
            Caption = "Purchase Request List"
            Set itmX = lvwList.ColumnHeaders.Add(, , "Request ID")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Product ID")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Description")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Unit")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Qty")
            lvwList.ColumnHeaders(1).Width = "1000"
            lvwList.ColumnHeaders(2).Width = "1000"
            lvwList.ColumnHeaders(3).Width = "3500"
            lvwList.ColumnHeaders(4).Width = "1000"
            lvwList.ColumnHeaders(5).Width = "1000"
            
            sSQL = "Select cPRNo, cItemno, cDesc, cUnit, nQty " & _
                    "From Requisition_T " & _
                    "WHERE cCompanyID = '" & Trim(COID) & "'  and lServe = 0 ORDER BY cItemno "

            rsPickList.Open sSQL, connList, adOpenForwardOnly, adLockReadOnly, adCmdText
                
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cPRNo))
                itmX.SubItems(1) = IIf(IsNull(rsPickList!cItemNo), "", Trim(rsPickList!cItemNo))
                itmX.SubItems(2) = IIf(IsNull(rsPickList!cDesc), "", Trim(rsPickList!cDesc))
                itmX.SubItems(3) = IIf(IsNull(rsPickList!cUnit), "", Trim(rsPickList!cUnit))
                itmX.SubItems(4) = IIf(IsNull(rsPickList!nQty), "0", Trim(rsPickList!nQty))
                rsPickList.MoveNext
            Loop
        
        Case Accounts
            Caption = "Account List"
            Set itmX = lvwList.ColumnHeaders.Add(, , "Account Number")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Account Title")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Account Type")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Balance")
            lvwList.ColumnHeaders(1).Width = "1500"
            lvwList.ColumnHeaders(2).Width = "2750"
            lvwList.ColumnHeaders(3).Width = "1250"
            lvwList.ColumnHeaders(4).Width = "1700"
            lvwList.ColumnHeaders(4).Alignment = lvwColumnRight

            sSQL = "SELECT cAcctNo,cTitle,cType,nBalance FROM ACCOUNT WHERE cCompanyID = '" & Trim(COID) & "' ORDER BY cAcctNo "
            rsPickList.Open sSQL, connList, adOpenKeyset
            If sFilterString <> "" Then rsPickList.Filter = "cAcctNo LIKE '" & sFilterString & "%'"
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cAcctNo) & "")
                itmX.SubItems(1) = Trim(rsPickList!cTitle) & ""
                itmX.SubItems(2) = Trim(rsPickList!cType) & ""
                itmX.SubItems(3) = Format(rsPickList!nBalance, "#,##0.#0") & ""
                
                rsPickList.MoveNext
            Loop
            
        Case Accounts_General
            Caption = "Account List (General Accounts)"
            Set itmX = lvwList.ColumnHeaders.Add(, , "Account Number")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Account Title")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Account Type")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Level")
            lvwList.ColumnHeaders(1).Width = "1500"
            lvwList.ColumnHeaders(2).Width = "2750"
            lvwList.ColumnHeaders(3).Width = "1500"
            lvwList.ColumnHeaders(4).Width = "1450"
            lvwList.ColumnHeaders(4).Alignment = lvwColumnRight

            sSQL = "SELECT cAcctNo,cTitle,cType,cLevel FROM ACCOUNT WHERE cCompanyID = '" & Trim(COID) & "' " & _
                    "AND cCategory = '" & Trim(RepName) & "' AND cType = 'GENERAL' " & _
                    "ORDER BY cAcctNo "
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cAcctNo) & "")
                itmX.SubItems(1) = Trim(rsPickList!cTitle) & ""
                itmX.SubItems(2) = Trim(rsPickList!cType) & ""
                itmX.SubItems(3) = Format(rsPickList!cLevel, "#,##0") & ""
                
                rsPickList.MoveNext
            Loop
        
        Case AssetAccounts
            Caption = "Account List (Assets)"
            Set itmX = lvwList.ColumnHeaders.Add(, , "Account Title")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Account Number")
            lvwList.ColumnHeaders(1).Width = "5440"
            lvwList.ColumnHeaders(2).Width = "2500"

            sSQL = "SELECT cAcctNo,cTitle FROM ACCOUNT WHERE cCompanyID = '" & Trim(COID) & "' AND cAcctNo like '12%' ORDER BY cAcctNo "
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cTitle) & "")
                itmX.SubItems(1) = Trim(rsPickList!cAcctNo) & ""
                
                rsPickList.MoveNext
            Loop
        
        Case Asset
            Caption = "Asset List" & cStrTypeHdr
            Set itmX = lvwList.ColumnHeaders.Add(, , "Asset Description")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Asset ID")
            lvwList.ColumnHeaders(1).Width = "5440"
            lvwList.ColumnHeaders(2).Width = "2500"

            sSQL = "SELECT cAssetNo,cDesc FROM ASSET WHERE cCompanyID = '" & Trim(COID) & "' " & cStrType & " ORDER BY cDesc "
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cDesc) & "")
                itmX.SubItems(1) = Trim(rsPickList!cAssetNo) & ""
                
                rsPickList.MoveNext
            Loop
        Case Supply
            Caption = "Supplies List"
            Set itmX = lvwList.ColumnHeaders.Add(, , "Supply Description")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Supply ID")
            lvwList.ColumnHeaders(1).Width = "5440"
            lvwList.ColumnHeaders(2).Width = "2500"

            sSQL = "SELECT cSupplyNo, cDesc FROM SUPPLY WHERE cCompanyID = '" & Trim(COID) & "' ORDER BY cDesc "
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cDesc) & "")
                itmX.SubItems(1) = Trim(rsPickList!cSupplyNo) & ""
                
                rsPickList.MoveNext
            Loop
    
        Case Client_All
            Dim cClientType As String
            Caption = "Client List (All)"
            
            sSQL = "SELECT cCode,cName,lCompany,lIndividual FROM CLIENT_CUSTOMER WHERE cCompanyID = '" & Trim(COID) & "' ORDER BY cCode "
            
            rsPickList.Open sSQL, connList, adOpenKeyset
            If sFilterString <> "" Then rsPickList.Filter = "cCode LIKE '" & sFilterString & "%'"
            
            If sColumnVariable = "Code" Then
                Set itmX = lvwList.ColumnHeaders.Add(, , "Client ID")
                Set itmX = lvwList.ColumnHeaders.Add(, , "Client Name")
                Set itmX = lvwList.ColumnHeaders.Add(, , "Client Type")
                lvwList.ColumnHeaders(1).Width = "2500"
                lvwList.ColumnHeaders(2).Width = "5440"
                lvwList.ColumnHeaders(3).Width = "2000"
                
                Do Until rsPickList.EOF
                    Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cCode) & "")
                    itmX.SubItems(1) = Trim(rsPickList!cName) & ""
                    'If rsPickList!lCustomer = True Then
                    '    If rsPickList!lSupplier = True Then
                    '        cClientType = "Customer/Supplier"
                    '    Else
                            cClientType = "Customer"
                    '    End If
                    'Else
                    '    cClientType = "Supplier"
                    'End If
                    itmX.SubItems(2) = cClientType
                    rsPickList.MoveNext
                Loop
            ElseIf sColumnVariable = "Desc" Then
                Set itmX = lvwList.ColumnHeaders.Add(, , "Client Name")
                Set itmX = lvwList.ColumnHeaders.Add(, , "Client ID")
                Set itmX = lvwList.ColumnHeaders.Add(, , "Client Type")
                lvwList.ColumnHeaders(1).Width = "5440"
                lvwList.ColumnHeaders(2).Width = "2500"
                lvwList.ColumnHeaders(3).Width = "2000"
                
                Do Until rsPickList.EOF
                    Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cName) & "")
                    itmX.SubItems(1) = Trim(rsPickList!cCode) & ""
                    'If rsPickList!lCustomer = True Then
                    '    If rsPickList!lSupplier = True Then
                    '        cClientType = "Customer/Supplier"
                    '    Else
                            cClientType = "Customer"
                    '    End If
                    'Else
                    '    cClientType = "Supplier"
                    'End If
                    itmX.SubItems(2) = cClientType
                    rsPickList.MoveNext
                Loop
            End If
                
        Case Client
            Caption = "Client List (Customers/Supplier)"
            
            sSQL = "SELECT cCode,cName FROM CLIENT_CUSTOMER WHERE cCompanyID = '" & Trim(COID) & "' ORDER BY cCode "
            
            rsPickList.Open sSQL, connList, adOpenKeyset
            If sFilterString <> "" Then rsPickList.Filter = "cCode LIKE '" & sFilterString & "%'"
            
            If sColumnVariable = "Code" Then
                Set itmX = lvwList.ColumnHeaders.Add(, , "Client ID")
                Set itmX = lvwList.ColumnHeaders.Add(, , "Client Name")
                lvwList.ColumnHeaders(1).Width = "2500"
                lvwList.ColumnHeaders(2).Width = "5440"
                
                Do Until rsPickList.EOF
                    Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cCode) & "")
                    itmX.SubItems(1) = Trim(rsPickList!cName) & ""
                    rsPickList.MoveNext
                Loop
            ElseIf sColumnVariable = "Desc" Then
                Set itmX = lvwList.ColumnHeaders.Add(, , "Client Name")
                Set itmX = lvwList.ColumnHeaders.Add(, , "Client ID")
                lvwList.ColumnHeaders(1).Width = "5440"
                lvwList.ColumnHeaders(2).Width = "2500"
                
                Do Until rsPickList.EOF
                    Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cName) & "")
                    itmX.SubItems(1) = Trim(rsPickList!cCode) & ""
                    rsPickList.MoveNext
                Loop
            End If
        Case TimeDeposit
            Caption = "Time Deposit List"
            sSQL = "SELECT cLoanTDNo, nPrincipal, nRate, dDateFrom, dDateTo FROM LoanTD WHERE cType = 'TD' AND cCompanyID = '" & Trim(COID) & "'"
            rsPickList.Open sSQL, connList, adOpenForwardOnly, adLockReadOnly, adCmdText
            
            'If sColumnVariable = "Code" Then
                Set itmX = lvwList.ColumnHeaders.Add(, , "TD Number")
                Set itmX = lvwList.ColumnHeaders.Add(, , "Principal")
                Set itmX = lvwList.ColumnHeaders.Add(, , "Rate")
                Set itmX = lvwList.ColumnHeaders.Add(, , "Date From")
                Set itmX = lvwList.ColumnHeaders.Add(, , "Date To")
                lvwList.ColumnHeaders(1).Width = "1500"
                lvwList.ColumnHeaders(2).Width = "2000"
                lvwList.ColumnHeaders(3).Width = "2000"
                lvwList.ColumnHeaders(4).Width = "2000"
                lvwList.ColumnHeaders(5).Width = "2000"
                
                Do Until rsPickList.EOF
                    Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cLoanTDNo) & "")
                    itmX.SubItems(1) = Trim(rsPickList!nPrincipal) & ""
                    itmX.SubItems(2) = Trim(rsPickList!nRate) & ""
                    itmX.SubItems(3) = Trim(rsPickList!dDateFrom) & ""
                    itmX.SubItems(4) = Trim(rsPickList!dDateTo) & ""
                    
                    rsPickList.MoveNext
                Loop
            'End If
        Case Loan
            Caption = "Loan Release List"
            sSQL = "SELECT cLoanTDNo, nPrincipal, nRate, dDateFrom, dDateTo FROM LoanTD WHERE cType = 'LOAN' AND cCompanyID = '" & Trim(COID) & "'"
            rsPickList.Open sSQL, connList, adOpenForwardOnly, adLockReadOnly, adCmdText
            Set itmX = lvwList.ColumnHeaders.Add(, , "Loan ID")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Principal")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Rate")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Date From")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Date To")
            lvwList.ColumnHeaders(1).Width = "1500"
            lvwList.ColumnHeaders(2).Width = "2000"
            lvwList.ColumnHeaders(3).Width = "2000"
            lvwList.ColumnHeaders(4).Width = "2000"
            lvwList.ColumnHeaders(5).Width = "2000"
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cLoanTDNo) & "")
                itmX.SubItems(1) = Trim(rsPickList!nPrincipal) & ""
                itmX.SubItems(2) = Trim(rsPickList!nRate) & ""
                itmX.SubItems(3) = Trim(rsPickList!dDateFrom) & ""
                itmX.SubItems(4) = Trim(rsPickList!dDateTo) & ""
                
                rsPickList.MoveNext
            Loop
        Case Client_Customer
            Caption = "Client List (Customers)"
            
            sSQL = "SELECT cCode,cName FROM CLIENT_CUSTOMER WHERE cCompanyID = '" & Trim(COID) & "' ORDER BY cCode "
            
            If RepName = "CommissionCriteria" Then
                sSQL = "SELECT cCode,cName FROM CLIENT_CUSTOMER WHERE cCompanyID = '" & Trim(COID) & "'" & _
                        "AND cCode NOT IN (SELECT A.cID FROM COMMISSION_APPLYTO A LEFT OUTER JOIN COMMISSION B ON " & _
                        "A.cCompanyID = B.cCompanyID AND A.cComID = B.cComID WHERE B.cType = 'Terms')" & _
                        "ORDER BY cCode "
            End If
            
            rsPickList.Open sSQL, connList, adOpenKeyset
            If sFilterString <> "" Then rsPickList.Filter = "cCode LIKE '" & sFilterString & "%'"
            
            If sColumnVariable = "Code" Then
                Set itmX = lvwList.ColumnHeaders.Add(, , "Client ID")
                Set itmX = lvwList.ColumnHeaders.Add(, , "Client Name")
                lvwList.ColumnHeaders(1).Width = "2500"
                lvwList.ColumnHeaders(2).Width = "5440"
                
                Do Until rsPickList.EOF
                    Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cCode) & "")
                    itmX.SubItems(1) = Trim(rsPickList!cName) & ""
                    rsPickList.MoveNext
                Loop
            ElseIf sColumnVariable = "Desc" Then
                Set itmX = lvwList.ColumnHeaders.Add(, , "Client Name")
                Set itmX = lvwList.ColumnHeaders.Add(, , "Client ID")
                lvwList.ColumnHeaders(1).Width = "5440"
                lvwList.ColumnHeaders(2).Width = "2500"
                
                Do Until rsPickList.EOF
                    Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cName) & "")
                    itmX.SubItems(1) = Trim(rsPickList!cCode) & ""
                    rsPickList.MoveNext
                Loop
            End If
                
        Case Client_Supplier
            Caption = "Client List (Suppliers)"
            sSQL = "SELECT cCode,cName FROM CLIENT_SUPPLIER WHERE cCompanyID = '" & Trim(COID) & "' ORDER BY cCode "
            
            rsPickList.Open sSQL, connList, adOpenKeyset
            If sFilterString <> "" Then rsPickList.Filter = "cCode LIKE '" & sFilterString & "%'"
            
            If sColumnVariable = "Code" Then
                Set itmX = lvwList.ColumnHeaders.Add(, , "Client ID")
                Set itmX = lvwList.ColumnHeaders.Add(, , "Client Name")
                lvwList.ColumnHeaders(1).Width = "2500"
                lvwList.ColumnHeaders(2).Width = "5440"
    
                Do Until rsPickList.EOF
                    Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cCode) & "")
                    itmX.SubItems(1) = Trim(rsPickList!cName) & ""
                    rsPickList.MoveNext
                Loop
            ElseIf sColumnVariable = "Desc" Then
                Set itmX = lvwList.ColumnHeaders.Add(, , "Client Name")
                Set itmX = lvwList.ColumnHeaders.Add(, , "Client ID")
                lvwList.ColumnHeaders(1).Width = "5440"
                lvwList.ColumnHeaders(2).Width = "2500"
    
                Do Until rsPickList.EOF
                    Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cName) & "")
                    itmX.SubItems(1) = Trim(rsPickList!cCode) & ""
                    rsPickList.MoveNext
                Loop
            End If
        
        Case User
            Caption = "User List"
            Set itmX = lvwList.ColumnHeaders.Add(, , "User Name")
            Set itmX = lvwList.ColumnHeaders.Add(, , "User ID")
            lvwList.ColumnHeaders(1).Width = "5440"
            lvwList.ColumnHeaders(2).Width = "2500"
            
            sSQL = "select UserId, FirstName from SEC_USER"
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!FirstName) & "")
                itmX.SubItems(1) = Trim(rsPickList!UserID) & ""
                
                rsPickList.MoveNext
            Loop
            
        Case Company
            Caption = "Unit List"
            Set itmX = lvwList.ColumnHeaders.Add(, , "Company Name")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Company ID")
            lvwList.ColumnHeaders(1).Width = "5440"
            lvwList.ColumnHeaders(2).Width = "2500"
            
            sSQL = "select cCompanyId, cCompanyName from COMPANY "
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cCompanyName) & "")
                itmX.SubItems(1) = Trim(rsPickList!cCompanyID) & ""
                
                rsPickList.MoveNext
            Loop
            
        Case Role
            Caption = "Role List"
            Set itmX = lvwList.ColumnHeaders.Add(, , "Role Name")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Role Code")
            lvwList.ColumnHeaders(1).Width = "5440"
            lvwList.ColumnHeaders(2).Width = "2500"
            
            sSQL = "select cRole, cRoleName from ROLES "
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cRoleName) & "")
                itmX.SubItems(1) = Trim(rsPickList!CROLE) & ""
                
                rsPickList.MoveNext
            Loop
        Case PurchaseInvoice
            Caption = "Invoice List"
            Set itmX = lvwList.ColumnHeaders.Add(, , "WRR Number")
            Set itmX = lvwList.ColumnHeaders.Add(, , "WRR Gross")
            lvwList.ColumnHeaders(1).Width = "5440"
            lvwList.ColumnHeaders(2).Width = "2500"
            
            sSQL = "select cWRRNo, nGross from WRR where cCode = '" & sClient & "' AND cCompanyID = '" & COID & "'"
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cWRRNo) & "")
                itmX.SubItems(1) = Format(rsPickList!nGross, "#,##0.#0") & ""
                rsPickList.MoveNext
            Loop
        Case Bank
            Caption = "Bank List"
            Set itmX = lvwList.ColumnHeaders.Add(, , "Bank Name")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Bank ID")
            lvwList.ColumnHeaders(1).Width = "5440"
            lvwList.ColumnHeaders(2).Width = "2500"

            sSQL = "SELECT cBankID,cBankName FROM BANK WHERE cCompanyID = '" & Trim(COID) & "' ORDER BY cBankName "
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cBankName) & "")
                itmX.SubItems(1) = Trim(rsPickList!cBankID) & ""
                
                rsPickList.MoveNext
            Loop
        
        Case PRForCanvass
            List_PRForCanvass
        
        Case SalesInvoice
            Caption = "Sales Invoice List"
            Set itmX = lvwList.ColumnHeaders.Add(, , "Invoice Number")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Invoice Type")
            lvwList.ColumnHeaders(1).Width = "5440"
            lvwList.ColumnHeaders(2).Width = "2500"
            
            sSQL = "select cInvNo, cType from SALES WHERE cCode = '" & sClient & "' AND cCompanyID = '" & COID & "'"
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cInvNo) & "")
                itmX.SubItems(1) = Trim(rsPickList!cType) & ""
                
                rsPickList.MoveNext
            Loop
            
        Case CHECKS
            Caption = "Check List"
            Set itmX = lvwList.ColumnHeaders.Add(, , "Check Number")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Check Date")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Bank ID")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Check Amount")
            lvwList.ColumnHeaders(1).Width = "2000"
            lvwList.ColumnHeaders(2).Width = "1750"
            lvwList.ColumnHeaders(3).Width = "1500"
            lvwList.ColumnHeaders(4).Width = "2000"
            lvwList.ColumnHeaders(4).Alignment = lvwColumnRight
            
            sSQL = "SELECT cCheckNo, dCheckDate, cBankID, nAmount from CHECKS WHERE cCompanyID = '" & COID & "' " & RepName & " ORDER BY cCheckNo"
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cCheckNo) & "")
                itmX.SubItems(1) = Format(rsPickList!dCheckDate, "MM/dd/yyyy") & ""
                itmX.SubItems(2) = Trim(rsPickList!cBankID) & ""
                itmX.SubItems(3) = Format(rsPickList!nAmount, "#,##0.#0") & ""
                
                rsPickList.MoveNext
            Loop
        
        Case Checks_Deposit, Checks_Deposit_Ref
        If mType = Checks_Deposit Then
            OR_Str = "Check Number"
            Check_Str = "OR Number"
            
            OR_Num = "cCheckNo"
            Check_Num = "cTranNo"
            
        ElseIf mType = Checks_Deposit_Ref Then
            OR_Str = "OR Number"
            Check_Str = "Check Number"
            
            OR_Num = "cTranNo"
            Check_Num = "cCheckNo"
        
        End If
        
            Caption = "Check List (For Deposit)"
            Set itmX = lvwList.ColumnHeaders.Add(, , OR_Str)
            Set itmX = lvwList.ColumnHeaders.Add(, , "Check Date")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Bank ID")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Check Amount")
            Set itmX = lvwList.ColumnHeaders.Add(, , Check_Str)
            
            lvwList.ColumnHeaders(1).Width = "2000"
            lvwList.ColumnHeaders(2).Width = "1750"
            lvwList.ColumnHeaders(3).Width = "1500"
            lvwList.ColumnHeaders(4).Width = "2000"
            lvwList.ColumnHeaders(4).Alignment = lvwColumnRight
            lvwList.ColumnHeaders(5).Width = "1500"
            
            'sSQL = "SELECT cCheckNo, dCheckDate, cBankID, nAmount,cTranNo  from CHECKS WHERE cCompanyID = '" & COID & "' " & RepName & " AND dCheckDate <= '" & Format(frmARDepositSlip.dtbDate.Text, "MM/dd/yyyy") & "' ORDER BY  dCheckDate ASC "
            
            
            sSQL = "SELECT cCheckNo, dCheckDate, cBankID, nAmount,A.cTranNo  " & _
                        " From CHECKS A" & _
                        " LEFT JOIN (SELECT cTranNo,lCancelled FROM INTAR " & _
                        " Union All " & _
                        " SELECT cTranNo,lCancelled FROM PR) B ON A.cTranNo = B.cTranNo " & _
                        " WHERE cCompanyID = '000-00'  AND lDeposited = 0  /* AND dCheckDate <= '01/24/2011' */ " & _
                        " AND cCheckNo  NOT IN " & _
                        " (SELECT cCheckNo FROM DEPOSIT_T B LEFT JOIN DEPOSIT A ON B.cTranNo = A.cTranNo AND B.cCompanyID = A.cCompanyID WHERE A.lApproved = 1 ) AND  B.lCancelled = 0 " & _
                        " AND A.cTranNo NOT IN ('AR-3297','AR-3729','OR-40650','OR-40731','OR-41328','000012','000013','OR-40607') ORDER BY  dCheckDate DESC "
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList.Fields(OR_Num)) & "")
                itmX.SubItems(1) = Format(rsPickList!dCheckDate, "MM/dd/yyyy") & ""
                itmX.SubItems(2) = Trim(rsPickList!cBankID) & ""
                itmX.SubItems(3) = Format(rsPickList!nAmount, "#,##0.#0") & ""
                itmX.SubItems(4) = Trim(rsPickList.Fields(Check_Num)) & ""
                rsPickList.MoveNext
            Loop
        Case CheckBook
            Caption = "Bank's CheckBook List"
            Set itmX = lvwList.ColumnHeaders.Add(, , "Checkbook Code")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Current Check No.")
            lvwList.ColumnHeaders(1).Width = "5440"
            lvwList.ColumnHeaders(2).Width = "2500"
            
            sSQL = "SELECT B.cCheckBookNo, B.cCheckCurrent FROM BANK A, BANKCHECK B WHERE A.cBankID = B.cBankID AND CAST(B.cCheckCurrent AS NUMERIC) <= CAST(B.cCheckTo AS NUMERIC) AND B.cBankID = '" & RepName & "' "
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cCheckBookno))
                itmX.SubItems(1) = rsPickList!cCheckcurrent
                
                rsPickList.MoveNext
            Loop
                    
        Case CheckBookNo
            Caption = "Bank's CheckBook List"
            Set itmX = lvwList.ColumnHeaders.Add(, , "Checkbook Code")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Current Check No.")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Bank ID")
            lvwList.ColumnHeaders(1).Width = "3000"
            lvwList.ColumnHeaders(2).Width = "2500"
            lvwList.ColumnHeaders(3).Width = "2000"
            
            sSQL = "SELECT B.cCheckBookNo, B.cCheckCurrent, A.cBankID FROM BANK A, BANKCHECK B WHERE A.cBankID = B.cBankID AND CAST(B.cCheckCurrent AS INT) <= CAST(B.cCheckTo AS INT) ORDER BY B.cCheckBookNo"
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cCheckBookno))
                itmX.SubItems(1) = rsPickList!cCheckcurrent
                itmX.SubItems(2) = rsPickList!cBankID
                
                rsPickList.MoveNext
            Loop
            
        Case BANKCOMP
            Caption = "Bank - Company List"
            Set itmX = lvwList.ColumnHeaders.Add(, , "Bank Code")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Bank Name")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Checkbook No.")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Current Check No.")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Bank Account No.")
            lvwList.ColumnHeaders(1).Width = "1000"
            lvwList.ColumnHeaders(2).Width = "3000"
            lvwList.ColumnHeaders(3).Width = "1000"
            lvwList.ColumnHeaders(4).Width = "1500"
            lvwList.ColumnHeaders(5).Width = "1500"
            
            
'            sSQL = "SELECT A.cBankID, A.cBankName, A.cCheckbookNo, B.cCheckCurrent, isnull(B.cBankAcctNo, '') as cBankAcctNo " & _
'                    "FROM BANKCOMP A " & _
'                    "LEFT OUTER JOIN BANKCHECK B ON A.cBankID = B.cBankID AND A.cCheckBookNo = B.cCheckBookNo " & _
'                    "WHERE A.cCompanyID = '" & sUnitId & "' " & _
'                    "ORDER BY A.cBankName, A.cCheckBookNo "
'I added this code
            sSQL = "SELECT A.cBankID, A.cBankName, A.cCheckbookNo, B.cCheckCurrent, isnull(B.cBankAcctNo, '') as cBankAcctNo " & _
                    "FROM BANKCOMP A " & _
                    "LEFT OUTER JOIN BANKCHECK B ON A.cBankID = B.cBankID AND A.cCheckBookNo = B.cCheckBookNo " & _
                    "WHERE A.cCompanyID = '" & sUnitId & "' AND  cast(B.cCheckCurrent AS NUMERIC (20, 2)) <= cast(B.cCheckTo as NUMERIC (20,2))" & _
                    "ORDER BY A.cBankName, A.cCheckBookNo "
                    
                    
                    
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cBankID))
                itmX.SubItems(1) = rsPickList!cBankName
                itmX.SubItems(2) = rsPickList!cCheckBookno
                itmX.SubItems(3) = rsPickList!cCheckcurrent
                itmX.SubItems(4) = rsPickList!cBankAcctNo
                
                rsPickList.MoveNext
            Loop
            
        Case Voucher
            Caption = "Voucher List"
            Set itmX = lvwList.ColumnHeaders.Add(, , "Voucher Number")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Voucher Date")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Voucher Amount")
            lvwList.ColumnHeaders(1).Width = "1500"
            lvwList.ColumnHeaders(2).Width = "1500"
            lvwList.ColumnHeaders(3).Width = "1500"
            
            sSQL = "select cTranNo, dDate, (nTDebit - nTCredit) as nAmount from VOUCHER where lPosted = 0 order by cTranNo "
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cTranNo))
                itmX.SubItems(1) = rsPickList!dDate
                itmX.SubItems(2) = rsPickList!nAmount
                
                rsPickList.MoveNext
            Loop
        
        Case IssuedCheck
            Caption = "Issued Check List"
            Set itmX = lvwList.ColumnHeaders.Add(, , "Check No.")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Bank ID")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Amount")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Check Date")
            lvwList.ColumnHeaders(1).Width = "1500"
            lvwList.ColumnHeaders(2).Width = "1500"
            lvwList.ColumnHeaders(3).Width = "1500"
            lvwList.ColumnHeaders(4).Width = "1500"
            
            sSQL = "SELECT cCheckNo, cBankID, nAmount, dCheckDate FROM ISSUED " & _
                    "WHERE cCompanyID = '" & Trim(sUnitId) & "' AND lCleared = 0 AND lCancelled = 0 " & _
                    "AND cCheckNo not in (select cCheckNo from cleared_t WHERE cBankID = '" & RTrim(vStrContainer1) & "') " & _
                    "AND cBankID = '" & RTrim(vStrContainer1) & "' ORDER BY cBankID, cCheckNo "
                    
                    
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF

                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cCheckNo))
                itmX.SubItems(1) = rsPickList!cBankID
                itmX.SubItems(2) = Format(rsPickList!nAmount, "#,##0.#0") & ""
                itmX.SubItems(3) = rsPickList!dCheckDate
                
                rsPickList.MoveNext
                
            Loop
            
        Case SecDepartment
            Caption = "Security List (Department)"
            Set itmX = lvwList.ColumnHeaders.Add(, , "Description")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Department ID")
            lvwList.ColumnHeaders(1).Width = "5440"
            lvwList.ColumnHeaders(2).Width = "2500"

            sSQL = "SELECT DeptID,Description FROM SEC_DEPARTMENT"
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!Description) & "")
                itmX.SubItems(1) = Trim(rsPickList!DeptID) & ""
                
                rsPickList.MoveNext
            Loop
        
        Case SecRole
            Caption = "Security List (User Role)"
            Set itmX = lvwList.ColumnHeaders.Add(, , "Description")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Role ID")
            lvwList.ColumnHeaders(1).Width = "5440"
            lvwList.ColumnHeaders(2).Width = "2500"

            sSQL = "SELECT RoleID,Description FROM SEC_ROLE"
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!Description) & "")
                itmX.SubItems(1) = Trim(rsPickList!RoleID) & ""
                
                rsPickList.MoveNext
            Loop
    
        Case APDRCR_DR
            Caption = "DR List for AP Adjustment"
            Set itmX = lvwList.ColumnHeaders.Add(, , "Transaction No.")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Date")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Type")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Gross")
            lvwList.ColumnHeaders(1).Width = "2000"
            lvwList.ColumnHeaders(2).Width = "1750"
            lvwList.ColumnHeaders(3).Width = "1500"
            lvwList.ColumnHeaders(4).Width = "2000"
            lvwList.ColumnHeaders(4).Alignment = lvwColumnRight
            
            sSQL = "SELECT cDRNo, dDate, cType, nGross FROM DR WHERE cType = 'Purchase Return' " & _
                    "AND cDelCode = '" & RepName & "'"
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cDRNo) & "")
                itmX.SubItems(1) = Format(rsPickList!dDate, "MM/dd/yyyy") & ""
                itmX.SubItems(2) = Trim(rsPickList!cType) & ""
                itmX.SubItems(3) = Format(rsPickList!nGross, "#,##0.#0") & ""
                
                rsPickList.MoveNext
            Loop
        
        Case ARDRCR_WRR
            Caption = "WRR List for AR Adjustment"
            Set itmX = lvwList.ColumnHeaders.Add(, , "Transaction No.")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Date")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Type")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Gross")
            lvwList.ColumnHeaders(1).Width = "2000"
            lvwList.ColumnHeaders(2).Width = "1750"
            lvwList.ColumnHeaders(3).Width = "1500"
            lvwList.ColumnHeaders(4).Width = "2000"
            lvwList.ColumnHeaders(4).Alignment = lvwColumnRight
            
            sSQL = "SELECT cWRRNo, dDate, cType, nGross FROM WRR WHERE cType = 'Sales Return' " & _
                    "AND cCode = '" & RepName & "'"
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cWRRNo) & "")
                itmX.SubItems(1) = Format(rsPickList!dDate, "MM/dd/yyyy") & ""
                itmX.SubItems(2) = Trim(rsPickList!cType) & ""
                itmX.SubItems(3) = Format(rsPickList!nGross, "#,##0.#0") & ""
                
                rsPickList.MoveNext
            Loop
    
        Case RFP_Voucher
            Caption = "RFP List for Voucher"
            Set itmX = lvwList.ColumnHeaders.Add(, , "Transaction No.", 2000)
            Set itmX = lvwList.ColumnHeaders.Add(, , "Date Needed", 1750)
            Set itmX = lvwList.ColumnHeaders.Add(, , "Payment For", 3000)
            
            CreateScript (RFP_Voucher)
            
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cTranNo) & "")
                itmX.SubItems(1) = Trim(rsPickList!dDateNeeded) & ""
                itmX.SubItems(2) = Trim(rsPickList!cPayfor) & ""
                rsPickList.MoveNext
            Loop
        Case Department
            Caption = "List"
            Set itmX = lvwList.ColumnHeaders.Add(, , "Description")
            Set itmX = lvwList.ColumnHeaders.Add(, , "ID")
            lvwList.ColumnHeaders(1).Width = "5440"
            lvwList.ColumnHeaders(2).Width = "2500"

            sSQL = "SELECT cDesc,cValue FROM PARAMETER_USER WHERE cCompanyID = '" & COID & "' AND cType = 'DEPARTMENT'"
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cDesc) & "")
                itmX.SubItems(1) = Trim(rsPickList!cValue) & ""
                
                rsPickList.MoveNext
            Loop
        
        Case Item
            Caption = "Item"
            Set itmX = lvwList.ColumnHeaders.Add(, , "Item No.")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Description")
            Set itmX = lvwList.ColumnHeaders.Add(, , "SR Price")
            Set itmX = lvwList.ColumnHeaders.Add(, , "WS Price")
            lvwList.ColumnHeaders(1).Width = "1500"
            lvwList.ColumnHeaders(2).Width = "3000"
            lvwList.ColumnHeaders(3).Width = "1500"
            lvwList.ColumnHeaders(4).Width = "1500"
            
            sSQL = "SELECT cItemNo, cDesc, nSRPrice, nWSPrice from ITEM where cCompanyId = '" & COID & "' " & RepName
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cItemNo))
                itmX.SubItems(1) = Trim(rsPickList!cDesc)
                itmX.SubItems(2) = Format(rsPickList!nSRPrice, "#,##0.#0")
                itmX.SubItems(3) = Format(rsPickList!nWSPrice, "#,##0.#0")
                
                rsPickList.MoveNext
            Loop
            
        Case Item_Inv
            Caption = "Product List"
            Set itmX = lvwList.ColumnHeaders.Add(, , "Product ID")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Description")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Balance")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Cost")
            lvwList.ColumnHeaders(1).Width = "1500"
            lvwList.ColumnHeaders(2).Width = "2500"
            lvwList.ColumnHeaders(3).Width = "1500"
            lvwList.ColumnHeaders(4).Width = "1500"
            lvwList.ColumnHeaders(3).Alignment = lvwColumnRight
            lvwList.ColumnHeaders(4).Alignment = lvwColumnRight
            
            sSQL = "SELECT A.cItemNo, B.cDesc, ISNULL((A.nInitial - A.nOutgoing + A.nIncoming + A.nRetSales - A.nRetPur - A.nIssues + A.nBack - A.nProdOut + A.nProdIn + A.nAdjustment - A.nSampleDemo), 0) AS nBalance, B.nAveCost " & _
                    "FROM V_ProductBalanceInquiry_Module A " & _
                    "LEFT OUTER JOIN ITEM B ON A.cItemNo = B.cItemNo AND A.cCompanyID = B.cCompanyID " & _
                    "WHERE A.cCompanyID = '" & COID & "' " & RepName
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cItemNo))
                itmX.SubItems(1) = Trim(rsPickList!cDesc)
                itmX.SubItems(2) = Format(IIf(IsNull(rsPickList!nBalance), 0, rsPickList!nBalance), "#,##0.#0")
                itmX.SubItems(3) = Format(IIf(IsNull(rsPickList!nAveCost), 0, rsPickList!nAveCost), "#,##0.#0")
                
                rsPickList.MoveNext
            Loop
            
        Case DeliveryReceipt
            Caption = "Delivery Receipt"
            Set itmX = lvwList.ColumnHeaders.Add(, , "DR No.")
            Set itmX = lvwList.ColumnHeaders.Add(, , "SO No.")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Code")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Address")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Term")
            lvwList.ColumnHeaders(1).Width = "1500"
            lvwList.ColumnHeaders(2).Width = "1500"
            lvwList.ColumnHeaders(3).Width = "1500"
            lvwList.ColumnHeaders(4).Width = "3000"
            lvwList.ColumnHeaders(5).Width = "1500"
            
            'sSQL = "select distinct(a.cdrno), b.csono, (select c.cName from client c where b.cCode = c.cCode AND C.cCompanyID = '" & COID & "') as cCode, b.cAddress, b.cTerm from dr_t a, so b Where a.crefno = b.csono and a.cCompanyId = '" & COID & "' "
            sSQL = "SELECT DISTINCT(A.cDRNo), B.cSONo, C.cName AS cCode, B.cAddress, B.cTerm, D.cTranNo " & _
                "FROM DR_T A " & _
                "LEFT OUTER JOIN SO B ON A.cRefNo = B.cSONo AND A.cCompanyID = B.cCompanyID " & _
                "LEFT OUTER JOIN CLIENT_CUSTOMER C ON B.cCode = C.cCode AND B.cCompanyID = C.cCompanyID " & _
                "LEFT OUTER JOIN TRIPTICKET_T D ON A.cDRNo = D.cDRNo AND A.cRefNo = D.cSONo AND A.cCompanyID = D.cCompanyID " & _
                "WHERE B.cSONo IS NOT NULL AND D.cTranno IS NULL AND A.cCompanyID = '" & COID & "'"
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cDRNo))
                itmX.SubItems(1) = Trim(rsPickList!csono)
                itmX.SubItems(2) = Trim(rsPickList!cCode)
                itmX.SubItems(3) = Trim(rsPickList!cAddress)
                itmX.SubItems(4) = Trim(rsPickList!cTerm)
                
                rsPickList.MoveNext
            Loop
    
        Case BankAccount
            Caption = "Bank Accounts"
            Set itmX = lvwList.ColumnHeaders.Add(, , "Bank Account No.")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Currency")
            lvwList.ColumnHeaders(1).Width = "2000"
            
            sSQL = "SELECT cBankAcctNo, cCurrencyID FROM fm_bkBankAccount WHERE cBankID = '" & sBankCode & "' order by cBankAcctNo"
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cBankAcctNo))
                itmX.SubItems(1) = rsPickList!cCurrencyID
                
                rsPickList.MoveNext
            Loop
    
        Case Project
            Caption = "Project Setup List"
            Set itmX = lvwList.ColumnHeaders.Add(, , "Project Decription")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Project No.")
            lvwList.ColumnHeaders(1).Width = "5440"
            lvwList.ColumnHeaders(2).Width = "2500"

            sSQL = "SELECT cProjDesc,cProjNo FROM PROJECT WHERE cCompanyID = '" & COID & "' AND cCode = '" & RepName & "' ORDER BY cCode "
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cProjDesc) & "")
                itmX.SubItems(1) = Trim(rsPickList!cProjNo) & ""
                
                rsPickList.MoveNext
            Loop
            
        Case APCM
            Caption = "AR - Credit Memo List"
            Set itmX = lvwList.ColumnHeaders.Add(, , "CM No.")
            Set itmX = lvwList.ColumnHeaders.Add(, , "")
            lvwList.ColumnHeaders(1).Width = "5440"
            lvwList.ColumnHeaders(2).Width = "2500"

            sSQL = "SELECT cTranNo, '' as cNull FROM AP WHERE cType = 'Credit' AND cCompanyID = '" & COID & "' AND lCancelled = 0 AND lApproved = 1 ORDER BY cTranNo"
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cTranNo) & "")
                itmX.SubItems(1) = Trim(rsPickList!cNull) & ""
                
                rsPickList.MoveNext
            Loop
            
        Case APDM
            Caption = "AR - Debit Memo List"
            Set itmX = lvwList.ColumnHeaders.Add(, , "DM No.")
            Set itmX = lvwList.ColumnHeaders.Add(, , "")
            lvwList.ColumnHeaders(1).Width = "5440"
            lvwList.ColumnHeaders(2).Width = "2500"

            sSQL = "SELECT cTranNo, '' as cNull FROM AP WHERE cType = 'Debit' AND cCompanyID = '" & COID & "' AND lCancelled = 0 AND lApproved = 1 ORDER BY cTranNo"
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cTranNo) & "")
                itmX.SubItems(1) = Trim(rsPickList!cNull) & ""
                
                rsPickList.MoveNext
            Loop
        
        Case ARCM
            Caption = "AR - Credit Memo List"
            Set itmX = lvwList.ColumnHeaders.Add(, , "CM No.")
            Set itmX = lvwList.ColumnHeaders.Add(, , "")
            lvwList.ColumnHeaders(1).Width = "5440"
            lvwList.ColumnHeaders(2).Width = "2500"

            sSQL = "SELECT cTranNo, '' as cNull FROM AR WHERE cType = 'Credit' AND cCompanyID = '" & COID & "' AND lCancelled = 0 AND lApproved = 1 ORDER BY cTranNo"
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cTranNo) & "")
                itmX.SubItems(1) = Trim(rsPickList!cNull) & ""
                
                rsPickList.MoveNext
            Loop
            
        Case ARDM
            Caption = "AR - Debit Memo List"
            Set itmX = lvwList.ColumnHeaders.Add(, , "DM No.")
            Set itmX = lvwList.ColumnHeaders.Add(, , "")
            lvwList.ColumnHeaders(1).Width = "5440"
            lvwList.ColumnHeaders(2).Width = "2500"

            sSQL = "SELECT cTranNo, '' as cNull FROM AR WHERE cType = 'Debit' AND cCompanyID = '" & COID & "' AND lCancelled = 0 AND lApproved = 1 ORDER BY cTranNo"
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cTranNo) & "")
                itmX.SubItems(1) = Trim(rsPickList!cNull) & ""
                
                rsPickList.MoveNext
            Loop
            
        Case RJOList
            List_RJOList
        
        Case Collection_Invoice
            Caption = "Invoice List"
            Set itmX = lvwList.ColumnHeaders.Add(, , "Invoice No.")
            Set itmX = lvwList.ColumnHeaders.Add(, , "")
            lvwList.ColumnHeaders(1).Width = "5440"
            lvwList.ColumnHeaders(2).Width = "2500"

            sSQL = "SELECT cInvNo, ' ' as cNull FROM SALES WHERE cCompanyID = '" & COID & "' AND lCancelled = 0 "
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cInvNo) & "")
                itmX.SubItems(1) = Trim(rsPickList!cNull) & ""
                
                rsPickList.MoveNext
            Loop

        Case Collection_OR
            Caption = "Official Receipt List"
            Set itmX = lvwList.ColumnHeaders.Add(, , "OR No.")
            Set itmX = lvwList.ColumnHeaders.Add(, , "")
            lvwList.ColumnHeaders(1).Width = "5440"
            lvwList.ColumnHeaders(2).Width = "2500"

            sSQL = "SELECT cTranNo, ' ' as cNull FROM PR WHERE cCompanyID = '" & COID & "' AND lCancelled = 0 "
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cTranNo) & "")
                itmX.SubItems(1) = Trim(rsPickList!cNull) & ""
                
                rsPickList.MoveNext
            Loop
            
        Case SI_Journal
            Caption = "Sales Invoice List"
            Set itmX = lvwList.ColumnHeaders.Add(, , "Invoice Number")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Invoice Type")
            lvwList.ColumnHeaders(1).Width = "5440"
            lvwList.ColumnHeaders(2).Width = "2500"
            
            sSQL = "select cInvNo, cType from SALES WHERE cCompanyID = '" & COID & "'"
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cInvNo) & "")
                itmX.SubItems(1) = Trim(rsPickList!cType) & ""
                
                rsPickList.MoveNext
            Loop
            
        Case ProductionInv
            Caption = "Work In Process Monitoring"
            Set itmX = lvwList.ColumnHeaders.Add(, , "Receipt No.")
            Set itmX = lvwList.ColumnHeaders.Add(, , "")
            lvwList.ColumnHeaders(1).Width = "5440"
            lvwList.ColumnHeaders(2).Width = "2500"

            CreateScript (ProductionInv)
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cTranNo) & "")
                itmX.SubItems(1) = Trim(rsPickList!cNull) & ""
                
                rsPickList.MoveNext
            Loop
            
        Case DepositJournal
            Caption = "Deposit Journal List"
            Set itmX = lvwList.ColumnHeaders.Add(, , "Deposit Journal No.")
            Set itmX = lvwList.ColumnHeaders.Add(, , "")
            lvwList.ColumnHeaders(1).Width = "5440"
            lvwList.ColumnHeaders(2).Width = "2500"

            sSQL = "SELECT DISTINCT cTranNo, ' ' as cNull FROM DEPOSIT_T WHERE cCompanyID = '" & COID & "' ORDER BY cTranNo"
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cTranNo) & "")
                itmX.SubItems(1) = Trim(rsPickList!cNull) & ""
                
                rsPickList.MoveNext
            Loop
            
        Case CMEVATEWT
            Caption = "EWT/EVAT List for Client " & cString
            Set itmX = lvwList.ColumnHeaders.Add(, , "Invoice No.")
            Set itmX = lvwList.ColumnHeaders.Add(, , "EWT/EVAT No.")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Type")
            Set itmX = lvwList.ColumnHeaders.Add(, , "Amount")
            Set itmX = lvwList.ColumnHeaders.Add(, , "AR Tran No.")
            lvwList.ColumnHeaders(1).Width = "2000"
            lvwList.ColumnHeaders(2).Width = "1600"
            lvwList.ColumnHeaders(3).Width = "1600"
            lvwList.ColumnHeaders(4).Width = "1600"
            lvwList.ColumnHeaders(5).Width = "1600"
            
            CreateScript (CMEVATEWT)

            
            rsPickList.Open sSQL, connList, adOpenKeyset
            
            Do Until rsPickList.EOF
                Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cInvNo) & "")
                itmX.SubItems(1) = Trim(rsPickList!nRefNo) & ""
                itmX.SubItems(2) = Trim(rsPickList!cType) & ""
                itmX.SubItems(3) = Trim(rsPickList!nAmount) & ""
                itmX.SubItems(4) = Trim(rsPickList!cARNo) & ""
                
                rsPickList.MoveNext
            Loop
        Case RVoucher
            ListRVoucher
            
        Case RPurchaseRequest
            List_RPurchaseRequest
    
    End Select
    
    If rsPickList.State = adStateOpen Then rsPickList.Close
    Set rsPickList = Nothing
    Set connList = Nothing
    
    lvwList.Checkboxes = False
    Select Case mType
        Case Item
            cmbAll.Visible = True
            lvwList.Checkboxes = True
        Case Item_Inv
            cmbAll.Visible = True
            lvwList.Checkboxes = True
        Case Product, ProductPM
            If RepName = "CommissionCriteria" Or RepName = "PriceMatrix" Then
                cmbAll.Visible = True
                lvwList.Checkboxes = True
            End If
        Case Client_Customer
            If RepName = "CommissionCriteria" Then
                cmbAll.Visible = True
                lvwList.Checkboxes = True
            End If
        Case Salesman
            If RepName = "CommissionCriteria" Then
                cmbAll.Visible = True
                lvwList.Checkboxes = True
            End If
        Case IssuedCheck, Checks_Deposit, RVoucher, RPurchaseRequest, Checks_Deposit_Ref
            cmbAll.Visible = True
            lvwList.Checkboxes = True
     
     

    End Select
    
    If Not lModal Then Unload frmWait
    lModal = False
    
    lvwList.Visible = True
    
    lvwList_ColumnClick lvwList.ColumnHeaders(1)
    
    If lvwList.ListItems.Count > 0 Then lvwList.SelectedItem = lvwList.ListItems(1)
    
    MousePointer = vbDefault
    
TheSource:
    'The connection cannot be used to perform this operation. It is either closed or invalid in this context.
    If Err.Number = 3709 Then
        Set connList = Nothing
        Unload frmWait
        MousePointer = vbDefault
        Unload Me
    End If
End Sub

Private Sub optCode_Click()
    sColumnVariable = "Code"
    ShowForm
End Sub

Private Sub optName_Click()
    sColumnVariable = "Desc"
    ShowForm
End Sub

Private Sub txtFind_Change()
    mCriteria = Me.txtFind
End Sub

Private Sub List_RJOList()
    Caption = "Request for Job Order List"
    Set itmX = lvwList.ColumnHeaders.Add(, , "RJO No.")
    Set itmX = lvwList.ColumnHeaders.Add(, , "RJO Type")
    lvwList.ColumnHeaders(1).Width = "2500"
    lvwList.ColumnHeaders(2).Width = "5440"

    sSQL = "SELECT cTranNo, cType FROM JOREQUEST WHERE cCompanyID = '" & COID & "' AND lCancelled = 0 " & _
            "AND cTranNo NOT IN (SELECT DISTINCT cRefRJONo from JO WHERE cCompanyID = '" & COID & "' AND lCancelled = 0)"
    rsPickList.Open sSQL, connList, adOpenKeyset
    
    Do Until rsPickList.EOF
        Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cTranNo) & "")
        itmX.SubItems(1) = Trim(rsPickList!cType) & ""
        
        rsPickList.MoveNext
    Loop
End Sub

Private Sub List_PRForCanvass()
 Caption = "Purchase Request for Canvass [" & RepName & "]"
    Set itmX = lvwList.ColumnHeaders.Add(, , "PR Number")
    Set itmX = lvwList.ColumnHeaders.Add(, , "PR Date")
    lvwList.ColumnHeaders(1).Width = "2500"
    lvwList.ColumnHeaders(2).Width = "5440"
    
    sSQL = "SELECT DISTINCT A.cCompanyID, A.cPRNo, B.dDate " & _
            "FROM REQUISITION_T A " & _
            "LEFT OUTER JOIN REQUISITION B ON A.cCompanyID = B.cCompanyID AND A.cPRNo = B.cPRNo " & _
            "WHERE A.cCompanyID = '" & COID & "' AND B.lCancelled = 0 AND A.cItemType = '" & RepName & "' " & _
            "AND A.nIdentity NOT IN " & _
            "(SELECT DISTINCT A.nRefIdentity AS nIdentity " & _
            "FROM CANVASS A " & _
            "WHERE A.lCancelled = 0 AND A.cCompanyID = '" & COID & "') "
    rsPickList.Open sSQL, connList, adOpenKeyset
    
    Do Until rsPickList.EOF
        Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cPRNo) & "")
        itmX.SubItems(1) = Format(rsPickList!dDate, "yyyy/mm/dd") & ""
        
        rsPickList.MoveNext
    Loop
End Sub

Private Sub List_PRItemForCanvass()
    Caption = "Purchase Request Items for Canvass [" & RepName & "]"
    
    sSQL = "SELECT DISTINCT A.cCompanyID, A.cPRNo, B.dDate, A.cItemNo, A.cDesc, A.nQty, A.nIdentity " & _
            "FROM REQUISITION_T A " & _
            "LEFT OUTER JOIN REQUISITION B ON A.cCompanyID = B.cCompanyID AND A.cPRNo = B.cPRNo " & _
            "WHERE A.cCompanyID = '" & COID & "' AND B.lApproved =  1 and B.lCancelled = 0 AND A.cItemType = '" & RepName & "' " & _
            "AND B.cPRNo = '" & cString & "' AND A.nIdentity NOT IN " & _
            "(SELECT DISTINCT A.nRefIdentity AS nIdentity " & _
            "FROM CANVASS A " & _
            "WHERE A.lCancelled = 0 AND A.cCompanyID = '" & COID & "') "
    rsPickList.Open sSQL, connList, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If sColumnVariable = "Code" Then
        Set itmX = lvwList.ColumnHeaders.Add(, , "Product ID")
        Set itmX = lvwList.ColumnHeaders.Add(, , "Product Description")
        Set itmX = lvwList.ColumnHeaders.Add(, , "Quantity Needed")
        Set itmX = lvwList.ColumnHeaders.Add(, , "Ref. Identity")
        lvwList.ColumnHeaders(1).Width = "2000"
        lvwList.ColumnHeaders(2).Width = "3500"
        lvwList.ColumnHeaders(3).Width = "2000"
        lvwList.ColumnHeaders(4).Width = "0"
        lvwList.ColumnHeaders(3).Alignment = lvwColumnRight
        Do Until rsPickList.EOF
            Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cItemNo) & "")
            itmX.SubItems(1) = Trim(rsPickList!cDesc) & ""
            itmX.SubItems(2) = Format(rsPickList!nQty, "#,##0.#0") & ""
            itmX.SubItems(3) = rsPickList!nIdentity & ""
            rsPickList.MoveNext
        Loop
    ElseIf sColumnVariable = "Desc" Then
        Set itmX = lvwList.ColumnHeaders.Add(, , "Product Description")
        Set itmX = lvwList.ColumnHeaders.Add(, , "Product ID")
        Set itmX = lvwList.ColumnHeaders.Add(, , "Quantity Needed")
        Set itmX = lvwList.ColumnHeaders.Add(, , "Ref. Identity")
        lvwList.ColumnHeaders(1).Width = "3500"
        lvwList.ColumnHeaders(2).Width = "2000"
        lvwList.ColumnHeaders(3).Width = "2000"
        lvwList.ColumnHeaders(4).Width = "0"
        lvwList.ColumnHeaders(3).Alignment = lvwColumnRight
            
        Do Until rsPickList.EOF
            Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cDesc) & "")
            itmX.SubItems(1) = Trim(rsPickList!cItemNo) & ""
            itmX.SubItems(2) = Format(rsPickList!nQty, "#,##0.#0") & ""
            itmX.SubItems(3) = rsPickList!nIdentity & ""
            rsPickList.MoveNext
        Loop
    End If
End Sub

Private Sub CreateScript(QType As eType)
    Select Case QType
        Case CMEVATEWT
            sSQL = "SELECT A.* FROM " & _
                    "(SELECT A.cCompanyID, A.cCode, A.cInvNo, A.cARNo, A.nEWTNo AS nRefNo, A.nEWTAmt AS nAmount, 'EWT' AS cType " & _
                    "FROM EWT A WHERE A.nEWTNo NOT IN " & _
                    "(SELECT A.nRefNo FROM CM_EWTEVAT_T A " & _
                    "LEFT OUTER JOIN CM_EWTEVAT B ON A.cCompanyID = B.cCOmpanyID AND A.cTranNo = B.cTranNo " & _
                    "WHERE A.cCompanyID = '" & COID & "' AND A.cType = 'EWT' AND B.lCancelled = 0) " & _
                    "UNION ALL " & _
                    "SELECT A.cCompanyID, A.cCode, A.cInvNo, A.cARNo, A.nEVATNo AS nRefNo, A.nEVATAmt AS nAmount, 'EVAT' AS cType " & _
                    "FROM EVAT A WHERE A.nEVATNo NOT IN " & _
                    "(SELECT A.nRefNo FROM CM_EWTEVAT_T A " & _
                    "LEFT OUTER JOIN CM_EWTEVAT B ON A.cCompanyID = B.cCOmpanyID AND A.cTranNo = B.cTranNo " & _
                    "WHERE A.cCompanyID = '" & COID & "' AND A.cType = 'EVAT' AND B.lCancelled = 0)) A " & _
                    "LEFT OUTER JOIN AR B ON A.cCompanyID = B.cCompanyID AND A.cARNo = B.cTranNo " & _
                    "WHERE B.lApproved = 1 AND A.cCode = '" & RepName & "' AND A.cCompanyID = '" & COID & "' " & _
                    "ORDER BY A.cInvNo, A.nRefNo"
    
        Case RFP_Voucher
            sSQL = "EXEC [rsp_FilterRFPList]  '" & COID & "'  ,   '" & RepName & "' "
            'sSQL = "SELECT cTranNo, dDateNeeded, cPayfor FROM RFP WHERE cCompanyID = '" & COID & "' " & _
                    " and cCode = '" & RepName & "' and lApproved = 1 and cTranNo not in (select cRFPNo from VOUCHER where " & _
                    " cCompanyID = '" & COID & "' and cCode = '" & RepName & "' and cRFPNo is not null)"
        Case ProductionInv
            sSQL = "SELECT cTranNo, ' ' as cNull FROM Production WHERE cCompanyID = '" & COID & "' ORDER BY cTranNo"
        Case RVoucher
            sSQL = "SELECT cTranNo AS cCode, cName  FROM Voucher WHERE cCompanyID = '" & COID & "' ORDER BY cTranNo"
         Case RPurchaseRequest
            sSQL = "SELECT cPRNo AS cCode, dDate  FROM requisition WHERE cCompanyID = '" & COID & "' ORDER BY cPRNo"
        Case Else
            sSQL = ""
        
    End Select
End Sub

Private Sub ListRVoucher()
            Caption = "Check Voucher"
            CreateScript (RVoucher)
            rsPickList.Open sSQL, connList, adOpenForwardOnly, adLockReadOnly, adCmdText
            If sFilterString <> "" Then rsPickList.Filter = "cTranNo LIKE '" & sFilterString & "%'"
            If sColumnVariable = "Code" Then
    
                Set itmX = lvwList.ColumnHeaders.Add(, , "CV No.")
                Set itmX = lvwList.ColumnHeaders.Add(, , "Payee")
                lvwList.ColumnHeaders(1).Width = "2500"
                lvwList.ColumnHeaders(2).Width = "5440"

                Do Until rsPickList.EOF
                    Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cCode) & "")
                    itmX.SubItems(1) = Trim(rsPickList!cName) & ""
                    rsPickList.MoveNext
                Loop
            ElseIf sColumnVariable = "Desc" Then
                Set itmX = lvwList.ColumnHeaders.Add(, , "Payee")
                Set itmX = lvwList.ColumnHeaders.Add(, , "CV No.")
                lvwList.ColumnHeaders(1).Width = "5440"
                lvwList.ColumnHeaders(2).Width = "2500"

                Do Until rsPickList.EOF
                    Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cName) & "")
                    itmX.SubItems(1) = Trim(rsPickList!cCode) & ""
                    rsPickList.MoveNext
                Loop
            End If
End Sub

Private Sub List_RPurchaseRequest()
 Caption = "Purchase Request for Canvass [" & RepName & "]"
    Set itmX = lvwList.ColumnHeaders.Add(, , "PR Number")
    Set itmX = lvwList.ColumnHeaders.Add(, , "PR Date")
    lvwList.ColumnHeaders(1).Width = "2500"
    lvwList.ColumnHeaders(2).Width = "5440"
    
    sSQL = "SELECT DISTINCT A.cCompanyID, A.cPRNo, B.dDate " & _
            "FROM REQUISITION_T A " & _
            "LEFT OUTER JOIN REQUISITION B ON A.cCompanyID = B.cCompanyID AND A.cPRNo = B.cPRNo " & _
            "WHERE A.cCompanyID = '" & COID & "' AND B.lCancelled = 0 "
            '& _
'            "AND A.nIdentity NOT IN " & _
'            "(SELECT DISTINCT A.nRefIdentity AS nIdentity " & _
'            "FROM CANVASS A " & _
'            "WHERE A.lCancelled = 0 AND A.cCompanyID = '" & COID & "') "
    rsPickList.Open sSQL, connList, adOpenKeyset
    
    Do Until rsPickList.EOF
        Set itmX = lvwList.ListItems.Add(, , Trim(rsPickList!cPRNo) & "")
        itmX.SubItems(1) = Format(rsPickList!dDate, "yyyy/mm/dd") & ""
        
        rsPickList.MoveNext
    Loop
End Sub


