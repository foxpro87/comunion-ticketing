VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{B3FB64BF-91F9-11D7-A482-0008A14158BC}#2.22#0"; "ITGControls.ocx"
Begin VB.Form frmITGPicker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ITGPicker"
   ClientHeight    =   3945
   ClientLeft      =   -105
   ClientTop       =   -120
   ClientWidth     =   7425
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   7425
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraSearch 
      Caption         =   "Primary Search Column"
      Height          =   495
      Left            =   1905
      TabIndex        =   4
      Top             =   3375
      Width           =   2670
      Begin VB.OptionButton optCode 
         Caption         =   "ID/Code"
         Height          =   255
         Left            =   180
         TabIndex        =   6
         Top             =   180
         Width           =   915
      End
      Begin VB.OptionButton optName 
         Caption         =   "Name/Desc"
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   180
         Width           =   1230
      End
   End
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   45
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   3000
      Width           =   4515
   End
   Begin MSDataGridLib.DataGrid dtgList 
      Height          =   2865
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7410
      _ExtentX        =   13070
      _ExtentY        =   5054
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin ITGControls.ITGCommandButton cmdCancel 
      Height          =   345
      Left            =   6120
      TabIndex        =   2
      Top             =   3525
      Width           =   1230
      _ExtentX        =   2170
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
      Top             =   3525
      Width           =   1230
      _ExtentX        =   2170
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
End
Attribute VB_Name = "frmITGPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'IT Group Inc. 2005.09.23

Option Explicit

Enum ePickType
    QAccounts
    QBank
    QCustomer
    QCustomerExp
    QDistributor
    QDivision
    QEmployee
    QEmployeeDriver
    QFixedAsset 'All Assets
    QFixedAssetVehicle
    QFixedAssetVessel
    QFixedAssetEquipment
    QFixedAssetOthers
    QGroupings
    QLocation 'Setup for Trucker Route
    QLocationArea
    QLocationCity
    QLocationProv
    QMarketSegment
    QOCustomer 'Other Customer
    QPettyCashAd
    QPriceMatrix
    QProduct
    QProductDivision
    QProfitCenter
    QLead
    QProject
    QRepProject
    QReceipt
    QSalesman
    QSalesmanDiv
    QSecDepartment
    QSupplier
    QSupplierNT
    QSupply
    QServices
    QServiceType
    QTradeShow
    QTripTicket
    QTrucker
    QTruckerRoute
    QWarehouse
    QBranch
    QDepartment
    QWRR
    QSummaryExpense
    QInvoice
    QBilledTo   'Add
    QEMPSUP_NT 'Add
    QOR_AR 'Add
    QBusinessType
    QServiceType2
End Enum

Public zType As ePickType
Public zCode As String
Public zName As String
Public sColumnVariable As String
Public sTypeVariable As String ' Myk
Public sPCCode As String ' Myk
Public cCodeCustomer As String


Private rsPickList As New ADODB.Recordset
Private oConnection As New clsConnection
Private connList As ADODB.Connection

Private Sub cmdcancel_Click()
    Unload Me
    Set frmITGPicker = Nothing
End Sub

Private Sub cmdOK_Click()
    SelectOK
    Unload Me
End Sub

Private Sub dtgList_DblClick()
    If rsPickList.RecordCount <> 0 Then SelectOK
End Sub

Private Sub dtgList_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
        Set frmITGPicker = Nothing
    ElseIf KeyAscii = 8 Then
        If Trim(txtFind.Text) <> "" Then
            txtFind = Mid(txtFind.Text, 1, Len(txtFind.Text) - 1)
            If txtFind <> "" Then
                FilterString True
            Else
                FilterString False
            End If
        End If
    ElseIf KeyAscii = 13 Then
        SelectOK
    ElseIf KeyAscii = 39 Then
        KeyAscii = 0
    Else
        txtFind = txtFind.Text + UCase(Chr(KeyAscii))
        If Trim(txtFind) <> "" Then
            FilterString True
            dtgList.SetFocus
        End If
    End If
End Sub

Private Sub Form_Activate()
    If zType = QAccounts Or zType = QServices Or zType = QServiceType Or zType = QDivision Or zType = QServiceType2 Then
        optCode.Value = True
    Else
        optName.Value = True
    End If
End Sub

Private Sub Form_Load()
    lPickListActive = True
    txtFind.Text = sFilterString
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsPickList = Nothing
    Set connList = Nothing
    RepName = Empty
    sFilterString = Empty
    Me.sTypeVariable = Empty
    lPickListActive = False
End Sub

Private Sub SelectOK()
    zCode = ""
    zName = ""
    vStrContainer1 = ""
    
    If rsPickList.RecordCount = 0 Then Exit Sub
    
    Select Case zType
        Case QCustomer, QSalesman, QSupplier, QSupplierNT, QOCustomer, QCustomerExp, QSalesmanDiv, QEMPSUP_NT
            zCode = rsPickList!cCode
            zName = rsPickList!cName
        Case QBusinessType
            zName = rsPickList!cDesc
        Case QBilledTo
            zCode = rsPickList!cCode
            zName = rsPickList!cNameBilled
        Case QSecDepartment
            zCode = rsPickList!DeptID
            zName = rsPickList!Description
        Case QSupply
            zCode = rsPickList!cSupplyNo
            zName = rsPickList!cDesc
        Case QPettyCashAd
            zCode = rsPickList!cTranNo
            zName = rsPickList!cRequestedBy
        Case QTripTicket
            zCode = rsPickList!cTranNo
            zName = rsPickList!cVanNo
        Case QTrucker
            zCode = rsPickList!cTruckerCode
            zName = rsPickList!cName
        Case QTruckerRoute
            zCode = rsPickList!cRouteCode
            zName = rsPickList!cOrigin
        Case QTradeShow
            zCode = rsPickList!cCode
            zName = rsPickList!cName
        Case QDistributor
            zCode = rsPickList!cCode
            zName = rsPickList!cName
        Case QAccounts
            If GetValueFrTable("cType", "ACCOUNT", "cAcctNo = '" & rsPickList!cAcctNo & "'") = "General" Then
                If MsgBox("'General Account' has been selected. Do you want to continue?", vbYesNo + vbInformation, "Comunion") = vbYes Then
                    zCode = rsPickList!cAcctNo
                    zName = rsPickList!cTitle
                End If
            Else
                zCode = rsPickList!cAcctNo
                zName = rsPickList!cTitle
            End If
        Case QProduct
            zCode = rsPickList!cItemNo
            zName = rsPickList!cDesc
        Case QPriceMatrix
            zCode = rsPickList!cPMID
            zName = rsPickList!cDesc
        Case QLead
            zCode = rsPickList!cLeadID
            zName = rsPickList!cCompany
        Case QProject, QRepProject
            zCode = rsPickList!cProjNo
            zName = rsPickList!cProjTitle
        Case QEmployee
            zCode = rsPickList!cEmpCode
            zName = rsPickList!cEmpName
        Case QEmployeeDriver
            zCode = rsPickList!cEmpCode
            zName = rsPickList!cEmpName
        Case QLocation
            zCode = rsPickList!cID
            zName = rsPickList!cLocation
        Case QFixedAssetVehicle
            zCode = rsPickList!cAssetNo
            zName = rsPickList!cDesc
        Case QFixedAsset
            zCode = rsPickList!cAssetNo
            zName = rsPickList!cDesc
        Case QFixedAssetVessel
            zCode = rsPickList!cAssetNo
            zName = rsPickList!cDesc
        Case QFixedAssetEquipment
            zCode = rsPickList!cAssetNo
            zName = rsPickList!cDesc
        Case QFixedAssetOthers
            zCode = rsPickList!cAssetNo
            zName = rsPickList!cDesc
        Case QBank
            zCode = rsPickList!cBankID
            zName = rsPickList!cBankName
        Case QLocationCity
            zCode = rsPickList!cID
            zName = rsPickList!cLocation
        Case QLocationProv
            zCode = rsPickList!cID
            zName = rsPickList!cLocation
        Case QLocationArea
            zCode = rsPickList!cID
            zName = rsPickList!cLocation
        Case QMarketSegment, QProductDivision
            zCode = rsPickList!cClassCode
            zName = rsPickList!cDescription
        Case QProfitCenter
            zCode = rsPickList!cPCCode
            zName = rsPickList!cDescription
        Case QWarehouse
            zCode = rsPickList!cWH
            zName = rsPickList!cName
        Case QGroupings
            zCode = rsPickList!cID
            zName = rsPickList!cDescription
        Case QDivision
            zCode = rsPickList!cDivisionID
            zName = rsPickList!cDivName
        Case QServices
            zCode = rsPickList!cServiceID
            zName = rsPickList!cServName
        Case QServiceType
            zCode = rsPickList!cServiceTypeID
            zName = rsPickList!cServiceTypeName
        Case QServiceType2
            zCode = rsPickList!cServiceTypeID
            zName = rsPickList!cServiceTypeName
        Case QWRR
            zCode = rsPickList!cWRRNo
            zName = rsPickList!cName
        Case QSummaryExpense
            zCode = rsPickList!cTranNo
            zName = rsPickList!cPJMNo
        Case QInvoice
            zCode = rsPickList!cInvNo
            zName = rsPickList!cCode
        Case QReceipt
            zCode = rsPickList!cTranNo
            zName = rsPickList!cCode
            vStrContainer1 = rsPickList!cTranNo
        Case QOR_AR
            zCode = rsPickList!cTranNo
            zName = rsPickList!nAmount
            vStrContainer1 = rsPickList!cTranNo
    End Select
    Unload Me
End Sub

Private Sub FilterString(lReset As Boolean)
    With rsPickList
        If sColumnVariable = "Code" Then
            If lReset = True Then
                Select Case zType
                    Case QCustomer, QSalesman, QSalesmanDiv, QSupplier, QSupplierNT, QOCustomer, QCustomerExp, QEMPSUP_NT
                        .Filter = "cCode like '" & Trim(txtFind.Text) & "%'"
                    Case QBusinessType
                        .Filter = "cDesc like '" & Trim(txtFind.Text) & "%'"
                    Case QBilledTo
                        .Filter = "cNameBilled like '" & Trim(txtFind.Text) & "%'"
                    Case QSecDepartment
                        .Filter = "DeptID like '" & Trim(txtFind.Text) & "%'"
                    Case QSupply
                        .Filter = "cSupplyNo like '" & Trim(txtFind.Text) & "%'"
                    Case QPettyCashAd
                        .Filter = "cTranNo like '" & Trim(txtFind.Text) & "%'"
                    Case QTripTicket
                        .Filter = "cTranNo like '" & Trim(txtFind.Text) & "%'"
                    Case QTrucker
                        .Filter = "cTruckerCode like '" & Trim(txtFind.Text) & "%'"
                    Case QTruckerRoute
                        .Filter = "cRouteCode like '" & Trim(txtFind.Text) & "%'"
                    Case QTradeShow
                        .Filter = "cCode like '" & Trim(txtFind.Text) & "%'"
                    Case QDistributor
                        .Filter = "cCode like '" & Trim(txtFind.Text) & "%'"
                    Case QAccounts
                        .Filter = "cAcctNo like '" & Trim(txtFind.Text) & "%'"
                    Case QProduct
                        .Filter = "cItemNo like '" & Trim(txtFind.Text) & "%'"
                    Case QPriceMatrix
                        .Filter = "cPMID like '" & Trim(txtFind.Text) & "%'"
                    Case QEmployee
                        .Filter = "cEmpCode like '" & Trim(txtFind.Text) & "%'"
                    Case QEmployeeDriver
                        .Filter = "cEmpCode like '" & Trim(txtFind.Text) & "%'"
                    Case QLocation
                        .Filter = "cID like '" & Trim(txtFind.Text) & "%'"
                    Case QFixedAssetVehicle
                        .Filter = "cAssetNo like '" & Trim(txtFind.Text) & "%'"
                    Case QFixedAsset
                        .Filter = "cAssetNo like '" & Trim(txtFind.Text) & "%'"
                    Case QFixedAssetVessel
                        .Filter = "cAssetNo like '" & Trim(txtFind.Text) & "%'"
                    Case QFixedAssetEquipment
                        .Filter = "cAssetNo like '" & Trim(txtFind.Text) & "%'"
                    Case QFixedAssetOthers
                        .Filter = "cAssetNo like '" & Trim(txtFind.Text) & "%'"
                    Case QBank
                        .Filter = "cBankID like '" & Trim(txtFind.Text) & "%'"
                    Case QLocationCity
                        .Filter = "cID like '" & Trim(txtFind.Text) & "%'"
                    Case QLocationProv
                        .Filter = "cID like '" & Trim(txtFind.Text) & "%'"
                    Case QLocationArea
                        .Filter = "cID like '" & Trim(txtFind.Text) & "%'"
                    Case QMarketSegment, QProductDivision
                        .Filter = "cClassCode like '" & Trim(txtFind.Text) & "%'"
                    Case QProfitCenter
                        .Filter = "cPCCode like '" & Trim(txtFind.Text) & "%'"
                    Case QWarehouse
                        .Filter = "cWH like '" & Trim(txtFind.Text) & "%'"
                    Case QGroupings
                        .Filter = "cID like '" & Trim(txtFind.Text) & "%'"
                    Case QLead
                        .Filter = "cLeadID like '" & Trim(txtFind.Text) & "%'"
                    Case QProject, QRepProject
                        .Filter = "cProjNo like '" & Trim(txtFind.Text) & "%'"
                    Case QDivision
                        .Filter = "cDivisionID like '" & Trim(txtFind.Text) & "%'"
                    Case QServices
                        .Filter = "cServiceID like '" & Trim(txtFind.Text) & "%'"
                    Case QServiceType
                        .Filter = "cServiceTypeID like '" & Trim(txtFind.Text) & "%'"
                    Case QServiceType2
                        .Filter = "cServiceTypeID like '" & Trim(txtFind.Text) & "%'"
                    Case QWRR
                        .Filter = "cWRRNo like '" & Trim(txtFind.Text) & "%'"
                    Case QSummaryExpense
                        .Filter = "cTranNo like '" & Trim(txtFind.Text) & "%'"
                    Case QReceipt
                        .Filter = "cTranNo like '" & Trim(txtFind.Text) & "%'"
                    Case QOR_AR
                        .Filter = "cTranNo like '" & Trim(txtFind.Text) & "%'"
                End Select
            Else
                Select Case zType
                    Case QCustomer, QSalesman, QSalesmanDiv, QSupplier, QSupplierNT, QOCustomer, QCustomerExp, QEMPSUP_NT
                        .Filter = "cCode <> ''"
                    Case QBusinessType
                        .Filter = "cDesc <> ''"
                    Case QBilledTo
                        .Filter = "cNameBilled <> ''"
                    Case QSecDepartment
                        .Filter = "DeptID <> ''"
                    Case QSupply
                        .Filter = "cSupplyNo <> ''"
                    Case QPettyCashAd
                        .Filter = "cTranNo <> ''"
                    Case QTripTicket
                        .Filter = "cTranNo <> ''"
                    Case QTrucker
                        .Filter = "cTruckerCode <> ''"
                    Case QTruckerRoute
                        .Filter = "cRouteCode <> ''"
                    Case QTradeShow
                        .Filter = "cCode <> ''"
                    Case QDistributor
                        .Filter = "cCode <> ''"
                    Case QAccounts
                        .Filter = "cAcctNo <> ''"
                    Case QProduct
                        .Filter = "cItemNo <> ''"
                    Case QPriceMatrix
                        .Filter = "cPMID <> ''"
                    Case QEmployee
                        .Filter = "cEmpCode <> ''"
                    Case QEmployeeDriver
                        .Filter = "cEmpCode <> ''"
                    Case QLocation
                        .Filter = "cID <> ''"
                    Case QFixedAssetVehicle
                        .Filter = "cAssetNo <> ''"
                    Case QFixedAsset
                        .Filter = "cAssetNo <> ''"
                    Case QFixedAssetVessel
                        .Filter = "cAssetNo <> ''"
                    Case QFixedAssetEquipment
                        .Filter = "cAssetNo <> ''"
                    Case QFixedAssetOthers
                        .Filter = "cAssetNo <> ''"
                    Case QBank
                        .Filter = "cBankID <> ''"
                    Case QLocationCity
                        .Filter = "cID <> ''"
                    Case QLocationProv
                        .Filter = "cID <> ''"
                    Case QLocationArea
                        .Filter = "cID <> ''"
                    Case QMarketSegment, QProductDivision
                        .Filter = "cClassCode <> ''"
                    Case QProfitCenter
                        .Filter = "cPCCode <> ''"
                    Case QWarehouse
                        .Filter = "cWH <> ''"
                    Case QGroupings
                        .Filter = "cID <> ''"
                    Case QLead
                        .Filter = "cLeadID <> ''"
                    Case QProject, QRepProject
                        .Filter = "cProjNo <> ''"
                    Case QDivision
                        .Filter = "cDivisionID <> ''"
                    Case QServices
                        .Filter = "cServiceID <> ''"
                    Case QServiceType
                        .Filter = "cServiceTypeID <> ''"
                    Case QServiceType2
                        .Filter = "cServiceTypeID <> ''"
                    Case QWRR
                        .Filter = "cWRRNo <> ''"
                    Case QSummaryExpense
                        .Filter = "cTranNo <> ''"
                    Case QInvoice
                        .Filter = "cInvNo <> ''"
                    Case QReceipt
                        .Filter = "cTranNo <> ''"
                    Case QOR_AR
                        .Filter = "cTranNo <> ''"
                End Select
            End If
        Else
            If lReset = True Then
                Select Case zType
                    Case QCustomer, QSalesman, QSalesmanDiv, QSupplier, QSupplierNT, QOCustomer, QWarehouse, QCustomerExp, QEMPSUP_NT
                        .Filter = "cName like '" & Trim(txtFind.Text) & "%'"
                    Case QBusinessType
                        .Filter = "cDesc like '" & Trim(txtFind.Text) & "%'"
                    Case QBilledTo
                        .Filter = "cNameBilled like '" & Trim(txtFind.Text) & "%'"
                    Case QPettyCashAd
                        .Filter = "cRequestedBy like '" & Trim(txtFind.Text) & "%'"
                    Case QSupply
                        .Filter = "cDesc like '" & Trim(txtFind.Text) & "%'"
                    Case QTripTicket
                        .Filter = "cVanNo like '" & Trim(txtFind.Text) & "%'"
                    Case QTrucker
                        .Filter = "cName like '" & Trim(txtFind.Text) & "%'"
                    Case QSecDepartment
                        .Filter = "Description like '" & Trim(txtFind.Text) & "%'"
                    Case QTruckerRoute
                        .Filter = "cOrigin like '" & Trim(txtFind.Text) & "%'"
                    Case QTradeShow
                        .Filter = "cName like '" & Trim(txtFind.Text) & "%'"
                    Case QDistributor
                        .Filter = "cName like '" & Trim(txtFind.Text) & "%'"
                    Case QAccounts
                        .Filter = "cTitle like '" & Trim(txtFind.Text) & "%'"
                    Case QProduct
                        .Filter = "cDesc like '" & Trim(txtFind.Text) & "%'"
                    Case QPriceMatrix
                        .Filter = "cDesc like '" & Trim(txtFind.Text) & "%'"
                    Case QEmployee
                        .Filter = "cEmpName like '" & Trim(txtFind.Text) & "%'"
                    Case QEmployeeDriver
                        .Filter = "cEmpName like '" & Trim(txtFind.Text) & "%'"
                    Case QLocation
                        .Filter = "cLocation like '" & Trim(txtFind.Text) & "%'"
                    Case QFixedAssetVehicle
                        .Filter = "cDesc like '" & Trim(txtFind.Text) & "%'"
                    Case QFixedAsset
                        .Filter = "cDesc like '" & Trim(txtFind.Text) & "%'"
                    Case QFixedAssetVessel
                        .Filter = "cDesc like '" & Trim(txtFind.Text) & "%'"
                    Case QFixedAssetEquipment
                        .Filter = "cDesc like '" & Trim(txtFind.Text) & "%'"
                    Case QFixedAssetOthers
                        .Filter = "cDesc like '" & Trim(txtFind.Text) & "%'"
                    Case QBank
                        .Filter = "cBankName like '" & Trim(txtFind.Text) & "%'"
                    Case QLocationCity
                        .Filter = "cLocation like '" & Trim(txtFind.Text) & "%'"
                    Case QLocationProv
                        .Filter = "cLocation like '" & Trim(txtFind.Text) & "%'"
                    Case QLocationArea
                        .Filter = "cLocation like '" & Trim(txtFind.Text) & "%'"
                    Case QMarketSegment, QProductDivision, QProfitCenter, QGroupings
                        .Filter = "cDescription like '" & Trim(txtFind.Text) & "%'"
                    Case QLead
                        .Filter = "cCompany like '" & Trim(txtFind.Text) & "%'"
                    Case QProject, QRepProject
                        .Filter = "cProjTitle like '" & Trim(txtFind.Text) & "%'"
                    Case QDivision
                        .Filter = "cDivName like '" & Trim(txtFind.Text) & "%'"
                    Case QServices
                        .Filter = "cServName like '" & Trim(txtFind.Text) & "%'"
                    Case QServiceType
                        .Filter = "cServiceTypeName like '" & Trim(txtFind.Text) & "%'"
                    Case QServiceType2
                        .Filter = "cServiceTypeName like '" & Trim(txtFind.Text) & "%'"
                    Case QWRR
                        .Filter = "cName like '" & Trim(txtFind.Text) & "%'"
                    Case QSummaryExpense
                        .Filter = "cPJMNo like '" & Trim(txtFind.Text) & "%'"
'                    Case QOR_AR
'                          .Filter = "nAmount like '" & Trim(txtFind.Text) & "%'"
                End Select
            Else
                Select Case zType
                    Case QCustomer, QSalesman, QSalesmanDiv, QSupplier, QSupplierNT, QOCustomer, QCustomerExp, QEMPSUP_NT
                        .Filter = "cName <> ''"
                    Case QBusinessType
                        .Filter = "cDesc <> ''"
                    Case QSecDepartment
                        .Filter = "Description <> ''"
                    Case QSupply
                        .Filter = "cDesc <> ''"
                    Case QPettyCashAd
                        .Filter = "cRequestedBy <> ''"
                    Case QTripTicket
                        .Filter = "cVanNo <> ''"
                    Case QTrucker
                        .Filter = "cName <> ''"
                    Case QTruckerRoute
                        .Filter = "cOrigin <> ''"
                    Case QTradeShow
                        .Filter = "cName <> ''"
                    Case QDistributor
                        .Filter = "cName <> ''"
                    Case QAccounts
                        .Filter = "cTitle <> ''"
                    Case QProduct
                        .Filter = "cDesc <> ''"
                    Case QPriceMatrix
                        .Filter = "cDesc <> ''"
                    Case QEmployee
                        .Filter = "cEmpName <> ''"
                    Case QEmployeeDriver
                        .Filter = "cEmpName <> ''"
                    Case QLocation
                        .Filter = "cLocation <> ''"
                    Case QFixedAssetVehicle
                        .Filter = "cDesc <> ''"
                    Case QFixedAsset
                        .Filter = "cDesc <> ''"
                    Case QFixedAssetVessel
                        .Filter = "cDesc <> ''"
                    Case QFixedAssetEquipment
                        .Filter = "cDesc <> ''"
                    Case QFixedAssetOthers
                        .Filter = "cDesc <> ''"
                    Case QBank
                        .Filter = "cBankName <> ''"
                    Case QLocationCity
                        .Filter = "cLocation <> ''"
                    Case QLocationProv
                        .Filter = "cLocation <> ''"
                    Case QLocationArea
                        .Filter = "cLocation <> ''"
                    Case QMarketSegment, QProductDivision, QProfitCenter, QGroupings
                        .Filter = "cDescription <> ''"
                    Case QLead
                        .Filter = "cCompany <> ''"
                    Case QProject, QRepProject
                        .Filter = "cProjTitle <> ''"
                    Case QDivision
                        .Filter = "cDivName <> ''"
                    Case QServices
                        .Filter = "cServName <> ''"
                    Case QServiceType
                        .Filter = "cServiceTypeName <> ''"
                    Case QServiceType2
                        .Filter = "cServiceTypeName <> ''"
                    Case QWRR
                        .Filter = "cName <> ''"
                    Case QSummaryExpense
                        .Filter = "cPJMNo <> ''"
                End Select
            End If
        End If
    End With
End Sub

Public Sub ShowForm1()
On Error GoTo TheSource

    dtgList.ClearFields
    If rsPickList.State = adStateOpen Then rsPickList.Close
    DoEvents
    oConnection.OpenNewConnection connList
    'MousePointer = vbHourglass
    If Not lModal Then FormWaitShow "Loading list . . ."
    
    Select Case zType
        Case QCustomer
            Caption = "Customer List"
            If sColumnVariable = "Code" Then
                LoadColumn "cCode", "cName", "CLIENT_CUSTOMER", 1, "Customer ID", "Customer Name", IIf(sTypeVariable = "", "", "and cCustomerType = '" & Trim(sTypeVariable) & "' "), "cCode"
            Else
                LoadColumn "cName", "cCode", "CLIENT_CUSTOMER", 2, "Customer Name", "Customer ID", IIf(sTypeVariable = "", "", "and cCustomerType = '" & Trim(sTypeVariable) & "' "), "cName"
            End If
        Case QBusinessType
            Caption = "Business Type List"
            If sColumnVariable = "Code" Then
                LoadColumn "nIdentity", "cDesc", "BusinessType", 1, "ID", "Description", IIf(sTypeVariable = "", "", "and cCustomerType = '" & Trim(sTypeVariable) & "' "), "nIdentity"
            Else
                LoadColumn "cDesc", "nIdentity", "BusinessType", 2, "Description", "ID", IIf(sTypeVariable = "", "", "and cCustomerType = '" & Trim(sTypeVariable) & "' "), "cDesc"
            End If
            
        Case QCustomerExp
            Caption = "Customer List"
            If sColumnVariable = "Code" Then
                LoadColumn "cCode", "cName", "CLIENT_EXPORT", 1, "Customer ID", "Customer Name", IIf(sTypeVariable = "", "", "and cCustomerType = '" & Trim(sTypeVariable) & "' "), "cCode"
            Else
                LoadColumn "cName", "cCode", "CLIENT_EXPORT", 2, "Customer Name", "Customer ID", IIf(sTypeVariable = "", "", "and cCustomerType = '" & Trim(sTypeVariable) & "' "), "cName"
            End If
        Case QSecDepartment
            Caption = "Department List"
            If sColumnVariable = "Code" Then
                LoadColumn "DeptID", "Description", "SEC_DEPARTMENT", 1, "Department ID", "Department Name", " ", "DeptID", True
            Else
                LoadColumn "Description", "DeptID", "SEC_DEPARTMENT", 2, "Department Name", "Department ID", " ", "Description", True
            End If
        Case QSalesman
            Caption = "Project Manager List"
            If sColumnVariable = "Code" Then
                LoadColumn "cCode", "cName", "SALESMAN", 1, "PJM ID", "PJM Name", " ", "cCode"
            Else
                LoadColumn "cName", "cCode", "SALESMAN", 2, "PJM Name", "PJM ID", " ", "cName"
            End If
        Case QSalesmanDiv
            Caption = "Project Manager List"
            If sColumnVariable = "Code" Then
                LoadColumn "cCode", "cName", "SALESMAN", 1, "PJM ID", "PJM Name", "and cDivisionID = '" & Trim(RepName) & "'", "cCode"
            Else
                LoadColumn "cName", "cCode", "SALESMAN", 2, "PJM Name", "PJM ID", "and cDivisionID = '" & Trim(RepName) & "'", "cName"
            End If
        Case QTripTicket
            Caption = "Trip Ticket List"
            If sColumnVariable = "Code" Then
                LoadColumn "cTranNo", "cVanNo", "TRIPTICKET", 1, "Trip Ticket", "Van", " ", "cTranNo"
            Else
                LoadColumn "cVanNo", "cTranNo", "TRIPTICKET", 2, "Van", "Trip Ticket", " ", "cVanNo"
            End If
        Case QPettyCashAd
            Caption = "Petty Cash Advance List"
            If sColumnVariable = "Code" Then
                LoadColumn "cTranNo", "cRequestedBy", "PETTY", 1, "Petty Cash Advance no.", "Requested by", " ", "cTranNo"
            Else
                LoadColumn "cRequestedBy", "cTranNo", "PETTY", 2, "Requested by", "Petty Cash Advance no.", " ", "cRequestedBy"
            End If
        Case QTrucker
            Caption = "Trucker List"
            If sColumnVariable = "Code" Then
                LoadColumn "cTruckerCode", "cName", "TRUCKER", 1, "Trucker ID", "Trucker Name", " ", "cTruckerCode"
            Else
                LoadColumn "cName", "cTruckerCode", "TRUCKER", 2, "Trucker Name", "Trucker ID", " ", "cName"
            End If
        Case QTruckerRoute
            Caption = "Trucker Route List"
            If sColumnVariable = "Code" Then
                LoadColumn "cRouteCode", "cOrigin", "TRUCKER_ROUTE", 1, "Trucker Route ID", "Trucker Route Name", " ", "cRouteCode"
            Else
                LoadColumn "cOrigin", "cRouteCode", "TRUCKER_ROUTE", 2, "Trucker Route Name", "Trucker Route ID", " ", "cOrigin"
            End If
        Case QTradeShow
            Caption = "Trade Show List"
            If sColumnVariable = "Code" Then
                LoadColumn "cCode", "cName", "CLIENT_CUSTOMER", 1, "Trade Show ID", "Trade Show Name", "and cCustomerType = 'Trade Show'", "cCode"
            Else
                LoadColumn "cName", "cCode", "CLIENT_CUSTOMER", 2, "Trade Show Name", "Trade Show ID", "and cCustomerType = 'Trade Show'", "cName"
            End If
        Case QDistributor
            Caption = "Distributor List"
            If sColumnVariable = "Code" Then
                LoadColumn "cCode", "cName", "CLIENT_CUSTOMER", 1, "Distributor ID", "Distributor Name", "and cCustomerType = 'Distributor'", "cCode"
            Else
                LoadColumn "cName", "cCode", "CLIENT_CUSTOMER", 2, "Distributor Name", "Distributor ID", "and cCustomerType = 'Distributor'", "cName"
            End If
        Case QAccounts
            Caption = "Accounts List"
            If sColumnVariable = "Code" Then
                LoadColumn "cAcctNo", "cTitle", "ACCOUNT", 1, "Account No.", "Account Title", " ", "cAcctNo"
            Else
                LoadColumn "cTitle", "cAcctNo", "ACCOUNT", 2, "Account Title", "Account No.", " ", "cTitle"
            End If
        Case QSupplier
            Caption = "Supplier List"
            If sColumnVariable = "Code" Then
                LoadColumn "cCode", "cName", "CLIENT_SUPPLIER", 1, "Supplier ID", "Supplier Name", " ", "cCode"
            Else
                LoadColumn "cName", "cCode", "CLIENT_SUPPLIER", 2, "Supplier Name", "Supplier ID", " ", "cName"
            End If
        Case QBilledTo
        
            Caption = "List of Billed Name"
            If sColumnVariable = "Code" Then
                LoadColumn "cCode", "cNameBilled", "CLIENT_CUSTOMER_BilledTo", 1, "Code", "Name", " ", "cCode"
            Else
                LoadColumn "cNameBilled", "cCode", "CLIENT_CUSTOMER_BilledTo", 2, "Name", "Code", " AND  cCode = '" & cCodeCustomer & "'", "cNameBilled"
            End If
            
        Case QEMPSUP_NT
            Caption = "Supplier List (Non-Trade) And Employee List"
            If sColumnVariable = "Code" Then
                LoadColumn2 "cCode", "cName", "rsp_supplierNT_employee", 1, "Code", "Name", "Code"
            Else
                LoadColumn2 "cName", "cCode", "rsp_supplierNT_employee", 2, "Name", "Code", "Name"
            End If
        
        
        Case QSupplierNT
            Caption = "Supplier List (Non-Trade)"
            If sColumnVariable = "Code" Then
                LoadColumn "cCode", "cName", "CLIENT_SUPPLIER_NT", 1, "Supplier ID", "Supplier Name", " ", "cCode"
            Else
                LoadColumn "cName", "cCode", "CLIENT_SUPPLIER_NT", 2, "Supplier Name", "Supplier ID", " ", "cName"
            End If
        Case QSupply
            Caption = "Supply List"
            If sColumnVariable = "Code" Then
                LoadColumn "cSupplyNo", "cDesc", "SUPPLY", 1, "Supply ID", "Supply Name", " ", "cSupplyNo"
            Else
                LoadColumn "cDesc", "cSupplyNo", "SUPPLY", 2, "Supply Name", "Supply ID", " ", "cDesc"
            End If
        Case QProduct
            Caption = "Product List"
            If sColumnVariable = "Code" Then
                LoadColumn "cItemNo", "cDesc", "ITEM", 1, "Product ID", "Product Description", " ", "cItemNo"
            Else
                LoadColumn "cDesc", "cItemNo", "ITEM", 2, "Product Description", "Product ID", " ", "cDesc"
            End If
        Case QPriceMatrix
            Caption = "Price Matrix List"
            If sColumnVariable = "Code" Then
                LoadColumn "cPMID", "cDesc", "PM", 1, "Price Matrix ID", "Price Matrix Description", " ", "cPMID"
            Else
                LoadColumn "cDesc", "cPMID", "PM", 2, "Price Matrix Description", "Price Matrix ID", " ", "cDesc"
            End If
        Case QEmployee
            Caption = "Employee List"
            If sColumnVariable = "Code" Then
                LoadColumn "cEmpCode", "cEmpName", "EMPLOYEE", 1, "Employee ID", "Employee Name", " ", "cEmpCode"
            Else
                LoadColumn "cEmpName", "cEmpCode", "EMPLOYEE", 2, "Employee Name", "Employee ID", " ", "cEmpName"
            End If
        Case QEmployeeDriver
            Caption = "Driver List"
            If sColumnVariable = "Code" Then
                LoadColumn "cEmpCode", "cEmpName", "EMPLOYEE", 1, "Employee ID", "Employee Name", "and cPosition = 'Driver'", "cEmpCode"
            Else
                LoadColumn "cEmpName", "cEmpCode", "EMPLOYEE", 2, "Employee Name", "Employee ID", "and cPosition = 'Driver'", "cEmpName"
            End If
        Case QOCustomer
            Caption = "Other Customer List"
            If sColumnVariable = "Code" Then
                LoadColumn "cCode", "cName", "CLIENT_OTHERS", 1, "Customer ID", "Customer Name", " ", "cCode"
            Else
                LoadColumn "cName", "cCode", "CLIENT_OTHERS", 2, "Customer Name", "Customer ID", " ", "cName"
            End If
        Case QBank
            Caption = "Bank List"
            If sColumnVariable = "Code" Then
                LoadColumn "cBankID", "cBankName", "BANK", 1, "Bank ID", "Bank Name", " ", "cBankID"
            Else
                LoadColumn "cBankName", "cBankID", "BANK", 2, "Bank Name", "Bank ID", " ", "cBankName"
            End If
        Case QLocation
            Caption = "Location List"
            If sColumnVariable = "Code" Then
                LoadColumn "cID", "cLocation", "LOCATION", 1, "Location ID", "Location Name", " ", "cID"
            Else
                LoadColumn "cLocation", "cID", "LOCATION", 2, "Location Name", "Location ID", " ", "cLocation"
            End If
        Case QFixedAssetVehicle
            Caption = "Fixed Asset List"
            If sColumnVariable = "Code" Then
                LoadColumn "cAssetNo", "cDesc", "ASSET", 1, "Fixed Asset No.", "Fixed Asset Description", "and cType = 'Vehicle'", "cAssetNo"
            Else
                LoadColumn "cDesc", "cAssetNo", "ASSET", 2, "Fixed Asset Description", "Fixed Asset No.", "and cType = 'Vehicle'", "cDesc"
            End If
        Case QFixedAsset
            Caption = "Fixed Asset List"
            If sColumnVariable = "Code" Then
                LoadColumn "cAssetNo", "cDesc", "ASSET", 1, "Fixed Asset No.", "Fixed Asset Description", " ", "cAssetNo"
            Else
                LoadColumn "cDesc", "cAssetNo", "ASSET", 2, "Fixed Asset Description", "Fixed Asset No.", " ", "cDesc"
            End If
        Case QFixedAssetVessel
            Caption = "Fixed Asset List"
            If sColumnVariable = "Code" Then
                LoadColumn "cAssetNo", "cDesc", "ASSET", 1, "Fixed Asset No.", "Fixed Asset Description", " and cType = 'Vessel'", "cAssetNo"
            Else
                LoadColumn "cDesc", "cAssetNo", "ASSET", 2, "Fixed Asset Description", "Fixed Asset No.", " and cType = 'Vessel'", "cDesc"
            End If
        Case QFixedAssetEquipment
            Caption = "Fixed Asset List"
            If sColumnVariable = "Code" Then
                LoadColumn "cAssetNo", "cDesc", "ASSET", 1, "Fixed Asset No.", "Fixed Asset Description", " and cType = 'Equipment'", "cAssetNo"
            Else
                LoadColumn "cDesc", "cAssetNo", "ASSET", 2, "Fixed Asset Description", "Fixed Asset No.", " and cType = 'Equipment'", "cDesc"
            End If
        Case QFixedAssetOthers
            Caption = "Fixed Asset List"
            If sColumnVariable = "Code" Then
                LoadColumn "cAssetNo", "cDesc", "ASSET", 1, "Fixed Asset No.", "Fixed Asset Description", " and cType = 'Others'", "cAssetNo"
            Else
                LoadColumn "cDesc", "cAssetNo", "ASSET", 2, "Fixed Asset Description", "Fixed Asset No.", " and cType = 'Others'", "cDesc"
            End If
        Case QLocationCity
            Caption = "Location List"
            If sColumnVariable = "Code" Then
                LoadColumn "cID", "cLocation", "LOCATION", 1, "Location ID", "Location Name", " and cType = 'City'", "cID"
            Else
                LoadColumn "cLocation", "cID", "LOCATION", 2, "Location Name", "Location ID", " and cType = 'City'", "cLocation"
            End If
        Case QLocationProv
            Caption = "Location List"
            If sColumnVariable = "Code" Then
                LoadColumn "cID", "cLocation", "LOCATION", 1, "Location ID", "Location Name", " and cType = 'Region'", "cID"
            Else
                LoadColumn "cLocation", "cID", "LOCATION", 2, "Location Name", "Location ID", " and cType = 'Region'", "cLocation"
            End If
        Case QLocationArea
            Caption = "Location List"
            If sColumnVariable = "Code" Then
                LoadColumn "cID", "cLocation", "LOCATION", 1, "Location ID", "Location Name", " and cType = 'Area'", "cID"
            Else
                LoadColumn "cLocation", "cID", "LOCATION", 2, "Location Name", "Location ID", " and cType = 'Area'", "cLocation"
            End If
        Case QMarketSegment
            Caption = "Market Segment List"
            If sColumnVariable = "Code" Then
                LoadColumn "cClassCode", "cDescription", "CLASSIFICATION", 1, "MS Code", "Description", " and cType = 'MS'", "cClassCode"
            Else
                LoadColumn "cDescription", "cClassCode", "CLASSIFICATION", 2, "Description", "MS Code", " and cType = 'MS'", "cDescription"
            End If
        Case QProductDivision
            Caption = "Product Division List"
            If sColumnVariable = "Code" Then
                LoadColumn "cClassCode", "cDescription", "CLASSIFICATION", 1, "MS Code", "Description", " and cType = 'PD'", "cClassCode"
            Else
                LoadColumn "cDescription", "cClassCode", "CLASSIFICATION", 2, "Description", "MS Code", " and cType = 'PD'", "cDescription"
            End If
        Case QProfitCenter
            Caption = "Profit Center List"
            If sColumnVariable = "Code" Then
                LoadColumn "cPCCode", "cDescription", "PROFITCENTER", 1, "PC Code", "Description", " ", "cPCCode"
            Else
                LoadColumn "cDescription", "cPCCode", "PROFITCENTER", 2, "Description", "PC Code", " ", "cDescription"
            End If
        Case QWarehouse
            Caption = "Warehouse List"
            If sColumnVariable = "Code" Then
                LoadColumn "cWH", "cName", "WHSE", 1, "WHSE Code", "Description", " ", "cWH"
            Else
                LoadColumn "cName", "cWH", "WHSE", 2, "Description", "WHSE Code", " ", "cName"
            End If
        Case QGroupings
            Caption = "Groupings List"
            If sColumnVariable = "Code" Then
                LoadColumn "cID", "cDescription", "GROUPINGS", 1, "Group Code", "Description", "and cGroupNo = '" & Trim(RepName) & "' ", "cID"
            Else
                LoadColumn "cDescription", "cID", "GROUPINGS", 2, "Description", "Group Code", "and cGroupNo = '" & Trim(RepName) & "' ", "cID"
            End If
        Case QLead
            Caption = "Leads List"
            If sColumnVariable = "Code" Then
                LoadColumn "cLeadID", "cCompany", "PMS_LEAD", 1, "Lead ID", "Company Name", " ", "cLeadID"
            Else
                LoadColumn "cCompany", "cLeadID", "PMS_LEAD", 2, "Company Name", "Lead ID", " ", "cCompany"
            End If
        Case QProject
            Caption = "Project List"
            If sColumnVariable = "Code" Then
                LoadColumn "cProjNo", "cProjTitle", "PMS_PROJECT", 1, "Project ID", "Project Title", "and cStatus = 'Active' ", "cProjNo"
            Else
                LoadColumn "cProjTitle", "cProjNo", "PMS_PROJECT", 2, "Project Title", "Project ID", "and cStatus = 'Active' ", "cProjTitle"
            End If
        Case QRepProject
            Caption = "Project List"
            If sColumnVariable = "Code" Then
                LoadColumn "cProjNo", "cProjTitle", "PMS_PROJECT", 1, "Project ID", "Project Title", " ", "cProjNo"
            Else
                LoadColumn "cProjTitle", "cProjNo", "PMS_PROJECT", 2, "Project Title", "Project ID", " ", "cProjTitle"
            End If
        Case QDivision
            Caption = "Division List"
            If sColumnVariable = "Code" Then
                LoadColumn "cDivisionID", "cDivName", "PMS_DIVISION", 1, "Division ID", "Division Name", " ", "cDivisionID"
            Else
                LoadColumn "cDivName", "cDivisionID", "PMS_DIVISION", 2, "Division Name", "Division ID", " ", "cDivName"
            End If
        Case QServices
            Caption = "Strategic Business Unit"
            If sColumnVariable = "Code" Then
                LoadColumn "cServiceID", "cServName", "PMS_SERVICE", 1, "SBU Code", "Strategic Business Units (SBUs)", " ", "cServiceID"
            Else
                LoadColumn "cServName", "cServiceID", "PMS_SERVICE", 2, "Strategic Business Units (SBUs)", "SBU Code", " ", "cServName"
            End If
        Case QServiceType
            Caption = "Service Type"
            If sColumnVariable = "Code" Then
                LoadColumn "cServiceTypeID", "cServiceTypeName", "PMS_SERVICE_TYPE", 1, "Service Type ID", "Service Type Name", " AND cServiceTypeID NOT IN ('09','00') ", "cServiceTypeID"
            Else
                LoadColumn "cServiceTypeName", "cServiceTypeID", "PMS_SERVICE_TYPE", 2, "Service Type Name", "Service Type ID", " AND cServiceTypeID NOT IN ('09','00') ", "cServiceTypeName"
            End If
        Case QServiceType2
            Caption = "Service Type"
            If sColumnVariable = "Code" Then
                LoadColumn "cServiceTypeID", "cServiceTypeName", "PMS_SERVICE_TYPE", 1, "Service Type ID", "Service Type Name", "  ", "cServiceTypeID"
            Else
                LoadColumn "cServiceTypeName", "cServiceTypeID", "PMS_SERVICE_TYPE", 2, "Service Type Name", "Service Type ID", "  ", "cServiceTypeName"
            End If
        Case QWRR
            Caption = "Receiving Report"
            If sColumnVariable = "Code" Then
                LoadColumn "cWRRNo", "cName", "WRR", 1, "RR No.", "Supplier Name", " ", "cWRRNo"
            Else
                LoadColumn "cWRRNo", "cName", "WRR", 2, "RR No.", "Supplier Name", " ", "cWRRNo"
            End If
        Case QSummaryExpense
            Caption = "Receiving Report"
            If sColumnVariable = "Code" Then
                LoadColumn "cTranNo", "cPJMNo", "PMS_EXPENSE", 1, "SE No.", "Employee ID", " ", "cTranNo"
            Else
                LoadColumn "cTranNo", "cPJMNo", "PMS_EXPENSE", 2, "SE No.", "Employee ID", " ", "cTranNo"
            End If
        Case QInvoice
            Caption = "Service Invoice"
            If sColumnVariable = "Code" Then
                LoadColumn "cInvNo", "cCode", "SALES", 1, "Invoice No.", "Client", " ", "cInvNo"
            Else
                LoadColumn "cInvNo", "cCode", "SALES", 2, "Invoice No.", "Client", " ", "cInvNo"
            End If
        Case QReceipt
            Caption = "Official Receipt"
            If sColumnVariable = "Code" Then
                LoadColumn "cTranNo", "cCode", "PR", 1, "Trans. No.", "Client", " ", "cTranNo"
            Else
                LoadColumn "cTranNo", "cCode", "PR", 2, "Trans. No.", "Client", " ", "cTranNo"
            End If
            Case QOR_AR
            Caption = "Official Receipt And Acknowledgement Receipt"
            If sColumnVariable = "Code" Then
                LoadColumn2 "cTranNo", "nAmount", "rsp_PickList_OR_AR", 1, "Trans. No.", "Amount", "Code"
            Else
                LoadColumn2 "nAmount", "cTranNo", "rsp_PickList_OR_AR", 2, "Amount", "Trans. No.", "Amount"
            End If
            
    End Select
    If Not lModal Then Unload frmWait
    lModal = False
    FilterString True
    
    'MousePointer = vbDefault
    
TheSource:
    If Err.Number = 3709 Then
        Set connList = Nothing
        Unload frmWait
        'MousePointer = vbDefault
        Unload Me
    End If
End Sub

Private Sub optCode_Click()
    sColumnVariable = "Code"
    ShowForm1
    dtgList.SetFocus
End Sub

Private Sub optName_Click()
    sColumnVariable = "Desc"
    ShowForm1
    dtgList.SetFocus
End Sub

Private Sub LoadColumn(FirstArg As String, SecondArg As String, TableArg As String, _
            CodeArg As Integer, Caption1 As String, Caption2 As String, Cond As String, OrderArg As String, Optional lWOCompany As Boolean)
    
    If lWOCompany Then
        If UCase(Left(Trim(Cond), 3)) = "AND" Then Cond = Trim(Mid(Trim(Cond), 4, Len(Trim(Cond))))
        
        If Trim(Cond) = "" Then
            sSQL = "SELECT " & FirstArg & ", " & SecondArg & " FROM " & TableArg & _
                " ORDER BY " & Trim(OrderArg)
        Else
            sSQL = "SELECT " & FirstArg & ", " & SecondArg & " FROM " & TableArg & _
                " WHERE " & Trim(Cond) & _
                " ORDER BY " & Trim(OrderArg)
        End If
    Else
    sSQL = "SELECT " & FirstArg & ", " & SecondArg & " FROM " & TableArg & _
           " WHERE cCompanyID = '" & Trim(COID) & "' " & Trim(Cond) & _
           " ORDER BY " & Trim(OrderArg)
    End If
    
    rsPickList.Open sSQL, connList, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set dtgList.DataSource = rsPickList
    With rsPickList
        dtgList.Columns(0).DataField = (FirstArg)
        dtgList.Columns(1).DataField = (SecondArg)
    End With
    If CodeArg = 1 Then
        dtgList.Columns(0).Width = 1500
        dtgList.Columns(1).Width = 5300
    Else
        dtgList.Columns(0).Width = 5300
        dtgList.Columns(1).Width = 1500
    End If
    dtgList.Columns(0).Caption = Caption1
    dtgList.Columns(1).Caption = Caption2

End Sub
    
Private Sub LoadColumn2(FirstArg As String, SecondArg As String, TableArg As String, _
            Optional CodeArg As Integer, Optional Caption1 As String, Optional Caption2 As String, Optional Cond As String, Optional OrderArg As String, Optional lWOCompany As Boolean)
    
    sSQL = "EXEC " & TableArg & " '" & COID & "' ,  '" & Cond & "'"

    rsPickList.Open sSQL, connList, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set dtgList.DataSource = rsPickList
    With rsPickList
        dtgList.Columns(0).DataField = (FirstArg)
        dtgList.Columns(1).DataField = (SecondArg)
    End With
    If CodeArg = 1 Then
        dtgList.Columns(0).Width = 1500
        dtgList.Columns(1).Width = 5300
    Else
        dtgList.Columns(0).Width = 5300
        dtgList.Columns(1).Width = 1500
    End If
    

    dtgList.Columns(0).Caption = Caption1
    dtgList.Columns(1).Caption = Caption2

End Sub

