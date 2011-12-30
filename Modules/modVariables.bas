Attribute VB_Name = "modVariables"
Option Explicit

'Variable for Licensing ComUnion
Public Type eSerial
    lSeparator1 As String * 5
    CompanyID As String * 20
    lSeparator2 As String * 5
    HSerial As String * 15
    lSeparator3 As String * 5
    ActivationKey As String * 25
    SerialKey As String * 25
    lSeparator4 As String * 5
End Type

Public lRegistered As Boolean
Public dTranDate As Date

Public Const HandCursor = 32649&
Public Declare Function SetCursor Lib "USER32" (ByVal hCursor As Long) As Long
Public Declare Function LoadCursor Lib "USER32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Choice As String
Public lTasksFlag As String

Declare Function vbEncodeIPtr Lib "p2smon.dll" (x As Object) As String
Declare Function CreateReportOnRuntimeDS Lib "p2smon.dll" (x As Object, ByVal reportPath$, ByVal fieldDefFilePath$, ByVal bOverWriteExistingFiles%, ByVal bLaunchDesigner%) As Integer
Declare Function OpenReportOnRuntimeDS Lib "p2smon.dll" (x As Object, ByVal reportPath$, ByVal fieldDefFilePath$, ByVal bOverWriteExistingFiles%, ByVal bLaunchDesigner%) As Integer
Declare Function CreateFieldDefFile Lib "p2smon.dll" (x As Object, ByVal fieldDefFilePath$, ByVal bOverWriteExistingFiles%) As Integer

'Global declarations
Global glbCN As ADODB.Connection
Global glbCNShape As ADODB.Connection
Global glbSSQL As String

'iTG Module Constants
Global Const CCDEPOSIT = "CCDEPOSIT"  'CREDIT & COLLECTION DEPOSIT SLIP
Global Const CCOR = "CCOR" 'CREDIT & COLLECTION RECEIPT
Global Const CCPULLED = "CCPULLED" 'CREDIT & COLLECTION PULLED-OUT CHECK
Global Const CCREPLACE = "CCREPLACE" 'CREDIT & COLLECTION CHECK REPLACEMENT
Global Const CCRETURN = "CCRETURN" 'CREDIT & COLLECTION RETURN/BOUNCED CHECK
Global Const DOWNLOAD = "DOWNLOAD" 'DOWNLOADING
Global Const IPDR = "IPDR" 'INVENTORY & PRODUCTION DELIVERY RECEIPT
Global Const IPINVADJ = "IPINVADJ" 'INVENTORY & PRODUCTION ADJUSTMENTS
Global Const IPINVTRAN = "IPINVTRAN" 'INVENTORY & PRODUCTION TRANSFER
Global Const IPPO = "IPPO"  'INVENTORY & PRODUCTION ORDER
Global Const IPWRR = "IPWRR" 'INVENTORY & PRODUCTION RECEIVING REPORT
Global Const LDTRIP = "LDTRIP" 'LOGICTICS & DELIVERY TRIP TICKET
Global Const MACCOUNT = "MACCOUNT" 'MAINTENANCE ACCOUNT
Global Const MBANK = "MBANK" 'MAINTENANCE BANK
Global Const MCHECKS = "MCHECKS" 'MAINTENANCE CHECKS ON HAND
Global Const MCLIENT = "MCLIENT" 'MAINTENANCE CLIENT
Global Const MCLIENTEXP = "MCLIENTEXP" 'MAINTENANCE CLIENT EXPORT
Global Const MCOMPANY = "MCOMPANY" 'MAINTENANCE COMPANY
Global Const MCURRENCY = "MCURRENCY" 'MAINTENANCE CURRENCY
Global Const MEWT = "MEWT" 'MAINTENANCE EWT
Global Const MEVAT = "MEVAT" 'MAINTENANCE EVAT
Global Const MGROUP = "MGROUP" 'MAINTENANCE PRODUCT GROUPINGS
Global Const MMS = "MMS" 'MAINTENANCE MARKET SEGMENT
Global Const MPD = "MPD" 'MAINTENANCE PRODUCT DIVISION
Global Const MPRICE = "MPRICE" 'MAINTENANCE PRICE LIST
Global Const MPRODUCT = "MPRODUCT" 'MAINTENANCE PRODUCT
Global Const MSALESMAN = "MSALESMAN" 'MAINTENANCE SALESMAN
Global Const MWHSE = "MWHSE" 'MAINTENANCE WAREHOUSE
Global Const MFASSETS = "MFASSETS" 'MAINTENANCE FIXED ASSETS
Global Const MPROFIT = "MPROFIT" 'MAINTENANCE PROFIT CENTER
Global Const MCOMCRI = "MCOMCRI" 'MAINTENANCE COMMISSION CRITERIA
Global Const MCOMSETUP = "MCOMSETUP" 'MAINTENANCE COMMISSION SETUP
Global Const OBCMDM = "OBCMDM" 'ORDER & BILLING ADJUSTMENTS
Global Const OBCMITEM = "OBCMITEM" 'ORDER & BILLING CM ADJUSTMENTS PER PRODUCT
Global Const OBPROJECT = "OBPROJECT" 'ORDER & BILLING PROJECT SETUP
Global Const OBSI = "OBSI" 'ORDER & BILLING SALES INVOICE
Global Const OBSIPROJ = "OBSIPROJ" 'ORDER & BILLING SALES INVOICE PROJECT
Global Const OBSO = "OBSO" 'ORDER & BILLING SALES ORDER
Global Const PPCMDM = "PPCMDM" 'PURCHASING & PAYABLES ADJUSTMENTS
Global Const PPPO = "PPPO"  'PURCHASING & PAYABLES PURCHASE ORDER
Global Const PPPR = "PPPR" 'PURCHASING & PAYABLES PURCHASE REQUISITION
Global Const PPRFP = "PPRFP" 'PURCHASING & PAYABLES REQUEST FOR PAYMENT
Global Const REPORT = "REPORT" 'REPORTS
Global Const SACCESS = "SACCESS" 'SECURITY USER ACCESS LEVEL
Global Const SCACCESS = "SCACCESS" 'SECURITY COMPANY ACCESS
Global Const SDEPT = "SDEPT" 'SECURITY DEPARTMENT
Global Const sModule = "SMODULE" 'SECURITY MODULE
Global Const SROLE = "SROLE" 'SECURITY USER ROLE
Global Const SUSER = "SUSER" 'SECURITY USER PROFILE
Global Const TSETUP = "TSETUP" 'TOOLS SYSTEM SETUP
Global Const TSPARAM = "TSPARAM" 'TOOLS SYSTEM PARAMETER
Global Const TUPARAM = "TUPARAM" 'TOOLS USER PARAMETER
'--FOR GENERAL LEDGER ONLY MODULES--
Global Const GLJOURNAL = "GLJOURNAL" 'GENERAL LEDGER JOURNAL ENTRY
Global Const GLINTVOU = "GLINTVOU" 'GENERAL LEDGER INTERNAL VOUCHER
Global Const GLARCPT = "GLARCPT" 'GENERAL LEDGER ACKNOWLEDGMENT RECEIPT
Global Const GLDM = "GLDM" 'GENERAL LEDGER DEBIT MEMO
Global Const GLCM = "GLCM" 'GENERAL LEDGER CREDIT MEMO
Global Const PPVOUCHER = "PPVOUCHER" 'PURCHASING & PAYABLES VOUCHER
Global Const PPCHKISS = "PPCHKISS" 'PURCHASING & PAYABLES CHECK ISSUANCE
Global Const PPCHKCLR = "PPCHKCLR" 'PURCHASING & PAYABLES CHECK CLEARING
Global Const PPCHKBBL = "PPCHKBBL" 'PURCHASING & PAYABLES BANK CHECK BOOKLET LIST
Global Const PPBANKCL = "PPBANKCL" 'PURCHASING & PAYABLES BANK COMPANY LIST
Global Const TGLINT = "TGLINT" 'TOOLS GENERAL LEDGER INTERFACE

'Public declarations
Public FrmName As Form
Public rs As New Recordset
Public tmpRS As New Recordset
Public objBox As TextBox
Public objCtl As Control
Public i As Integer
Public itmX
Public vBookMark As Variant
Public RepName As String
Public cString As String
Public lBoolean As Boolean
Public lCloseWindow As Boolean
Public lModal As Boolean
Public lSaving As Boolean
Public lPickListActive As Boolean
Public nCount As Long
Public sID As String
Public sIDLprint As String
Public sFilterString As String
Public UserFullName As String
Public FrmCaption As String
Public ClosingREP As String
Public UserRole As String

Global Const sWOGenAcct = "" '" AND cType <> 'General' "

'Change transaction number variable
Public sNewTranNo As String

'Variables for connection
Public sServer As String
Public sDBname As String
Public sDBDriver As String
Public sUserName As String
Public sDBPassword As String
Public sUserID As String
Public sUserDept As String
Public sDivision As String

Public SecUserID As String
Public SecUserRole As String
Public SecUserName As String

'Variable for Terminal ID
Public SecTerminalID As Integer

'Variable for Company ID
Public COID As String
Public COName As String
Public sBankCode As String
Public sBankComp As String
Public sUnitId As String

Public sClient As String
Public cSQLFind As String

'Variable for Downloading
Public sDownload As String

'Variable for Temporary container
Public vStrContainer1 As String
Public vStrContainer2 As String
Public vStrContainer3 As String
Public vStrContainer4 As String
Public vStrContainer5 As String
Public vStrContainer6 As String
Public vNumContainer1 As Double
Public vNumContainer2 As Double

'Pull-out slip option variable
Public lPullOutSlip As Boolean

'Purchase order slip option variable
Public lNoCanvass As Boolean

'*********************************
'System Report Global Variables
'By : Jon Fonacier
'*********************************
Public cReport As String
Public cReportTitle As String
Public cCompany As String
Public cAddress1 As String
Public cAddress2 As String
Public cCriteria1 As String
Public cCriteria2 As String
Public nVarInterval As Integer
Public cFilter As String
Public cFilter1 As String

Public gReportDataFrom As String
Public gReportDataTo As String
Public gReportFilterBy As String

Public gReportDateFrom As String
Public gReportDateTo As String
Public cReportName As String
Public cStorProcName As String
Public cFilePreview As String
Public cTypeReport As String

Public cID As String
Public cName As String
Public nLeadTime As Integer
Public nBuffer As Integer
Public nDaysFrom As Integer
Public nDaysTo As Integer
Public cWeek1From As String
Public cWeek1To As String
Public cWeek2From As String
Public cWeek2To As String
Public cWeek3From As String
Public cWeek3To As String
Public cWeek4From As String
Public cWeek4To As String
Public cWeek5From As String
Public cWeek5To As String

'*********************************
'Ad Hoc Report Form Variable
'*********************************
Public lAdHoc As Boolean
Public cCommand As String
Public nFileLen As Integer
Public CrystalApplication As CRAXDRT.Application
Public CrystalReport As CRAXDRT.REPORT
Public CrystalDatabase As CRAXDRT.Database
Public CrystalTables As CRAXDRT.DatabaseTables
Public CrystalTable As CRAXDRT.DatabaseTable
Public AdoRS As ADODB.Recordset
'*********************************
'Receipt Printing Report
'*********************************
Public cModule As String
Public cModulePrint As String
Public cSQL As String

Public post As Integer
Public modulecode As String
Public cFrom As String
Public cTo As String


'added
'Public mFrmName As Form


'Approval Action variables
Enum eGAction
    G_Cancel
    G_OK
End Enum
Public gAction As eGAction

'logicglassupdate
Public cStrTypeHdr As String
Public cStrType As String

'Variables For Auto Email
Public TimeCollection() As String
Public IsReportSend() As Boolean
Public PassTime As String

'Variable for Autonumber
Public CurrentYear As String
Public CurrentMonth As String
Public nCurrentCounter As Integer
'New Code - Autonumber
Public AutonumFormat As String
Public nNumStart As String
Public nNumLen As Integer
Public gblsNumeric As String
Public gblQty As Integer


'Variable for Printer
Public sThermalPrinterName As String

'Variable for Login Time
Public dLoginTime As Date

'Variables for Printing Reports
Public OutputFileName As String
Public pRPTPath As String
Public cReceiptName As String
Public EmpReport As String
Public lUsePassword As Boolean

Public Property Get sSQL() As String
    sSQL = glbSSQL
End Property

Public Property Let sSQL(ByVal vNewValue As String)
    glbSSQL = vNewValue
End Property


