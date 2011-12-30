VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'IT Group Inc. 2005.07.17

Option Explicit

'Sub GetAccessLevel()
'
'    Set rs = New Recordset
'
'    sSQL = "SELECT DISTINCT [Module] FROM SEC_ACCESSLEVEL WHERE RoleID = '" & SecUserRole & "' AND cCompanyID = '" & COID & "'"
'    rs.Open sSQL, cn, adOpenForwardOnly, adLockReadOnly
'    Do Until rs.EOF
'        Select Case UCase(Trim(rs!Module))
'
'
'            Case "SERVITYPE"
'                ITGLedgerMain.mnuAR.Enabled = True
'                ITGLedgerMain.mnuProjServType.Enabled = True
'            '---------------------------------------------------------
'            ' Order and Billing Security
'            Case "OBSI"
'                ITGLedgerMain.mnuAR.Enabled = True
'                ITGLedgerMain.mnuARSalesInvoice.Enabled = True
'            Case "OBCMEE"
'                ITGLedgerMain.mnuAR.Enabled = True
'                ITGLedgerMain.mnuARCMEWTEVAT.Enabled = True
'
'            ' Credit and Collection Security
'            Case "CCOR"
'                ITGLedgerMain.mnuCC.Enabled = True
'                ITGLedgerMain.mnuARCollection.Enabled = True
'            Case "CCDEPOSIT"
'                ITGLedgerMain.mnuCC.Enabled = True
'                ITGLedgerMain.mnuARDepositSlip.Enabled = True
'            Case "CCRETURN"
'                ITGLedgerMain.mnuCC.Enabled = True
'                ITGLedgerMain.mnuRecievablesReturnCheck.Enabled = True
'            Case "CCPULLED"
'                ITGLedgerMain.mnuCC.Enabled = True
'                ITGLedgerMain.mnuRecievablesPulledOutCheck.Enabled = True
'            Case "CCREPLACE"
'                ITGLedgerMain.mnuCC.Enabled = True
'                ITGLedgerMain.mnuReceivablesCheckReplacement.Enabled = True
'
'            ' Supplies and Maintenance Security
'            Case "IPSRS"
'                ITGLedgerMain.mnuSam.Enabled = True
'                ITGLedgerMain.mnuInvRequisitionSlip.Enabled = True
'            Case "IPSIF"
'                ITGLedgerMain.mnuSam.Enabled = True
'                ITGLedgerMain.mnuInvSupplyIssue.Enabled = True
'
'            ' Purchasing and Payables Security
'            Case "PPPR"
'                ITGLedgerMain.mnuAp.Enabled = True
'                ITGLedgerMain.mnuPayablesPurchaseRequest.Enabled = True
'            Case "PPCANVASS"
'                ITGLedgerMain.mnuAp.Enabled = True
'                ITGLedgerMain.mnuPayablesCanvass.Enabled = True
'            Case "PPPO"
'                ITGLedgerMain.mnuAp.Enabled = True
'                ITGLedgerMain.mnuPayablesPurchaseORder.Enabled = True
'            Case "PPRFP"
'                ITGLedgerMain.mnuAp.Enabled = True
'                ITGLedgerMain.mnuPayablesRFP.Enabled = True
'            Case "PPCMDM"
'                ITGLedgerMain.mnuAp.Enabled = True
'                ITGLedgerMain.mnuPayablesAdjustment.Enabled = True
'            Case "PPVOUCHER"
'                ITGLedgerMain.mnuAp.Enabled = True
'                ITGLedgerMain.mnuPayablesVoucher.Enabled = True
'            Case "PPCHKISS"
'                ITGLedgerMain.mnuAp.Enabled = True
'                ITGLedgerMain.mnuPayablesCI.Enabled = True
'            Case "PPCHKCLR"
'                ITGLedgerMain.mnuAp.Enabled = True
'                ITGLedgerMain.mnuPayablesCC.Enabled = True
'            Case "PPSETUP"
'                ITGLedgerMain.mnuAp.Enabled = True
'                ITGLedgerMain.mnuPayablesSetup.Enabled = True
'            Case "PPCHKBBL"
'                ITGLedgerMain.mnuAp.Enabled = True
'                ITGLedgerMain.mnuPayablesBL.Enabled = True
'            Case "PPBANKCL"
'                ITGLedgerMain.mnuAp.Enabled = True
'                ITGLedgerMain.mnuPayablesBCL.Enabled = True
'
'            ' Inventory and Production Security
'            Case "IPWRR"
'                ITGLedgerMain.mnuInventory.Enabled = True
'                ITGLedgerMain.mnuPayablesWRR.Enabled = True
'            Case "IPWRR_AS"
'                ITGLedgerMain.mnuInventory.Enabled = True
'                ITGLedgerMain.mnuPayablesWRR_Asset.Enabled = True
'            Case "IPRGS"
'                ITGLedgerMain.mnuInventory.Enabled = True
'                ITGLedgerMain.mnuPayablesRS.Enabled = True
'
'            ' File Maintenance Security
'            Case "MACCOUNT"
'                ITGLedgerMain.mnuMaint.Enabled = True
'                ITGLedgerMain.mnuMaintAccount.Enabled = True
'            Case "MCLIENT"
'                ITGLedgerMain.mnuMaint.Enabled = True
'                ITGLedgerMain.mnuMaintClient.Enabled = True
'            Case "MSALESMAN"
'                ITGLedgerMain.mnuMaint.Enabled = True
'                ITGLedgerMain.mnuMaintSalesman.Enabled = True
'            Case "MEMPLOYEE"
'                ITGLedgerMain.mnuMaint.Enabled = True
'                ITGLedgerMain.mnuMaintEmployee.Enabled = True
'            Case "MPRODUCT"
'                ITGLedgerMain.mnuMaint.Enabled = True
'                ITGLedgerMain.mnuMaintProduct.Enabled = True
'            Case "MSUPPLY"
'                ITGLedgerMain.mnuMaint.Enabled = True
'                ITGLedgerMain.mnuMaintSupply.Enabled = True
'            Case "MWHSE"
'                ITGLedgerMain.mnuMaint.Enabled = True
'                ITGLedgerMain.mnuMaintWH.Enabled = True
'            Case "MMS"
'                ITGLedgerMain.mnuMaint.Enabled = True
'                ITGLedgerMain.mnuMaintMS.Enabled = True
'            Case "MPD"
'                ITGLedgerMain.mnuMaint.Enabled = True
'                ITGLedgerMain.mnuMaintPD.Enabled = True
'            Case "MCURRENCY"
'                ITGLedgerMain.mnuMaint.Enabled = True
'                ITGLedgerMain.mnuMaintCurrency.Enabled = True
'            Case "MEWT"
'                ITGLedgerMain.mnuMaint.Enabled = True
'                ITGLedgerMain.mnuMaintEWT.Enabled = True
'            Case "MEVAT"
'                ITGLedgerMain.mnuMaint.Enabled = True
'                ITGLedgerMain.mnuMaintEVAT.Enabled = True
'            Case "MCHECKS"
'                ITGLedgerMain.mnuMaint.Enabled = True
'                ITGLedgerMain.mnuMaintChecksOnHand.Enabled = True
'            Case "MBANK"
'                ITGLedgerMain.mnuMaint.Enabled = True
'                ITGLedgerMain.mnuMaintBank.Enabled = True
'            Case "MFASSETS"
'                ITGLedgerMain.mnuMaint.Enabled = True
'                ITGLedgerMain.mnuMaintFixedAssets.Enabled = True
'            Case "MPROFIT"
'                ITGLedgerMain.mnuMaint.Enabled = True
'                ITGLedgerMain.mnuMaintProfitCenter.Enabled = True
'            Case "MCOMPANY"
'                ITGLedgerMain.mnuMaint.Enabled = True
'                ITGLedgerMain.mnuMaintCompany.Enabled = True
'            Case "MOCUSTOMER"
'                ITGLedgerMain.mnuMaint.Enabled = True
'                ITGLedgerMain.mnuMaintClientOthers.Enabled = True
'            Case "MSUPPT"
'                ITGLedgerMain.mnuMaint.Enabled = True
'                ITGLedgerMain.mnuMaintSupplier.Enabled = True
'                ITGLedgerMain.mnuMaintSupplierTrade.Enabled = True
'            Case "MSUPPNT"
'                ITGLedgerMain.mnuMaint.Enabled = True
'                ITGLedgerMain.mnuMaintSupplier.Enabled = True
'                ITGLedgerMain.mnuMaintSupplierNT.Enabled = True
'            Case "MLOCATION"
'                ITGLedgerMain.mnuMaint.Enabled = True
'                ITGLedgerMain.mnuLogLocations.Enabled = True
'            Case "MFUND"
'                ITGLedgerMain.mnuMaint.Enabled = True
'                ITGLedgerMain.mnuMaintLoan.Enabled = True
'                ITGLedgerMain.mnuMaintTD.Enabled = True
'            Case "MAINTBUS"
'                ITGLedgerMain.mnuMaint.Enabled = True
'                ITGLedgerMain.mnuMaintBuss.Enabled = True
'
'            ' Reports Security
'            Case "REPAR"
'                'ITGLedgerMain.mnuReports.Enabled = True
'                'ITGLedgerMain.mnuRepAR.Enabled = True
'            Case "REPAP"
'                'ITGLedgerMain.mnuReports.Enabled = True
'                'ITGLedgerMain.mnuRepAP.Enabled = True
'            Case "REPINV"
'                'ITGLedgerMain.mnuReports.Enabled = True
'                'ITGLedgerMain.mnuRepINV.Enabled = True
'            Case "REPFIN"
'                'ITGLedgerMain.mnuReports.Enabled = True
'                'ITGLedgerMain.mnuRepFin.Enabled = True
'            Case "REPCM"
'                'ITGLedgerMain.mnuReports.Enabled = True
'                'ITGLedgerMain.mnuRepCM.Enabled = True
'            Case "REPML"
'                'ITGLedgerMain.mnuReports.Enabled = True
'                'ITGLedgerMain.mnuRepML.Enabled = True
'            Case "REPDIS"
'                'ITGLedgerMain.mnuReports.Enabled = True
'                'ITGLedgerMain.mnuRepDis.Enabled = True
'
'            ' Access Security
'            Case "SACCESS"
'                ITGLedgerMain.mnuSec.Enabled = True
'                ITGLedgerMain.mnuSecAccessLevel.Enabled = True
'            Case "SCACCESS"
'                ITGLedgerMain.mnuSec.Enabled = True
'                ITGLedgerMain.mnuSecCompanyAccess.Enabled = True
'            Case "SDEPT"
'                ITGLedgerMain.mnuSec.Enabled = True
'                ITGLedgerMain.mnuSecDepartment.Enabled = True
'            Case "SMODULE"
'                ITGLedgerMain.mnuSec.Enabled = True
'                ITGLedgerMain.mnuSecModule.Enabled = True
'            Case "SROLE"
'                ITGLedgerMain.mnuSec.Enabled = True
'                ITGLedgerMain.mnuSecUserRole.Enabled = True
'            Case "SUSER"
'                ITGLedgerMain.mnuSec.Enabled = True
'                ITGLedgerMain.mnuSecUserProfile.Enabled = True
'
'            ' Tools Security
'            Case "TSETUP"
'                ITGLedgerMain.mnuTools.Enabled = True
'                ITGLedgerMain.mnuToolsSetup.Enabled = True
'            Case "TSPARAM"
'                ITGLedgerMain.mnuTools.Enabled = True
'                ITGLedgerMain.mnuToolsSysParam.Enabled = True
'            Case "TUPARAM"
'                ITGLedgerMain.mnuTools.Enabled = True
'                ITGLedgerMain.mnuToolsUserParam.Enabled = True
'            Case "TMEXPLORE"
'                ITGLedgerMain.mnuTools.Enabled = True
'                ITGLedgerMain.mnuToolsExplorer.Enabled = True
'            '--Added
'            Case "TCLOSEMOD"
'                ITGLedgerMain.mnuTools.Enabled = True
'                ITGLedgerMain.mnuToolsClosingModule.Enabled = True
'            Case "TRECOMPUT"
'                ITGLedgerMain.mnuTools.Enabled = True
'                ITGLedgerMain.mnuToolsRecompute.Enabled = True
'            Case "TGLACTIVITY"
'                ITGLedgerMain.mnuTools.Enabled = True
'                ITGLedgerMain.mnuToolsGLActivity.Enabled = True
'
'            '-----------------------------------------------------------------
'
'
'            Case "TSOAPPROVE"
'                ITGLedgerMain.mnuTools.Enabled = True
'                ITGLedgerMain.mnuToolsSOApproval.Enabled = True
'            Case "TATRAIL"
'                ITGLedgerMain.mnuTools.Enabled = True
'                ITGLedgerMain.mnuToolsAuditTrail.Enabled = True
'            Case "TGLINT"
'                ITGLedgerMain.mnuTools.Enabled = True
'                ITGLedgerMain.mnuToolsGLInterface.Enabled = True
'        End Select
'    rs.MoveNext
'    Loop
'    rs.Close
'    Set rs = Nothing
'
'    'Set pull-out slip option variable
'    lPullOutSlip = GetValueFrTable("lBit", "SYSTEM_OPTION", "cCode = 'SR_PULLOUT_WRR'")
'
'    'Set NoCanvass
'    lNoCanvass = GetValueFrTable("lBit", "SYSTEM_OPTION", "cCode = 'PO_NOCANVASS'")
'
'End Sub
