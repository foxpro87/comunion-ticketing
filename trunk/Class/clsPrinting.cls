VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrinting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private cmdReportHeader As New ADODB.Command
Private cmdReportDetail As New ADODB.Command
Private rsReportHeader As New ADODB.Recordset
Private rsReportDetail() As New ADODB.Recordset 'Dynamic Array
Private rptTransID As String
Private rspHeader As String
Private rspDetail As String
Private rspComp As String

Dim i As Integer
Dim cParam As String
Dim nSubRpt As Integer
Dim cSubRpt
Dim cSubSPRpt                           'rspDetail

Public pCN As New Connection
Public pCOID As String
Public pModule As String

Public sDateTran As String
Public sTimeFrom As String
Public sTimeTo As String
Public RPTPath As String

Public ITGLedgerMDI As Object




Public Sub PrintReceipt(Optional cTransID As String)
    Dim sModule As String
    
    Dim rptPrint As CRAXDRT.REPORT
    Dim ReportApp As New CRAXDDRT.Application

    On Error GoTo TheSource

    Dim cSubRptName As String
    Dim cSubRptCmp As String
    
    'Set pITGLedgerMDI = ITGLedgerMDI
    
    rptTransID = cTransID
    pRPTPath = RPTPath
        
    Dim tmpRS As New Recordset
    tmpRS.Open "select cCode from sec_tran_report where cReportHeader = '" & pModule & "'", pCN, adOpenStatic, adLockReadOnly
    sModule = tmpRS!cCode
    tmpRS.Close
    
    Dim strSQL As String
    strSQL = "select lIsSP,nSubRptNo,cReportHeader,cReportDetail,cTranHeader,cTranDetail,cParam from sec_tran_report where cCode = '" & sModule & "' and cCompanyID = '" & pCOID & "'"
    tmpRS.Open strSQL, pCN, adOpenStatic, adLockReadOnly
    If tmpRS.RecordCount <> 0 Then
        cReceiptName = tmpRS!cReportHeader 'IIf(RepName = "", tmpRS!cReporHeader, RepName)
        cSubRptName = tmpRS!cReportDetail
        rspHeader = tmpRS!cTranHeader
        rspDetail = tmpRS!cTranDetail
        cParam = tmpRS!cParam
        nSubRpt = tmpRS!nSubRptNo

        cSubRpt = Split(cSubRptName, ",")
        cSubSPRpt = Split(rspDetail, ",")
    Else
        Exit Sub
    End If
    
    'RepName = ""
    tmpRS.Close
        
    OpenReportRS cTransID
    
    Set rptPrint = ReportApp.OpenReport(RPTPath & "\" & cReceiptName & ".rpt") 'App.Path & "\Crystal Reports\Transaction\" & cReceiptName & ".rpt")
    DoEvents
    
    'For Dynamic Array SubReport
    If nSubRpt >= 1 Then
        For i = 0 To nSubRpt - 1
            rptPrint.OpenSubreport(cSubRpt(i)).Database.SetDataSource rsReportDetail(i)
        Next i
    End If
    rptPrint.Database.SetDataSource rsReportHeader
    
    frmReportViewer.Show
    Dim obj As New Application
    With frmReportViewer.ITGReportViewer
        .ReportSource = rptPrint
        .PrintReport
        .ViewReport
        .DisplayGroupTree = False
        .Zoom (100)
    End With
             
    ExportToPDF rptPrint
    
    'Audit trail
    UpdateLogFile pModule, cTransID, "Printed"

TheSource:
    If err.Number = -2147206461 Then 'File not found

        MsgBox "Report file '" & cReceiptName & ".rpt' not found.", vbCritical, "ComUnion"
    ElseIf err.Number = 3265 Then
        MsgBox err.Description
        MsgBox "Stored Procedure '" & rspHeader & "' not found.", vbCritical, "ComUnion"
    End If

    Set rsReportHeader = Nothing
    
    For i = 0 To nSubRpt - 1
        Set rsReportDetail(i) = Nothing
    Next i
    
    Set cmdReportHeader = Nothing
    Set cmdReportDetail = Nothing
    Set rptPrint = Nothing
    Set ReportApp = Nothing

End Sub

Sub OpenReportRS(Optional cTransID As String)
  
    rsReportHeader.CursorLocation = adUseClient
    With cmdReportHeader
        .ActiveConnection = pCN
        .CommandTimeout = 1000
        .CommandText = rspHeader
        .CommandType = adCmdStoredProc
        
        Dim tmpRS As New Recordset
        Dim a
        Dim strSQL As String
        strSQL = "select cType,cParamName,cModule,cDesc,cValue from parameter_user where cType = 'SQLReportSettings' and cCompanyID ='" & pCOID & "'"
        tmpRS.Open strSQL, pCN, adOpenStatic, adLockReadOnly
        a = Split(cParam, ",")
        Do Until tmpRS.EOF
            For i = 0 To UBound(a)
                If (tmpRS!cValue = a(i)) Then
                    Select Case tmpRS!cParamName
                        Case "TranCompany"
                            .Parameters(a(i)) = pCOID
                        Case "TranNo"
                            .Parameters(a(i)) = rptTransID
'                        Case "TranType"
'                            .Parameters(a(i)) = cTypeVoucher    'temporary
                        Case "TranDate"
                            .Parameters(a(i)) = sDateTran
                        Case "TranUser"
                            .Parameters(a(i)) = sUserName
                            
                        Case "TimeFrom"
                            .Parameters(a(i)) = sTimeFrom
                        Case "TimeTo"
                            .Parameters(a(i)) = sTimeTo
                            
                    End Select
                End If
            Next i
            tmpRS.MoveNext
        Loop
        tmpRS.Close
        
    End With
    
    'For Sub Report Manipulation
    If nSubRpt >= 1 Then                                                                                           'if the report has a sub report
        ReDim rsReportDetail(0 To nSubRpt - 1) As New ADODB.Recordset   'Redim the dynamic array
        For i = 0 To nSubRpt - 1
            With cmdReportDetail
                .ActiveConnection = cn
                .CommandTimeout = 1000
                .CommandText = cSubSPRpt(i) 'rspDetail
                .CommandType = adCmdStoredProc
                .Parameters("@cTranNo") = rptTransID
                .Parameters("@cCompID") = COID
            End With
            rsReportDetail(i).CursorLocation = adUseClient
            Set rsReportDetail(i) = cmdReportDetail.Execute()
        Next i
    End If
       
    Set rsReportHeader = cmdReportHeader.Execute()

End Sub


Public Sub ShowRpt_Form() 'sLink As String)
    Set FrmName = frmReportViewer
    FrmName.Show
    FrmName.WindowState = 2
End Sub

Sub ExportToPDF(rptPrint As CRAXDRT.REPORT)
    On Error GoTo ErrHandler
    If tmpRS.State = 1 Then tmpRS.Close
    tmpRS.Open "select * from MailSetup where cEmailSetupFor = '" & pModule & "'", pCN, adOpenStatic, adLockReadOnly
    OutputFileName = tmpRS!cAttachment
    frmReportViewer.sPWord = IIf(lUsePassword, tmpRS!cEmailSetupFor & Format(Now, "mmddyy"), "")
    tmpRS.Close
    
    If FileExists(OutputFileName) = False Then
        CreateFile OutputFileName
    End If
        
    Set frmReportViewer.rpt = rptPrint
    With frmReportViewer.rpt
        .ExportOptions.DiskFileName = OutputFileName
        .ExportOptions.DestinationType = 1 'crEDTDiskFile
        .ExportOptions.FormatType = 31 'crEFTPortableDocFormat
        .ExportOptions.PDFExportAllPages = True
        .Export False
        .PrinterSetup 0
    End With
    
ErrHandler:
    If err.Number <> 0 Then
        MsgBox err.Description
    End If
End Sub

