VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'Public oReport As New clsReport
'Dim oConnection As New clsConnection
'Dim connReport As ADODB.Connection
'Dim rsT As New Recordset
'
'Public Sub Process()
'On Error GoTo ErrLog
'Dim MyReport As New REPORT
'Dim ReportApp As New CRAXDDRT.Application
'Dim crRep As CRAXDRT.REPORT
'
'Dim cmdReport As ADODB.Command
'Dim rsReport As ADODB.Recordset
'Dim PathFile As String
'
'Dim i As Integer
'Dim cParam As String
'
''For ttx data source--------------------------------------------------------------------------------------------------------------
'  Set cmdReport = New ADODB.Command
'  Set rsReport = New ADODB.Recordset
'  rsReport.CursorLocation = adUseClient
'
'    With cmdReport
'        .ActiveConnection = cn
'        .CommandTimeout = 1000
'        .CommandText = cStorProcName
'        .CommandType = adCmdStoredProc
'
'        .Parameters("@cFrom") = Trim$(gReportDataFrom)
'        .Parameters("@cTo") = Trim$(gReportDataTo)
'        .Parameters("@cCompID") = COID
'        .Parameters("@cCompName") = cCompany
'        .Parameters("@cAddress1") = cAddress1
'        .Parameters("@cAddress2") = cAddress2
'        .Parameters("@cReportTitle") = cReportTitle
'        .Parameters("@cCriteria1") = cCriteria1
'        .Parameters("@cCriteria") = cCriteria2
'        .Parameters("@dDateFrom") = gReportDateFrom
'        .Parameters("@dDateTo") = gReportDateTo
'        If GetValueFrTable("lAging", "REPORT", "cReportID = '" & cReport & "'") = True Then
'            .Parameters("@nInterval") = nVarInterval
'            .Parameters("@cValue1") = "1 - " & nVarInterval
'            .Parameters("@cValue2") = str(nVarInterval + 1) & " - " & str(nVarInterval * 2)
'            .Parameters("@cValue3") = str((nVarInterval * 2) + 1) & " - " & str(nVarInterval * 3)
'            .Parameters("@cValue4") = str((nVarInterval * 3) + 1) & " - " & str(nVarInterval * 4)
'            .Parameters("@cValue5") = "Over " & str(nVarInterval * 5)
'        End If
'        If GetValueFrTable("lFilter", "REPORT", "cReportID = '" & cReport & "'") = True Then
'            .Parameters("@cFilter") = cFilter
'        End If
'    End With
'
'    Set rsReport = cmdReport.Execute()
'
'    If rsReport.State = 0 Then
'        MsgBox "No Report to View...", vbInformation, "Caution"
'        Exit Sub
'    End If
'
'    PathFile = "C:\" & Mid(cReportName, 1, Len(cReportName) - 4)
'
'    Set MyReport = ReportApp.OpenReport(App.Path & "\crystal reports\" & cReportName)
'    MyReport.Database.SetDataSource rsReport
'
'    With MyReport
'        .FormulaFields.GetItemByName("cCompanyName").Text = "'" & cCompany & "'"
'        .FormulaFields.GetItemByName("cAddress1").Text = "'" & cAddress1 & "'"
'        .FormulaFields.GetItemByName("cAddress2").Text = "'" & cAddress2 & "'"
'        .FormulaFields.GetItemByName("cReportTitle").Text = "'" & cReportTitle & "'"
'        .FormulaFields.GetItemByName("cCriteria1").Text = "'" & cCriteria1 & "'"
'        .FormulaFields.GetItemByName("cCriteria2").Text = "'" & cCriteria2 & "'"
'        If GetValueFrTable("lAging", "REPORT", "cReportID = '" & cReport & "'") = True Then
'            .FormulaFields.GetItemByName("cHeader1").Text = "'" & "1 - " & nVarInterval & "'"
'            .FormulaFields.GetItemByName("cHeader2").Text = "'" & str(nVarInterval + 1) & " - " & str(nVarInterval * 2) & "'"
'            .FormulaFields.GetItemByName("cHeader3").Text = "'" & str((nVarInterval * 2) + 1) & " - " & str(nVarInterval * 3) & "'"
'            .FormulaFields.GetItemByName("cHeader4").Text = "'" & str((nVarInterval * 3) + 1) & " - " & str(nVarInterval * 4) & "'"
'            .FormulaFields.GetItemByName("cHeader5").Text = "'" & "Over " & str(nVarInterval * 5) & "'"
'        End If
'    End With
'    If Trim$(frmReportCriteria.cboPrint.Text) = "Preview" Then
'        With frmReportViewer.ITGReportViewer
'            .ReportSource = MyReport
'            .DisplayGroupTree = False
'            .ViewReport
'            .Zoom 100
'        End With
'        frmReportViewer.Caption = cReportTitle
'        frmReportViewer.Show
'    ElseIf Trim$(frmReportCriteria.cboPrint.Text) = "Printer" Then
'        MyReport.PrinterSetup 0
'    ElseIf Trim$(frmReportCriteria.cboPrint.Text) = "File" Then
'        MyReport.EnableParameterPrompting = False
'        MyReport.ExportOptions.DestinationType = crEDTDiskFile
'        If cFilePreview = "Text File" Then
'            PathFile = PathFile & ".TXT"
'            MyReport.ExportOptions.DiskFileName = PathFile
'            MyReport.ExportOptions.FormatType = crEFTText
'        ElseIf cFilePreview = "Word Document" Then
'            PathFile = PathFile & ".DOC"
'            MyReport.ExportOptions.DiskFileName = PathFile
'            MyReport.ExportOptions.FormatType = crEFTWordForWindows
'        ElseIf cFilePreview = "Excel" Then
'            PathFile = PathFile & ".XLS"
'            MyReport.ExportOptions.DiskFileName = PathFile
'            MyReport.ExportOptions.FormatType = crEFTExcel80Tabular
'            MyReport.ExportOptions.ExcelTabHasColumnHeadings = True
'        End If
'        MyReport.Export False
'    End If
'    Set rsReport = Nothing
'ErrLog:
'    If Err.Number <> 0 Then
'        MsgBox "An error occured while accessing the report, please contact your administrator!", vbCritical
'    End If
'End Sub
'
'Public Sub ProcessReport(rs As Recordset, Linkedserver As Boolean)
'On Error GoTo ErrLog
'Dim MyReport As New REPORT
'Dim ReportApp As New CRAXDDRT.Application
'
'Dim cmdReport As ADODB.Command
'Dim rsReport As ADODB.Recordset
'
'Dim cmdReportHeading As New ADODB.Command
'Dim rsReportHeading As New ADODB.Recordset
'
'Dim PathFile As String
'
'Dim i As Integer
'Dim cParam As String
'
'    rsReportHeading.CursorLocation = adUseClient
'
'    Set rsReport = New ADODB.Recordset
'    If Not Linkedserver Then
'        Set cmdReport = New ADODB.Command
'        rsReport.CursorLocation = adUseClient
'
'            With cmdReport
'                .ActiveConnection = cn
'                .CommandTimeout = 1000
'                .CommandText = cStorProcName
'                .CommandType = adCmdStoredProc
'                .Parameters("@dDateFrom") = gReportDateFrom
'                .Parameters("@dDateTo") = gReportDateTo
'                .Parameters("@cCompID") = COID
'
'                If rs.RecordCount <> 0 Then
'                   rs.MoveFirst
'                   Do Until rs.EOF
'                        .Parameters(Trim(rs!cSQLParam)) = IIf(IsNull(rs!cValue), "", rs!cValue)
'                        rs.MoveNext
'                   Loop
'                End If
'
'                If cReport = "RP01-20" Then
'                    .Parameters("@nInterval") = nVarInterval
'                ElseIf cReport = "60000617" Then
'                    .Parameters("@lBit") = "1"
'                    .Parameters("@cValue1") = ""
'                    .Parameters("@cValue2") = ""
'                    .Parameters("@cValue3") = ""
'                    .Parameters("@cValue4") = ""
'                    .Parameters("@cValue5") = ""
'                End If
'
'            End With
'
'
'            Set rsReport = cmdReport.Execute()
'        Else
'
'            If rs.RecordCount <> 0 Then
'               rs.MoveFirst
'               Do Until rs.EOF
'                    cParam = cParam & "'" & IIf(IsNull(rs!cValue), "", rs!cValue) & "',"
'                    rs.MoveNext
'               Loop
'            End If
'
'            cParam = Left(cParam, Len(cParam) - 1)
'
''            sSQL = "EXEC " & sLinkedServer & Trim(cStorProcName) & "'" & Trim(gReportDateFrom) & "','" & Trim(gReportDateTo) & "','" _
''                    & Trim(COID) & "'," & cParam
''            rsReport.Open sSQL, cn, adOpenKeyset, adLockReadOnly
'        End If
'
'
'        With cmdReportHeading
'            .ActiveConnection = cn
'            .CommandTimeout = 1000
'            .CommandText = "rsp_Report_Header"
'            .CommandType = adCmdStoredProc
'            .Parameters("@cCompID") = COID
'        End With
'        Set rsReportHeading = cmdReportHeading.Execute()
'
'    If rsReport.State = 0 Then
'        MsgBox "No Report to View...", vbInformation, "Caution"
'        Exit Sub
'    End If
'
'    PathFile = "C:\" & Mid(cReportName, 1, Len(cReportName) - 4)
'
'    Set MyReport = ReportApp.OpenReport(App.Path & "\crystal reports\" & cReportName)
'    MyReport.Database.SetDataSource rsReport
'    MyReport.OpenSubreport("Report_Header").Database.SetDataSource rsReportHeading
'
'    With MyReport
''        .FormulaFields.GetItemByName("cCompanyName").Text = "'" & cCompany & "'"
''        .FormulaFields.GetItemByName("cAddress1").Text = "'" & cAddress1 & "'"
''        .FormulaFields.GetItemByName("cAddress2").Text = "'" & cAddress2 & "'"
'        .FormulaFields.GetItemByName("cReportTitle").Text = "'" & cReportTitle & "'"    '----get the Report Title
''        .FormulaFields.GetItemByName("cCriteria1").Text = "'" & cCriteria1 & "'"
''        .FormulaFields.GetItemByName("cCriteria2").Text = "'" & cCriteria2 & "'"
''        If GetValueFrTable("lAging", "REPORT", "cReportID = '" & cReport & "'") = True Then
''            .FormulaFields.GetItemByName("cHeader1").Text = "'" & "1 - " & nVarInterval & "'"
''            .FormulaFields.GetItemByName("cHeader2").Text = "'" & str(nVarInterval + 1) & " - " & str(nVarInterval * 2) & "'"
''            .FormulaFields.GetItemByName("cHeader3").Text = "'" & str((nVarInterval * 2) + 1) & " - " & str(nVarInterval * 3) & "'"
''            .FormulaFields.GetItemByName("cHeader4").Text = "'" & str((nVarInterval * 3) + 1) & " - " & str(nVarInterval * 4) & "'"
''            .FormulaFields.GetItemByName("cHeader5").Text = "'" & "Over " & str(nVarInterval * 5) & "'"
''        End If
'  '      .PrintOut False
'    End With
'
'    If Trim$(frmITGReportCriteria.cboPrint.Text) = "Preview" Then
'        With frmReportViewer.ITGReportViewer
'            .ReportSource = MyReport
'            .ViewReport
'            .Zoom 100
'        End With
'        frmReportViewer.Caption = cReportTitle
'        frmReportViewer.Show
'    ElseIf Trim$(frmITGReportCriteria.cboPrint.Text) = "Printer" Then
'        MyReport.PrinterSetup 0
'    ElseIf Trim$(frmITGReportCriteria.cboPrint.Text) = "File" Then
'        MyReport.EnableParameterPrompting = False
'        MyReport.ExportOptions.DestinationType = crEDTDiskFile
'        If cFilePreview = "Text File" Then
'            PathFile = PathFile & ".TXT"
'            MyReport.ExportOptions.DiskFileName = PathFile
'            MyReport.ExportOptions.FormatType = crEFTText
'        ElseIf cFilePreview = "Word Document" Then
'            PathFile = PathFile & ".DOC"
'            MyReport.ExportOptions.DiskFileName = PathFile
'            MyReport.ExportOptions.FormatType = crEFTWordForWindows
'        ElseIf cFilePreview = "Excel" Then
'            PathFile = PathFile & ".XLS"
'            MyReport.ExportOptions.DiskFileName = PathFile
'            MyReport.ExportOptions.FormatType = crEFTExcel80Tabular
'            MyReport.ExportOptions.ExcelTabHasColumnHeadings = True
'        End If
'        MyReport.Export False
'    End If
'    Set rsReport = Nothing
'ErrLog:
'    If Err.Number <> 0 Then
'        MsgBox Err.Description + " Please contact your administrator!", vbCritical
'    End If
'End Sub
'
'
