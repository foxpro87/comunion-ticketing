VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'IT Group Inc. 2005.09.23

Option Explicit

Private mcnDBFConn As ADODB.Connection
Private mcnSQLConn As ADODB.Connection
Private mblnOverwrite As Boolean

Private mstrDBFPath As String

Private arrSales() As String
Private arrSales_T() As String
Private arrPR() As String
Private arrPR_T() As String
Private arrPR_C() As String
Private arrAR() As String
Private arrAR_T() As String
Private arrPRC() As String

'PO [WRR]
Private arrPO() As String
Private arrPO_T() As String

'RSALES
Private arrRSales() As String
Private arrRSales_T() As String

Private strDL As String

Public Property Get oDBFConn() As ADODB.Connection
    Set oDBFConn = mcnDBFConn
End Property

Public Property Set oDBFConn(ByVal toNewDBFConn As ADODB.Connection)
    Set mcnDBFConn = toNewDBFConn
    mstrDBFPath = mcnDBFConn.DefaultDatabase
End Property

Public Property Get oSQLConn() As ADODB.Connection
    Set oSQLConn = mcnSQLConn
End Property

Public Property Set oSQLConn(ByVal toNewSQLConn As ADODB.Connection)
    Set mcnSQLConn = toNewSQLConn
End Property

Public Property Get lOverwrite() As Boolean
    lOverwrite = mblnOverwrite
End Property

Public Property Let lOverwrite(ByVal tblnNewOverwrite As Boolean)
    mblnOverwrite = tblnNewOverwrite
End Property

Public Function DownloadSales(tdteFrom As Date, tdteTo As Date) As Boolean
    Dim llOK As Boolean, lcCmdString As String
    Dim lcCmdSelectSales As String, lcCmdSelectSales_T As String
    Dim lcCmdInsertSales As String, lcCmdInsertSales_T As String
    Dim rsSales As ADODB.Recordset
    Dim rsSales_T As ADODB.Recordset
        
    lcCmdSelectSales = "SELECT A.cInvNo, A.cCode, B.cName, A.dDate, ISNULL((CONVERT(NUMERIC, D.cValue)), 0) AS nTerm, A.cSMan, C.cCode AS cSMName " & _
        "FROM SALES A LEFT OUTER JOIN CLIENT_CUSTOMER B ON A.cCode = B.cCode AND A.cCompanyID = B.cCompanyID " & _
        "LEFT OUTER JOIN SALESMAN C ON A.cSMan = C.cCode AND A.cCompanyID = C.cCompanyID " & _
        "LEFT OUTER JOIN PARAMETER_USER D ON A.cTerm = D.cParamName AND A.cCompanyID = D.cCompanyID " & _
        "WHERE A.dDate BETWEEN '" & CStr(tdteFrom) & "' AND '" & CStr(tdteTo) & "' " & _
        " AND (A.cCode IS NOT NULL AND RTRIM(LTRIM(A.cCode)) <> '') AND A.lCancelled = 0 AND A.cCompanyID = '" & COID & "'" & _
        " ORDER BY A.dDate "
        
    lcCmdSelectSales_T = "SELECT A.cInvNo, A.cItemNo, A.cDesc, " & _
        "A.nQty, A.nPrice, A.cUnit " & _
        "FROM SALES_T A " & _
        "INNER JOIN SALES B ON A.cInvNo = B.cInvNo AND A.cCompanyID = B.cCompanyID " & _
        "WHERE (A.cInvNo IN (SELECT cInvNO FROM SALES " & _
        "WHERE dDate BETWEEN '" & CStr(tdteFrom) & "' AND " & _
        "'" & CStr(tdteTo) & "' AND (cCode IS NOT NULL AND " & _
        "RTRIM(LTRIM(cCode)) <> '') AND B.lCancelled = 0 AND A.cCompanyID = '" & COID & "')) " & _
        "ORDER BY A.cInvNo "
    
    llOK = CreateTable("SALES")
    If llOK Then
        llOK = CreateTable("SALES_T")
    End If
    
    If llOK Then
        Set rsSales = oSQLConn.Execute(lcCmdSelectSales)
        If Not rsSales Is Nothing Then
            Do While Not rsSales.EOF
                lcCmdInsertSales = "INSERT INTO SALES " & _
                    "(cInvNo, cCode, cName, dDate, nTerm, cSMan, cSMName) " & _
                    "VALUES (" & " '" & Trim(rsSales!cInvNo) & "', " & _
                                 "'" & Trim(rsSales!cCode) & "', " & _
                                 "'" & Replace(Trim(rsSales!cName), "'", " ") & "', " & _
                                 "'" & CStr(rsSales!dDate) & "', " & _
                                 "  " & Trim(str(rsSales!nTerm)) & " , " & _
                                 "'" & Trim(rsSales!cSMan) & "', " & _
                                 "'" & Trim(rsSales!cSMName) & "' " & _
                            ")"
                oDBFConn.Execute (lcCmdInsertSales)
                rsSales.MoveNext
            Loop
        End If
              
        Set rsSales_T = oSQLConn.Execute(lcCmdSelectSales_T)
        If Not rsSales_T Is Nothing Then
            Do While Not rsSales_T.EOF
                lcCmdInsertSales_T = "INSERT INTO SALES_T " & _
                    "(cInvNo, cItemNo, cDesc, cUnit, nQty, nPrice) " & _
                    "VALUES (" & "'" & Trim(rsSales_T!cInvNo) & "', " & _
                                 "'" & Trim(rsSales_T!cItemNo) & "', " & _
                                 "'" & Replace(Trim(rsSales_T!cDesc), "'", " ") & "', " & _
                                 "'" & CStr(rsSales_T!cUnit) & "', " & _
                                 "  " & Trim(str(rsSales_T!nQty)) & " , " & _
                                 "  " & Trim(str(rsSales_T!nPrice)) & " " & _
                            ")"
                oDBFConn.Execute (lcCmdInsertSales_T)
                rsSales_T.MoveNext
            Loop
        End If
    End If
    
    Set rsSales = Nothing
    Set rsSales_T = Nothing
    DownloadSales = llOK
End Function

'************PR_C ONLY*************'
Public Function DownloadPR_C(tdteFrom As Date, tdteTo As Date) As Boolean
Dim llOK As Boolean, lcCmdString As String
Dim lcCmdSelectPR_C As String
Dim lcCmdInsertPR_C As String
Dim rsPR_C As ADODB.Recordset
    
    lcCmdSelectPR_C = "SELECT A.cTransNo as TransNo, B.dDate as TDate, B.cCode as Code, C.cName as Name, A.cType as Type, A.cCheckNo as [Check], " & _
                "A.dCheckDate as Date, A.nAmtCollect as Amount FROM PR_C A INNER JOIN PR B " & _
                "ON A.cTransNo = B.cTransNo INNER JOIN CUSTOMER C " & _
                "ON B.cCode = C.cCode WHERE " & _
                "B.dDate BETWEEN '" & CStr(tdteFrom) & "' AND '" & CStr(tdteTo) & "' AND B.lPost = 1" & _
                " ORDER BY B.dDate "
    
    strDL = "PR_C2"
    llOK = CreateTable("PRC")
        
    If llOK Then
    
        Set rsPR_C = oSQLConn.Execute(lcCmdSelectPR_C)
        If Not rsPR_C Is Nothing Then
            Do While Not rsPR_C.EOF
                lcCmdInsertPR_C = "INSERT INTO PR_C " & _
                    "(TransNo, TDate, Code, [Name], Type, [Check], [Date], Amount) " & _
                    "VALUES (" & "'" & Trim(rsPR_C!TransNo) & "', " & _
                                 "'" & CStr(rsPR_C!TDate) & "', " & _
                                 "'" & Trim(rsPR_C!Code) & "', " & _
                                 "'" & CStr(rsPR_C!Name) & "', " & _
                                 "'" & Trim(rsPR_C!Type) & "', " & _
                                 "'" & Trim(rsPR_C!Check) & "', " & _
                                 "'" & CStr(rsPR_C!Date) & "', " & _
                                 " " & Trim(str(rsPR_C!Amount)) & " " & _
                            ")"
                oDBFConn.Execute (lcCmdInsertPR_C)
                rsPR_C.MoveNext
            Loop
        End If
    End If
    
    Set rsPR_C = Nothing
    DownloadPR_C = llOK
    
End Function

Public Function DownloadPR(tdteFrom As Date, tdteTo As Date) As Boolean
    Dim llOK As Boolean, lcCmdString As String
    Dim lcCmdSelectPR As String, lcCmdSelectPR_T As String, lcCmdSelectPR_C As String
    Dim lcCmdInsertPR As String, lcCmdInsertPR_T As String, lcCmdInsertPR_C As String
    Dim rsPR As ADODB.Recordset
    Dim rsPR_T As ADODB.Recordset
    Dim rsPR_C As ADODB.Recordset
    
    lcCmdSelectPR = "SELECT A.cTranNo, A.cCode, B.cName, A.dDate " & _
        "FROM PR A LEFT OUTER JOIN CLIENT_CUSTOMER B ON A.cCode = B.cCode AND A.cCompanyID = B.cCompanyID " & _
        "WHERE A.dDate BETWEEN '" & CStr(tdteFrom) & "' AND '" & CStr(tdteTo) & "' AND A.lCancelled = 0 AND A.cCompanyID = '" & COID & "' " & _
        "ORDER BY A.dDate"
        
    lcCmdSelectPR_T = "SELECT A.cTranNo, A.cInvNo, A.dInvDate AS dDate, A.nInvBalance AS nAmount, " & _
        "A.nApplied FROM PR_T A WHERE A.cTranNo IN " & _
        "(SELECT cTranNo FROM PR WHERE dDate " & _
        "BETWEEN '" & CStr(tdteFrom) & "' AND '" & CStr(tdteTo) & "' AND lCancelled = 0 AND cCompanyID = '" & COID & "') " & _
        "AND A.cCompanyID = '" & COID & "' ORDER BY cTranNo"

    lcCmdSelectPR_C = "SELECT A.cTranNo, A.cType, ISNULL(A.cBankID, '') AS cBank, ISNULL(A.cCheckNo, '') AS cCheckNo, A.dCheckDate, A.nAmount FROM PR_C A " & _
        "WHERE A.cTranNo IN (SELECT cTranNo FROM PR " & _
        "WHERE dDate BETWEEN '" & CStr(tdteFrom) & "' AND '" & CStr(tdteTo) & "' AND lCancelled = 0 AND cCompanyID = '" & COID & "') " & _
        "AND A.cCompanyID = '" & COID & "' ORDER BY cTranNo"
        
    llOK = CreateTable("PR")
    If llOK Then
        llOK = CreateTable("PR_T")
    End If
    
    If llOK Then
        llOK = CreateTable("PR_C")
    End If
    
    If llOK Then
        Set rsPR = oSQLConn.Execute(lcCmdSelectPR)
        If Not rsPR Is Nothing Then
            Do While Not rsPR.EOF
                lcCmdInsertPR = "INSERT INTO PR " & _
                    "(cTranNo, cCode, cName, dDate) " & _
                    "VALUES (" & " '" & Trim(rsPR!cTranno) & "', " & _
                                 "'" & Trim(rsPR!cCode) & "', " & _
                                 "'" & Replace(Trim(rsPR!cName), "'", " ") & "', " & _
                                 "'" & CStr(rsPR!dDate) & "' " & _
                            ")"
                oDBFConn.Execute (lcCmdInsertPR)
                rsPR.MoveNext
            Loop
        End If
              
        Set rsPR_T = oSQLConn.Execute(lcCmdSelectPR_T)
        If Not rsPR_T Is Nothing Then
            Do While Not rsPR_T.EOF
                lcCmdInsertPR_T = "INSERT INTO PR_T " & _
                    "(cTranNo, cInvNo, dDate, nAmount, nApplied ) " & _
                    "VALUES (" & "'" & Trim(rsPR_T!cTranno) & "', " & _
                                 "'" & Trim(rsPR_T!cInvNo) & "', " & _
                                 "'" & CStr(rsPR_T!dDate) & "', " & _
                                 "  " & Trim(str(rsPR_T!nAmount)) & " , " & _
                                 "  " & Trim(str(rsPR_T!nApplied)) & " " & _
                            ")"
                oDBFConn.Execute (lcCmdInsertPR_T)
                rsPR_T.MoveNext
            Loop
        End If
        
            
        Set rsPR_C = oSQLConn.Execute(lcCmdSelectPR_C)
        If Not rsPR_C Is Nothing Then
            Do While Not rsPR_C.EOF
                
                lcCmdInsertPR_C = "INSERT INTO PR_C " & _
                    "(cTranNo , cType, cBank, cCheckNo, dCheckDate, nAmount) " & _
                    "VALUES (" & "'" & Trim(rsPR_C!cTranno) & "', " & _
                                 "'" & Trim(rsPR_C!cType) & "', " & _
                                 "'" & Trim(rsPR_C!cBank) & "" & "', " & _
                                 "'" & Trim(rsPR_C!cCheckNo) & "', " & _
                                 "'" & CStr(rsPR_C!dCheckDate) & "', " & _
                                 "  " & Trim(str(rsPR_C!nAmount)) & "  " & _
                            ")"
                oDBFConn.Execute (lcCmdInsertPR_C)
                rsPR_C.MoveNext
            Loop
        End If
        
    End If
    
    Set rsPR = Nothing
    Set rsPR_T = Nothing
    Set rsPR_C = Nothing
    DownloadPR = llOK
End Function

Public Function DownloadAR(tdteFrom As Date, tdteTo As Date) As Boolean
    Dim llOK As Boolean, lcCmdString As String
    Dim lcCmdSelectAR As String, lcCmdSelectAR_T As String
    Dim lcCmdInsertAR As String, lcCmdInsertAR_T As String
    Dim rsAR As ADODB.Recordset
    Dim rsAR_T As ADODB.Recordset
        
    lcCmdSelectAR = "SELECT A.cTransNo AS cTranNo, A.cCode, B.cName, A.cType, A.cInvNo, A.dDate " & _
        "FROM AR A LEFT OUTER JOIN Customer B ON A.cCode = B.cCode " & _
        "WHERE A.dDate BETWEEN '" & CStr(tdteFrom) & "' AND '" & CStr(tdteTo) & "' AND A.lPost = 1" & _
        " ORDER BY A.dDate "
        
    lcCmdSelectAR_T = "SELECT cTransNo AS cTranNo, cAcctNo, cAcctName AS cTitle, ISNULL(cDesc, '') AS cRemarks, " & _
        "ISNULL(nDebit,0) AS nDebit, ISNULL(nCredit,0) AS nCredit FROM AR_T WHERE cTransNo IN (SELECT cTransNo FROM AR WHERE dDate  " & _
        "BETWEEN '" & CStr(tdteFrom) & "' AND '" & CStr(tdteTo) & "' AND lPost = 1" & _
        ") ORDER BY cTransNo"
    
    llOK = CreateTable("AR")
    If llOK Then
        llOK = CreateTable("AR_T")
    End If
    
    If llOK Then
        Set rsAR = oSQLConn.Execute(lcCmdSelectAR)
        If Not rsAR Is Nothing Then
            Do While Not rsAR.EOF
                lcCmdInsertAR = "INSERT INTO AR " & _
                    "(cTranNo, cCode, cName, cType, cInvNo, dDate) " & _
                    "VALUES (" & " '" & Trim(rsAR!cTranno) & "', " & _
                                 "'" & Trim(rsAR!cCode) & "', " & _
                                 "'" & Replace(Trim(rsAR!cName), "'", " ") & "', " & _
                                 "'" & Trim(rsAR!cType) & "', " & _
                                 "'" & Trim(rsAR!cInvNo) & "', " & _
                                 "'" & CStr(rsAR!dDate) & "' " & _
                            ")"
                oDBFConn.Execute (lcCmdInsertAR)
                rsAR.MoveNext
            Loop
        End If
              
        Set rsAR_T = oSQLConn.Execute(lcCmdSelectAR_T)
        If Not rsAR_T Is Nothing Then
            Do While Not rsAR_T.EOF
                lcCmdInsertAR_T = "INSERT INTO AR_T " & _
                    "(cTranNo, cAcctNo, cTitle, cRemarks, nDebit, nCredit) " & _
                    "VALUES (" & "'" & Trim(rsAR_T!cTranno) & "', " & _
                                 "'" & Trim(rsAR_T!cAcctNo) & "', " & _
                                 "'" & Replace(Trim(rsAR_T!cTitle), "'", " ") & "', " & _
                                  "'" & Replace(IIf(IsNull(Trim(rsAR_T!cRemarks)), "", Trim(rsAR_T!cRemarks)), "'", " ") & "', " & _
                                 "  " & IIf(IsNull(str(rsAR_T!nDebit)), 0, Trim(str(rsAR_T!nDebit))) & ", " & _
                                 "  " & IIf(IsNull(str(rsAR_T!nCredit)), 0, Trim(str(rsAR_T!nCredit))) & " " & _
                            ")"
                oDBFConn.Execute (lcCmdInsertAR_T)
                rsAR_T.MoveNext
            Loop
        End If
    End If
    
    Set rsAR = Nothing
    Set rsAR_T = Nothing
    DownloadAR = llOK
End Function

'PO [Warehouse Receiving (WRR)]
Public Function DownloadPO(tdteFrom As Date, tdteTo As Date) As Boolean
    Dim llOK As Boolean, lcCmdString As String
    Dim lcCmdSelectPO As String, lcCmdSelectPO_T As String
    Dim lcCmdInsertPO As String, lcCmdInsertPO_T As String
    Dim rsPO As ADODB.Recordset
    Dim rsPO_T As ADODB.Recordset
        
    lcCmdSelectPO = "SELECT A.cPONo AS cInvNo, A.dDate, A.cSupCode AS cCode, B.cName, 'Purchase' AS cType, " & _
        "CASE A.cType " & _
        "WHEN 'I' THEN 'Imported' " & _
        "WHEN 'L' THEN 'Local' END AS cPurType " & _
        "FROM PO A LEFT OUTER JOIN Supplier B ON A.cSupCode = B.cCode " & _
        "WHERE A.dDate BETWEEN '01-01-2001' AND '05-01-2003' " & _
        "AND (A.cSupCode IS NOT NULL AND RTRIM(LTRIM(A.cSupCode)) <> '') " & _
        "ORDER BY A.dDate"

    lcCmdSelectPO_T = "SELECT A.cPONo AS cInvNo, 'WHSE1' AS cWH, ISNULL(A.cItemNo_Old, A.cItemNo) AS cItemNo, A.cDesc, A.nQty, A.nPrice, ISNULL(A.cUnit, 'KGS') AS cUnit " & _
        "FROM PO_T A " & _
        "WHERE A.cPONo IN " & _
        "(SELECT cPONo FROM PO WHERE dDate BETWEEN '01-01-2001' AND " & _
        "'05-01-2003' AND (cSupCode IS NOT NULL AND " & _
        "RTRIM(LTRIM(cSupCode)) <> '')) " & _
        "ORDER BY A.cPONo"

    llOK = CreateTable("WRR")
    If llOK Then
        llOK = CreateTable("WRR_T")
    End If
    
    If llOK Then
        Set rsPO = oSQLConn.Execute(lcCmdSelectPO)
        If Not rsPO Is Nothing Then
            Do While Not rsPO.EOF
                lcCmdInsertPO = "INSERT INTO WRR " & _
                    "(cInvNo, dDate, cCode, cName, cType, cPurType) " & _
                    "VALUES (" & " '" & Trim(rsPO!cInvNo) & "', " & _
                                "'" & CStr(rsPO!dDate) & "', " & _
                                "'" & Trim(rsPO!cCode) & "', " & _
                                "'" & Replace(Trim(rsPO!cName), "'", " ") & "', " & _
                                "'" & Trim(rsPO!cType) & "', " & _
                                "'" & Trim(rsPO!cPurType) & "' " & _
                            ")"
                oDBFConn.Execute (lcCmdInsertPO)
                rsPO.MoveNext
            Loop
        End If
              
        Set rsPO_T = oSQLConn.Execute(lcCmdSelectPO_T)
        If Not rsPO_T Is Nothing Then
            Do While Not rsPO_T.EOF
                lcCmdInsertPO_T = "INSERT INTO WRR_T " & _
                    "(cInvNo, cWH, cItemNo, cDesc, nQty, nPrice, cUnit) " & _
                    "VALUES (" & "'" & Trim(rsPO_T!cInvNo) & "', " & _
                                 "'" & Trim(rsPO_T!cWH) & "', " & _
                                 "'" & Trim(rsPO_T!cItemNo) & "', " & _
                                 "'" & Replace(Trim(rsPO_T!cDesc), "'", " ") & "', " & _
                                 "  " & Trim(str(rsPO_T!nQty)) & ", " & _
                                 "  " & Trim(str(rsPO_T!nPrice)) & ", " & _
                                 "'" & CStr(rsPO_T!cUnit) & "' " & _
                            ")"
                oDBFConn.Execute (lcCmdInsertPO_T)
                rsPO_T.MoveNext
            Loop
        End If
    End If
    
    Set rsPO = Nothing
    Set rsPO_T = Nothing
    DownloadPO = llOK
End Function

'RSALES [Return Sales]
Public Function DownloadRSales(tdteFrom As Date, tdteTo As Date) As Boolean
    Dim llOK As Boolean, lcCmdString As String
    Dim lcCmdSelectRSales As String, lcCmdSelectRSales_T As String
    Dim lcCmdInsertRSales As String, lcCmdInsertRSales_T As String
    Dim rsRSales As ADODB.Recordset
    Dim rsRSales_T As ADODB.Recordset
        
    lcCmdSelectRSales = "SELECT A.cInvNo, A.dDate, A.cCode, B.cName, A.cRefInv AS cRefNo, C.cSM AS cSMan " & _
        "FROM RSales A LEFT OUTER JOIN Customer B ON A.cCode = B.cCode " & _
        "LEFT OUTER JOIN SALES C ON A.cRefInv = C.cInvNo " & _
        "WHERE A.dDate BETWEEN '" & CStr(tdteFrom) & "' AND '" & CStr(tdteTo) & "' " & _
        " AND (A.cCode IS NOT NULL AND RTRIM(LTRIM(A.cCode)) <> '') AND A.lPosted = 1" & _
        " ORDER BY A.dDate "
        
    lcCmdSelectRSales_T = "SELECT A.cInvNo, 'WHSE1' AS cWH, ISNULL(E.cItemNo_Old, A.cItemNo) AS cItemNo, A.cDesc, A.nQty, A.nPrice, F.cUnit " & _
        "FROM RSALES_T A " & _
        "LEFT OUTER JOIN RSALES R ON A.cInvNo = R.cInvNo " & _
        "LEFT OUTER JOIN SALES B ON R.cRefInv = B.cInvNo " & _
        "LEFT OUTER JOIN DR C ON B.cDRNo = C.cDRNo " & _
        "LEFT OUTER JOIN ATW_T D ON C.cAWNo = D.cAWNo " & _
        "LEFT OUTER JOIN PO_T E ON D.cPONo = E.cPONo " & _
        "LEFT OUTER JOIN ITEM F ON A.cItemNo = F.cCode " & _
        "WHERE (A.cInvNo IN (SELECT cInvNO FROM RSALES " & _
        "WHERE dDate BETWEEN '" & CStr(tdteFrom) & "' AND " & _
        "'" & CStr(tdteTo) & "' AND (cCode IS NOT NULL AND " & _
        "RTRIM(LTRIM(cCode)) <> '') AND lPosted = 1)) " & _
        "ORDER BY A.cInvNo"

    llOK = CreateTable("RSALES")
    If llOK Then
        llOK = CreateTable("RSALES_T")
    End If
    
    If llOK Then
        Set rsRSales = oSQLConn.Execute(lcCmdSelectRSales)
        If Not rsRSales Is Nothing Then
            Do While Not rsRSales.EOF
                lcCmdInsertRSales = "INSERT INTO RSALES " & _
                    "(cInvNo, dDate, cCode, cName, cRefNo, cSMan) " & _
                    "VALUES (" & " '" & Trim(rsRSales!cInvNo) & "', " & _
                                "'" & CStr(rsRSales!dDate) & "', " & _
                                "'" & Trim(rsRSales!cCode) & "', " & _
                                "'" & Replace(Trim(rsRSales!cName), "'", " ") & "', " & _
                                "'" & Trim(rsRSales!cRefNo) & "', " & _
                                "'" & Trim(rsRSales!cSMan) & "' " & _
                            ")"
                oDBFConn.Execute (lcCmdInsertRSales)
                rsRSales.MoveNext
            Loop
        End If
              
        Set rsRSales_T = oSQLConn.Execute(lcCmdSelectRSales_T)
        If Not rsRSales_T Is Nothing Then
            Do While Not rsRSales_T.EOF
                lcCmdInsertRSales_T = "INSERT INTO RSALES_T " & _
                    "(cInvNo, cWH, cItemNo, cDesc, nQty, nPrice, cUnit) " & _
                    "VALUES (" & "'" & Trim(rsRSales_T!cInvNo) & "', " & _
                                 "'" & Trim(rsRSales_T!cWH) & "', " & _
                                 "'" & Trim(rsRSales_T!cItemNo) & "', " & _
                                 "'" & Replace(Trim(rsRSales_T!cDesc), "'", " ") & "', " & _
                                 "  " & Trim(str(rsRSales_T!nQty)) & ", " & _
                                 "  " & Trim(str(rsRSales_T!nPrice)) & ", " & _
                                 "'" & CStr(rsRSales_T!cUnit) & "' " & _
                            ")"
                oDBFConn.Execute (lcCmdInsertRSales_T)
                rsRSales_T.MoveNext
            Loop
        End If
    End If
    
    Set rsRSales = Nothing
    Set rsRSales_T = Nothing
    DownloadRSales = llOK
End Function

Public Function FileExists(FileName As String) As Boolean
'The FileExists() function returns a value
'of TRUE if the specified file exists, or
'FALSE if it doesn't.
Dim Msg As String
' Turn on error trapping so error handler responds
' if any error is detected.
On Error GoTo CheckError
    FileExists = (Dir(FileName) <> "")
    ' Avoid executing error handler if no error occurs.
    Exit Function

CheckError:                 ' Branch here if error occurs.
' Define constants to represent intrinsic Visual Basic error
' codes.
Const mnErrDiskNotReady = 71, mnErrDeviceUnavailable = 68
' vbExclamation, vbOK, vbCancel, vbCritical, and vbOKCancel are
'constants defined in the VBA type library.
If (Err.Number = mnErrDiskNotReady) Then
    Msg = "Put a floppy disk in the drive and close the door."
    ' Display message box with an exclamation mark icon and with
    ' OK and Cancel buttons.
    If MsgBox(Msg, vbExclamation & vbOKCancel) = vbOK Then
        Resume
    Else
        Resume Next
    End If
ElseIf Err.Number = mnErrDeviceUnavailable Then
    Msg = "This drive or path does not exist: " & FileName
    MsgBox Msg, vbExclamation
    Resume Next
Else
    Msg = "Unexpected error #" & str(Err.Number) & " occurred: " _
    & Err.Description
    ' Display message box with Stop sign icon and OK button.
    MsgBox Msg, vbCritical
    Stop
End If
Resume
End Function

Public Sub IniFieldArr()
    ReDim arrSales(6, 1) As String
    arrSales(0, 0) = "cInvNo": arrSales(0, 1) = "CHAR(20)"
    arrSales(1, 0) = "cCode": arrSales(1, 1) = "CHAR(20)"
    arrSales(2, 0) = "cName": arrSales(2, 1) = "CHAR(100)"
    arrSales(3, 0) = "dDate": arrSales(3, 1) = "DATE"
    arrSales(4, 0) = "cSMan": arrSales(4, 1) = "CHAR(20)"
    arrSales(5, 0) = "cSMName": arrSales(5, 1) = "CHAR(100)"
    arrSales(6, 0) = "nTerm": arrSales(6, 1) = "INTEGER"
    
    ReDim arrSales_T(5, 1) As String
    arrSales_T(0, 0) = "cInvNo": arrSales_T(0, 1) = "CHAR(20)"
    arrSales_T(1, 0) = "cItemNo": arrSales_T(1, 1) = "CHAR(25)"
    arrSales_T(2, 0) = "cDesc": arrSales_T(2, 1) = "CHAR(100)"
    arrSales_T(3, 0) = "cUnit": arrSales_T(3, 1) = "CHAR(10)"
    arrSales_T(4, 0) = "nPrice": arrSales_T(4, 1) = "DOUBLE"
    arrSales_T(5, 0) = "nQty": arrSales_T(5, 1) = "DOUBLE"
    
    ReDim arrPR(3, 1) As String
    arrPR(0, 0) = "cTranNo": arrPR(0, 1) = "CHAR(20)"
    arrPR(1, 0) = "cCode": arrPR(1, 1) = "CHAR(20)"
    arrPR(2, 0) = "cName": arrPR(2, 1) = "CHAR(100)"
    arrPR(3, 0) = "dDate": arrPR(3, 1) = "DATE"
    
    ReDim arrPR_T(4, 1) As String
    arrPR_T(0, 0) = "cTranNo": arrPR_T(0, 1) = "CHAR(20)"
    arrPR_T(1, 0) = "cInvNo": arrPR_T(1, 1) = "CHAR(20)"
    arrPR_T(2, 0) = "dDate": arrPR_T(2, 1) = "DATE"
    arrPR_T(3, 0) = "nAmount": arrPR_T(3, 1) = "DOUBLE"
    arrPR_T(4, 0) = "nApplied": arrPR_T(4, 1) = "DOUBLE"
        
    ReDim arrPR_C(6, 1) As String
    arrPR_C(0, 0) = "cTranNo": arrPR_C(0, 1) = "CHAR(20)"
    arrPR_C(1, 0) = "cType": arrPR_C(1, 1) = "CHAR(5)"
    'additional field'
    arrPR_C(2, 0) = "cBank": arrPR_C(2, 1) = "CHAR(30)"
    'additional field'
    arrPR_C(3, 0) = "cCheckNo": arrPR_C(3, 1) = "CHAR(30)"
    arrPR_C(4, 0) = "dCheckDate": arrPR_C(4, 1) = "DATE"
    arrPR_C(5, 0) = "nAmount": arrPR_C(5, 1) = "DOUBLE"
    arrPR_C(6, 0) = "cInvNo": arrPR_C(6, 1) = "CHAR(20)"
    
    ReDim arrAR(5, 1) As String
    arrAR(0, 0) = "cTranNo": arrAR(0, 1) = "CHAR(20)"
    arrAR(1, 0) = "cCode": arrAR(1, 1) = "CHAR(20)"
    arrAR(2, 0) = "cName": arrAR(2, 1) = "CHAR(100)"
    arrAR(3, 0) = "dDate": arrAR(3, 1) = "DATE"
    arrAR(4, 0) = "cType": arrAR(4, 1) = "CHAR(1)"
    arrAR(5, 0) = "cInvNo": arrAR(5, 1) = "CHAR(20)"
    
    ReDim arrAR_T(5, 1) As String
    arrAR_T(0, 0) = "cTranNo": arrAR_T(0, 1) = "CHAR(20)"
    arrAR_T(1, 0) = "cAcctNo": arrAR_T(1, 1) = "CHAR(20)"
    arrAR_T(2, 0) = "cTitle": arrAR_T(2, 1) = "CHAR(100)"
    arrAR_T(3, 0) = "cRemarks": arrAR_T(3, 1) = "CHAR(100)"
    arrAR_T(4, 0) = "nDebit": arrAR_T(4, 1) = "DOUBLE"
    arrAR_T(5, 0) = "nCredit": arrAR_T(5, 1) = "DOUBLE"
    
    ReDim arrPRC(7, 1) As String
    arrPRC(0, 0) = "TransNo": arrPRC(0, 1) = "CHAR(10)"
    arrPRC(1, 0) = "TDate": arrPRC(1, 1) = "DATE"
    arrPRC(2, 0) = "Code": arrPRC(2, 1) = "CHAR(6)"
    arrPRC(3, 0) = "[Name]": arrPRC(3, 1) = "CHAR(20)"
    arrPRC(4, 0) = "Type": arrPRC(4, 1) = "CHAR(3)"
    arrPRC(5, 0) = "[Check]": arrPRC(5, 1) = "CHAR(15)"
    arrPRC(6, 0) = "[Date]": arrPRC(6, 1) = "DATE"
    arrPRC(7, 0) = "Amount": arrPRC(7, 1) = "DOUBLE"

    'PO [WRR]
    ReDim arrPO(5, 1) As String
    arrPO(0, 0) = "cInvNo": arrPO(0, 1) = "CHAR(20)"
    arrPO(1, 0) = "dDate": arrPO(1, 1) = "DATE"
    arrPO(2, 0) = "cCode": arrPO(2, 1) = "CHAR(20)"
    arrPO(3, 0) = "cName": arrPO(3, 1) = "CHAR(100)"
    arrPO(4, 0) = "cType": arrPO(4, 1) = "CHAR(20)"
    arrPO(5, 0) = "cPurType": arrPO(5, 1) = "CHAR(20)"
    
    ReDim arrPO_T(6, 1) As String
    arrPO_T(0, 0) = "cInvNo": arrPO_T(0, 1) = "CHAR(20)"
    arrPO_T(1, 0) = "cWH": arrPO_T(1, 1) = "CHAR(25)"
    arrPO_T(2, 0) = "cItemNo": arrPO_T(2, 1) = "CHAR(25)"
    arrPO_T(3, 0) = "cDesc": arrPO_T(3, 1) = "CHAR(100)"
    arrPO_T(4, 0) = "nQty": arrPO_T(4, 1) = "DOUBLE"
    arrPO_T(5, 0) = "nPrice": arrPO_T(5, 1) = "DOUBLE"
    arrPO_T(6, 0) = "cUnit": arrPO_T(6, 1) = "CHAR(10)"
    
    'RSALES
    ReDim arrRSales(5, 1) As String
    arrRSales(0, 0) = "cInvNo": arrRSales(0, 1) = "CHAR(20)"
    arrRSales(1, 0) = "dDate": arrRSales(1, 1) = "DATE"
    arrRSales(2, 0) = "cCode": arrRSales(2, 1) = "CHAR(20)"
    arrRSales(3, 0) = "cName": arrRSales(3, 1) = "CHAR(100)"
    arrRSales(4, 0) = "cRefNo": arrRSales(4, 1) = "CHAR(20)"
    arrRSales(5, 0) = "cSMan": arrRSales(5, 1) = "CHAR(20)"
    
    ReDim arrRSales_T(6, 1) As String
    arrRSales_T(0, 0) = "cInvNo": arrRSales_T(0, 1) = "CHAR(20)"
    arrRSales_T(1, 0) = "cWH": arrRSales_T(1, 1) = "CHAR(25)"
    arrRSales_T(2, 0) = "cItemNo": arrRSales_T(2, 1) = "CHAR(25)"
    arrRSales_T(3, 0) = "cDesc": arrRSales_T(3, 1) = "CHAR(100)"
    arrRSales_T(4, 0) = "nQty": arrRSales_T(4, 1) = "DOUBLE"
    arrRSales_T(5, 0) = "nPrice": arrRSales_T(5, 1) = "DOUBLE"
    arrRSales_T(6, 0) = "cUnit": arrRSales_T(6, 1) = "CHAR(10)"

End Sub

Public Function CreateTable(tstrTableName As String) As Boolean
    Dim lcCmdString As String
    Dim lcDestinationPath As String
    Dim a() As String
    Dim i As Integer
    
    Select Case tstrTableName
        Case Is = "SALES"
            a = arrSales
        Case Is = "SALES_T"
            a = arrSales_T
        Case Is = "PR"
            a = arrPR
        Case Is = "PR_T"
            a = arrPR_T
        Case Is = "PR_C"
            a = arrPR_C
        Case Is = "AR"
            a = arrAR
        Case Is = "AR_T"
            a = arrAR_T
        Case Is = "PRC"
            a = arrPRC
        'PO [WRR]
        Case Is = "WRR"
            a = arrPO
        Case Is = "WRR_T"
            a = arrPO_T
        'RSALES
        Case Is = "RSALES"
            a = arrRSales
        Case Is = "RSALES_T"
            a = arrRSales_T
    End Select
    
    If Trim(strDL) = "PR_C2" Then
        lcCmdString = "CREATE TABLE PR_C ("
        tstrTableName = "PR_C"
    Else
        lcCmdString = "CREATE TABLE " & tstrTableName & " ("
    End If
    
    For i = LBound(a) To UBound(a)
        lcCmdString = lcCmdString & a(i, 0) & " " & a(i, 1) & " NULL, "
    Next i

    lcCmdString = Left(lcCmdString, Len(lcCmdString) - 2) & ")"
    
    If UCase(Trim(mstrDBFPath)) = "C:\" Then mstrDBFPath = "C:"
'    If Trim(strDL) = "PR_C2" Then
'        lcDestinationPath = mstrDBFPath & "\PR_C.DBF"
'    Else
        lcDestinationPath = mstrDBFPath & "\" & tstrTableName & ".DBF"
'    End If
    
    If (FileExists(lcDestinationPath) And Me.lOverwrite) Then
        Kill lcDestinationPath
    End If
    
    On Error GoTo ErrHandler
    Me.oDBFConn.Execute lcCmdString
    CreateTable = True
    Exit Function
ErrHandler:
    CreateTable = False
End Function


Private Sub Class_Initialize()
    Me.IniFieldArr
End Sub


