VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Opens a recordset
Public Sub OpenRecordset(rsName As ADODB.Recordset, ColumnName As String, TableName As String, _
            Optional Condition As String, Optional WithoutCOID As Boolean)
    With rsName
        If Trim(Condition) <> "" Then
            If Not WithoutCOID Then
                Condition = Condition & " AND cCompanyID = '" & Trim(COID) & "'"
            End If
        Else
            If Not WithoutCOID Then
                Condition = "WHERE cCompanyID = '" & Trim(COID) & "'"
            End If
        End If
        
        sSQL = "SELECT " & ColumnName & " FROM " & TableName & " " & Condition
        
        If .State <> adStateOpen Then
            .CursorLocation = adUseClient
            .Open sSQL, cn, adOpenStatic, adLockBatchOptimistic
            .MarshalOptions = adMarshalModifiedOnly
            .CacheSize = 50
        End If
    End With
End Sub

Public Sub OpenDualRecord(rsName As ADODB.Recordset, ColumnName As String, ColumnName2 As String, TableName As String, TableName2 As String, Optional Condition As String)
    With rsName
        sSQL = "SELECT " & ColumnName & " FROM " & TableName & " A Left Outer Join " & TableName2 & " B On " & ColumnName2 & " " & Condition
        
        If .State <> adStateOpen Then
            .CursorLocation = adUseClient
            .Open sSQL, cn, adOpenStatic, adLockBatchOptimistic
            .MarshalOptions = adMarshalModifiedOnly
            .CacheSize = 50
        End If
    End With

End Sub

Public Sub OpenShapeRecordset(rsName As ADODB.Recordset, SQLCmd As String)
        
    With rsName
        If .State <> adStateOpen Then
            .CursorLocation = adUseClient
            .Open SQLCmd, cnShape, adOpenStatic, adLockBatchOptimistic
            .MarshalOptions = adMarshalModifiedOnly
            .CacheSize = 50
        End If
    End With
    
End Sub

Public Sub SearchMode(ColumnName As String, TableName As String, Optional Condition As String)
    
End Sub

'This sub binds the controls to the recordset
Public Sub BindControls(rs As ADODB.Recordset)
    
    For Each objCtl In FrmName
        If TypeName(objCtl) = "ITGTextBox" Then
            Set objCtl.DataSource = rs
        ElseIf TypeName(objCtl) = "ITGDateBox" Then
            Set objCtl.DataSource = rs
        ElseIf TypeName(objCtl) = "ComboBox" Then
            Set objCtl.DataSource = rs
        ElseIf TypeName(objCtl) = "ITGCheckBox" Then
            Set objCtl.DataSource = rs
        ElseIf TypeName(objCtl) = "PictureBox" Then
            Set objCtl.DataSource = rs
        ElseIf TypeName(objCtl) = "Image" Then
            Set objCtl.DataSource = rs
        ElseIf TypeName(objCtl) = "RichTextBox" Then
            Set objCtl.DataSource = rs
        ElseIf TypeName(objCtl) = "ITGCombobox" Then
            Set objCtl.DataSource = rs
        End If
    Next

End Sub

'This sub unbinds the controls to the recordset
Public Sub UnbindControls()
    
    For Each objCtl In FrmName
        If TypeName(objCtl) = "ITGTextBox" Then
            Set objCtl.DataSource = Nothing
        ElseIf TypeName(objCtl) = "ITGDateBox" Then
            Set objCtl.DataSource = Nothing
        ElseIf TypeName(objCtl) = "ComboBox" Then
            Set objCtl.DataSource = Nothing
        ElseIf TypeName(objCtl) = "ITGCheckBox" Then
            Set objCtl.DataSource = Nothing
        ElseIf TypeName(objCtl) = "PictureBox" Then
            Set objCtl.DataSource = Nothing
        ElseIf TypeName(objCtl) = "Image" Then
            Set objCtl.DataSource = Nothing
        ElseIf TypeName(objCtl) = "RichTextBox" Then
            Set objCtl.DataSource = Nothing
        ElseIf TypeName(objCtl) = "ITGCombobox" Then
            Set objCtl.DataSource = Nothing
        End If
    Next

End Sub

Public Sub Close_RecordSet(cRecordset As ADODB.Recordset)
    cRecordset.Close
    Set cRecordset = Nothing
End Sub

'Opens a recordset with new refreshed/connection
Public Sub OpenRecordsetWithCN(rsName As ADODB.Recordset, ColumnName As String, TableName As String, _
            cConn As ADODB.Connection, Optional Condition As String, Optional WithoutCOID As Boolean)
    With rsName
        If Trim(Condition) <> "" Then
            If Not WithoutCOID Then
                Condition = Condition & " AND cCompanyID = '" & Trim(COID) & "'"
            End If
        Else
            If Not WithoutCOID Then
                Condition = "WHERE cCompanyID = '" & Trim(COID) & "'"
            End If
        End If
        
        sSQL = "SELECT " & ColumnName & " FROM " & TableName & " " & Condition
        
        If .State <> adStateOpen Then
            .CursorLocation = adUseClient
            .Open sSQL, cConn, adOpenStatic, adLockBatchOptimistic
            .MarshalOptions = adMarshalModifiedOnly
            .CacheSize = 50
        End If
    End With
End Sub

'For computation of actual product balance
Public Sub ProductBalanceCommand(rsName As ADODB.Recordset, ProductID As String)
'Dim cmd As New ADODB.Command
'
'
'    With cmd
'        .ActiveConnection = cn
'        .CommandTimeout = 1000
'        .CommandText = "SP_ProductBalanceInquiry_Module"
'        .CommandType = adCmdStoredProc
'        .Parameters("@cItemNo") = ProductID
'        .Parameters("@cCompanyID") = COID
'    End With
'
'    Set rsName = cmd.Execute()
'
End Sub

'For computation of actual NNP Balance
Public Sub NNPBalanceCommand(rsName As ADODB.Recordset, NNPID As String)
Dim cmd As New ADODB.Command
    
    With cmd
        .ActiveConnection = cn
        .CommandTimeout = 1000
        .CommandText = "SP_NNPBalanceInquiry_Module"
        .CommandType = adCmdStoredProc
        .Parameters("@cNNPNo") = NNPID
        .Parameters("@cCompanyID") = COID
    End With
    
    Set rsName = cmd.Execute()

End Sub

'For computation of credit limit per client
Public Sub CreditLimitCommand(rsName As ADODB.Recordset, CustomerID As String, MSCode As String)
Dim cmd As New ADODB.Command
    
    With cmd
        .ActiveConnection = cn
        .CommandTimeout = 1000
        .CommandText = "SP_CG"
        .CommandType = adCmdStoredProc
        .Parameters("@cCode") = CustomerID
        .Parameters("@cMSCode") = MSCode
        .Parameters("@cCompanyID") = COID
    End With
    
    Set rsName = cmd.Execute()

End Sub

'For computation of credit limit per parent customer (group code)
Public Sub CreditLimitPCCommand(rsName As ADODB.Recordset, GroupCode As String, ClassCode As String)
Dim cmd As New ADODB.Command
    
    
    With cmd
        .ActiveConnection = cn
        .CommandTimeout = 1000
        .CommandText = "SP_CG_PC"
        .CommandType = adCmdStoredProc
        .Parameters("@cGroupCode") = GroupCode
        .Parameters("@cClassCode") = ClassCode
        .Parameters("@cCompanyID") = COID
    End With
    
    Set rsName = cmd.Execute()

End Sub

Public Sub ProductAgeing(rsName As ADODB.Recordset, ProductID As String, Interval As Integer)
Dim cmd As New ADODB.Command
    With cmd
        .ActiveConnection = cn
        .CommandTimeout = 1000
        .CommandText = "rsp_Inventory_Product_Aging"
        .CommandType = adCmdStoredProc
        .Parameters("@cCompName") = ""
        .Parameters("@cAddress1") = ""
        .Parameters("@cAddress2") = ""
        .Parameters("@cReportTitle") = ""
        .Parameters("@cCriteria1") = ""
        .Parameters("@cCriteria") = ""
        .Parameters("@cFrom") = ProductID
        .Parameters("@cTo") = ProductID
        .Parameters("@dDateFROM") = Now - Interval
        .Parameters("@dDateTo") = Now
        .Parameters("@cCompID") = COID
        .Parameters("@nInterval") = Interval
        .Parameters("@cValue1") = ""
        .Parameters("@cValue2") = ""
        .Parameters("@cValue3") = ""
        .Parameters("@cValue4") = ""
        .Parameters("@cValue5") = ""
        .Parameters("@cValue6") = ""
    End With
    Set rsName = cmd.Execute()
End Sub




