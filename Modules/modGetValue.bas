Attribute VB_Name = "modGetValue"
Option Explicit

'Gets value from the database table
Public Function GetValueFrTable(ReturnFld As String, FromTable As String, Criteria As String, _
            Optional lNoCompany As Boolean)
Dim varRetVal As Variant
Dim rsGet As New Recordset
Dim cComp As String

    If Not lNoCompany Then
        cComp = " AND cCompanyID = '" & Trim(COID) & "'"
    Else
        cComp = ""
    End If
        
    sSQL = "SELECT " & Trim(ReturnFld) _
         & " FROM " & Trim(FromTable) _
         & " WHERE " & Trim(Criteria) & " " & cComp
         
    rsGet.Open sSQL, cn, adOpenKeyset, adLockReadOnly

    While Not rsGet.EOF
        varRetVal = rsGet(Trim(ReturnFld))
        rsGet.MoveNext
    Wend
    
    If IsNull(varRetVal) Then
        GetValueFrTable = ""
    Else
        GetValueFrTable = varRetVal
    End If
    
    rsGet.Close
    Set rsGet = Nothing
    

End Function

'Get complete address from the database
Public Function GetAddress(FromTable As String, Criteria As String)
Dim varRetVal As Variant
Dim rsGet As New Recordset

    sSQL = "SELECT cAddress, cCity, cCountry" & _
            " FROM " & Trim(FromTable) & " " & _
            "WHERE " & Trim(Criteria) & " AND cCompanyID = '" & Trim(COID) & "'"

    rsGet.Open sSQL, cn, adOpenKeyset

    While Not rsGet.EOF
        If IsNull(rsGet!cAddress) = False Then varRetVal = rsGet!cAddress
        If IsNull(rsGet!cCity) = False Then varRetVal = varRetVal & ", " & rsGet!cCity
        If IsNull(rsGet!cCountry) = False Then varRetVal = varRetVal & ", " & rsGet!cCountry
        rsGet.MoveNext
    Wend
    
    If IsNull(varRetVal) Then
        GetAddress = ""
    Else
        GetAddress = varRetVal
    End If
    
    rsGet.Close
    Set rsGet = Nothing
    
End Function

'Load values into a combo box control
Public Sub LoadComboValues(ByRef objCombo As ComboBox, ByVal FieldName As String, _
                            ByVal TableName As String, Optional Cond As String, _
                            Optional Order As String)
Dim cnCombo As New ADODB.Connection
Dim strOldValue As String
Dim rs As New Recordset

    cnCombo.CursorLocation = adUseClient
    
    cnCombo.ConnectionString = "driver={" & sDBDriver & "};" & _
    "server=" & sServer & ";uid=sa;pwd=" & sDBPassword & ";database=" & sDBname

    cnCombo.Open
    
    If Trim(Order) = "" Then
        sSQL = "SELECT DISTINCT " & FieldName & " FROM " & TableName & " " & Cond
    Else
        sSQL = "SELECT DISTINCT " & FieldName & " FROM " & TableName & " " & Cond & " ORDER BY " & Order
    End If
    
    rs.Open sSQL, cnCombo, adOpenKeyset

    'strOldValue = objCombo

    objCombo.Clear
    Do While Not rs.EOF
        If Trim("" & rs(FieldName)) <> "" Then
            objCombo.AddItem Trim("" & rs(FieldName))
        End If
        rs.MoveNext
    Loop
    rs.Close
    cnCombo.Close
    Set rs = Nothing
    Set cnCombo = Nothing

End Sub

'Load values into a combo box control
Public Sub LoadComboValues1(ByRef objCombo As ITGCombobox, ByVal FieldName As String, _
                            ByVal TableName As String, Optional Cond As String, _
                            Optional Order As String)
Dim cnCombo As New ADODB.Connection
Dim strOldValue As String
Dim rs As New Recordset

    cnCombo.CursorLocation = adUseClient
    
    cnCombo.ConnectionString = "driver={" & sDBDriver & "};" & _
    "server=" & sServer & ";uid=sa;pwd=" & sDBPassword & ";database=" & sDBname

    cnCombo.Open
    
    If Trim(Order) = "" Then
        sSQL = "SELECT DISTINCT " & FieldName & " FROM " & TableName & " " & Cond
    Else
        sSQL = "SELECT DISTINCT " & FieldName & ", " & Order & " FROM " & TableName & " " & Cond & " ORDER BY " & Order
    End If
    
    rs.Open sSQL, cnCombo, adOpenKeyset

    'strOldValue = objCombo

    'objCombo.Clear
    Do While Not rs.EOF
        If Trim("" & rs(FieldName)) <> "" Then
            objCombo.AddItem Trim("" & rs(FieldName))
        End If
        rs.MoveNext
    Loop
    rs.Close
    cnCombo.Close
    Set rs = Nothing
    Set cnCombo = Nothing

End Sub


'Load values into a combo box control
Public Sub LoadUnitValues(ByRef objCombo As ComboBox, ByVal ItemNo As String)
Dim rs As New Recordset

    sSQL = "SELECT * FROM ( " & _
            "SELECT cUnit FROM ITEM WHERE cItemNo = '" & ItemNo & "' AND cCompanyID = '" & COID & "' " & _
            "UNION ALL " & _
            "SELECT cUnit FROM ITEM_UNIT WHERE cItemNo = '" & ItemNo & "' AND cCompanyID = '" & COID & "' " & _
            ") A ORDER BY cUnit "
    rs.Open sSQL, cn, adOpenKeyset

    objCombo.Clear
    Do While Not rs.EOF
        If Trim("" & rs!cUnit) <> "" Then
            objCombo.AddItem Trim("" & rs!cUnit)
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

End Sub

'Load values into an Option table
Public Function LoadOption(ByVal param1 As String, ByVal param2 As Integer) As String
Dim rs As New Recordset
    If param2 = 3 Then
        sSQL = "SELECT lBit FROM SYSTEM_OPTION where cCode='" & param1 & "' AND cCompanyID = '" & COID & "'"
    Else
        sSQL = "SELECT cValue FROM SYSTEM_OPTION where cCode='" & param1 & "' AND cCompanyID = '" & COID & "'"
    End If
    rs.Open sSQL, cn, adOpenKeyset
    If rs.RecordCount <> 0 Then
        LoadOption = rs(0)
    End If
    rs.Close
    Set rs = Nothing
End Function

'Field existing on database and/or recordset
Public Function IDExisting(rs As ADODB.Recordset, Field As String, _
        Table As String, Filter As String, Optional Condition As String, Optional lNoCompany As Boolean) As Boolean
Dim rsExist As ADODB.Recordset
Dim strID
    
    'sample: IDExisting(rsHeader, "cCode", "CUSTOMER", Trim(rsHeader!cCode))
    IDExisting = False
    
    If Trim(Condition) <> "" Then
        Condition = " AND " & Condition
    End If
        
    strID = GetValueFrTable(Field, Table, Field & " = '" & Filter & "' " & Condition & "", lNoCompany)
    If strID <> "" Then IDExisting = True
    
    If IDExisting Then Exit Function
    
    Set rsExist = New ADODB.Recordset
    Set rsExist = rs.Clone
    
    If Condition <> "" Then
        rsExist.Filter = Field & " = '" & Filter & "' " & Condition
    Else
        rsExist.Filter = Field & " = '" & Filter & "'"
    End If
    
    If rsExist.RecordCount > 1 Then
        IDExisting = True
    End If
    
    rsExist.Close
    Set rsExist = Nothing
    
End Function

'Field existing on recordset
Public Function ExistingOnRS(rs As ADODB.Recordset, Field As String, _
        Filter As String) As Boolean
Dim rsExist As ADODB.Recordset
Dim strID
    
    ExistingOnRS = False
    
    'sample: IDExisting(rsHeader, "cClassCode", Trim(rsHeader!cClassCode))
    Set rsExist = New ADODB.Recordset
    Set rsExist = rs.Clone
    rsExist.Filter = Field & " = '" & Filter & "'"
    
    If rsExist.RecordCount > 1 Then
        ExistingOnRS = True
    End If
    
    rsExist.Close
    Set rsExist = Nothing
    
End Function

'Numeric to words (for currency use only)
Public Function NumToWords(tnNumber As Currency, tcCurrency As String) As String
Dim lcOnes As String, lcTens As String
Dim lcWhole As String, lcFraction As String
Dim lnCtr As Long, lcParts As String
Dim laParts() As String, laOnes() As String, laTens() As String, i As Long
Dim lcCurrentVal As String, lcRetVal As String
Dim lcArg As String, lnHundred As Long, lnOnes As Long, lnTens As Long
    
    lcOnes = "Zero,One,Two,Three,Four,Five,Six,Seven,Eight,Nine," + _
             "Ten,Eleven,Twelve,Thirteen,Fourteen,Fifteen,Sixteen," + _
             "Seventeen,Eighteen,Nineteen"
    laOnes = Split(lcOnes, ",")
             
    lcTens = "Zero,Ten,Twenty-,Thirty-,Forty-,Fifty-,Sixty-,Seventy-,Eighty-,Ninety-"
    laTens = Split(lcTens, ",")
    
    lcWhole = String(12 - Len(Trim(str(Int(tnNumber)))), "0") + Trim(str(Int(tnNumber)))
    lcFraction = Trim(str((Round(tnNumber, 2) - Int(tnNumber)) * 100))
    lnCtr = 1
    lcParts = Space(0)
    For i = 1 To Len(lcWhole) Step 3
        lcParts = lcParts + Mid(lcWhole, i, 3) + ","
    Next i
    
    lcParts = Left(lcParts, Len(lcParts) - 1)
    laParts = Split(lcParts, ",")
    lcRetVal = Space(0)
    For i = 0 To UBound(laParts)
        lcArg = laParts(i)
        If Val(lcArg) > 0 Then
            lnHundred = Val(Mid(lcArg, 1, 1))
            lnTens = Val(Mid(lcArg, 2, 1))
            lnTens = IIf(lnTens = 1, Val(Mid(lcArg, 2, 2)), lnTens)
            lnOnes = IIf(lnTens > 9, 0, Val(Mid(lcArg, 3, 1)))
            If lnTens < 10 And lnTens > 0 Then
                lcCurrentVal = IIf(lnHundred = 0, "", laOnes(lnHundred) + " Hundred ") + _
                       laTens(lnTens) + IIf(lnOnes = 0, "", laOnes(lnOnes))
            Else
                lcCurrentVal = IIf(lnHundred = 0, "", laOnes(lnHundred) + " Hundred ") + _
                       IIf(lnTens = 0, "", laOnes(lnTens)) + IIf(lnOnes = 0, "", laOnes(lnOnes))
            End If
            If lnTens < 10 And lnOnes = 0 Then
                lcCurrentVal = Left(lcCurrentVal, Len(lcCurrentVal) - 1)
            End If
            
            Select Case i
                Case 2
                    lcCurrentVal = lcCurrentVal + " Thousand "
                Case 1
                    lcCurrentVal = lcCurrentVal + " Million "
                Case 0
                    lcCurrentVal = lcCurrentVal + " Billion "
            End Select
            
            lcRetVal = lcRetVal + lcCurrentVal
        End If
    Next i
    lcRetVal = lcRetVal + IIf(lcFraction = "0", " " & Trim(tcCurrency) & " only", _
        IIf(Val(lcWhole) > 0, " " & Trim(tcCurrency) & " and " + _
        lcFraction + "/100 only", lcFraction + "/100 only"))
    NumToWords = lcRetVal
End Function

'Encrypting string
Public Function Encrypt(str As String) As String
Dim intCntr, intKey As Integer
Dim strTemp As String
Dim tmp As String
  
    intKey = Len(str) - 1
    strTemp = ""
    For intCntr = 0 To Len(str) - 1
        tmp = Chr(Asc(Left(Right(str, Len(str) - intCntr), 1)) - (intKey + intCntr))
        If tmp = "'" Then
            tmp = "*"
        End If
        strTemp = strTemp & tmp
        intKey = intKey + 3
    Next intCntr
    Encrypt = strTemp & IIf(Chr(intKey) <> "'", Chr(intKey), "*")
    
End Function

'Decrypting string
Public Function Decrypt(str As String) As String
Dim intCntr, intKey As Integer
Dim strTemp As String
  
    intKey = Len(str) - 2
    strTemp = ""
    For intCntr = 0 To Len(str) - 2
       If Left(Right(str, Len(str) - intCntr), 1) = "*" Then
            strTemp = strTemp & Chr(Asc("'") + (intKey + intCntr))
       
       Else
            strTemp = strTemp & Chr(Asc(Left(Right(str, Len(str) - intCntr), 1)) + (intKey + intCntr))
              
       End If
    '==============
        'strTemp = strTemp & Chr(Asc(Left(Right(str, Len(str) - intCntr), 1)) + (intKey + intCntr))
        intKey = intKey + 3
    
    Next intCntr
    Decrypt = strTemp
    
End Function

Public Function GetExpenseAmount(ProjectID As String) As Double
On Error GoTo ErrHandler
Dim rsGetExp As New ADODB.Recordset
Set rsGetExp = New ADODB.Recordset
rsGetExp.CursorLocation = adUseClient
rsGetExp.Open " exec [rsp_Expenses_Project] '" & COID & "', '" & Mid(ProjectID, 1, 8) & "' ", cn, 3, 3
If Not rsGetExp.EOF Then
    GetExpenseAmount = rsGetExp!ExpenseAmount
End If
Set rsGetExp = Nothing
ErrHandler:
If err.Number = "-2147217900" Then
    MsgBox "Stored Procedured 'rsp_Expenses_Project' not Found!", vbInformation, "Error"
End If
End Function

Public Function GetServiceFeeAmount(ProjectID As String) As Double
On Error GoTo ErrHandler
Dim rsGetExp As New ADODB.Recordset
Set rsGetExp = New ADODB.Recordset
rsGetExp.CursorLocation = adUseClient
rsGetExp.Open " exec [rsp_Project_ServiceAmount] '" & COID & "', '" & ProjectID & "' ", cn, 3, 3
If Not rsGetExp.EOF Then
    GetServiceFeeAmount = IIf(IsNull(rsGetExp!nAmount), 0, rsGetExp!nAmount)
End If
Set rsGetExp = Nothing
ErrHandler:
If err.Number = "-2147217900" Then
    MsgBox "Stored Procedured 'rsp_Project_ServiceAmount' not Found!", vbInformation, "Error"
End If
End Function


Public Function GetProbableServiceAmount(ProjectID As String) As Double
GetProbableServiceAmount = GetValueFrTable("nRemuneration", "PMS_LEAD", "cLeadID = '" & Mid(ProjectID, 1, 8) & "'")
GetProbableServiceAmount = GetProbableServiceAmount * 0.02
End Function

