Attribute VB_Name = "modControlNumber"
Option Explicit

Public Function GetAutoCtrlNo(rs As ADODB.Recordset, sRefCode As String, _
                sTblName As String, sFldName As String) As String
                
Dim lCtrlNoOk As Boolean
Dim sAutoNo As String
    sAutoNo = Trim(GetValueFrTable("cValue", "SYSTEM_OPTION", " cCode='" & Trim(sRefCode) & "'"))
    
    Do Until lCtrlNoOk
    If IDExisting(rs, sFldName, sTblName, Left(sAutoNo, 5) + Format(CStr((CInt(Right(sAutoNo, 4)) + 1)), "0000")) Then
        sAutoNo = Left(sAutoNo, 5) + Format(CStr((CInt(Right(sAutoNo, 4)) + 1)), "0000")
    Else
        lCtrlNoOk = True
    End If
    Loop
    
    lCtrlNoOk = False
    
    If Mid(sAutoNo, 4, 2) <> Format(Date, "yy") Then
        sAutoNo = Left(sAutoNo, 3) + Format(Date, "yy") + "0000"
    Else
        sAutoNo = Left(sAutoNo, 3) + Format(Date, "yy") + Format(Right(sAutoNo, 4) + 1, "0000")
    End If
        
    Do Until lCtrlNoOk
        If IDExisting(rs, sFldName, sTblName, sAutoNo) Then
            sAutoNo = Left(sAutoNo, 5) + Format(CStr((CInt(Right(sAutoNo, 4)) + 1)), "0000")
        Else
            lCtrlNoOk = True
        End If
    Loop
    
    Select Case sRefCode
    
        Case "AUTO_NUMBER_ARCM", "AUTO_NUMBER_RBF", "AUTO_NUMBER_PO", "AUTO_NUMBER_OR"
            sAutoNo = Left(sAutoNo, 3) + Format(CStr((CInt(Right(sAutoNo, 4)) + 1)), "000000")
    End Select
    
    

    
    GetAutoCtrlNo = sAutoNo
    FrmName.strNo = sAutoNo
    
End Function


Public Function ChkCtrlNo(rs As ADODB.Recordset, sRefCode As String, _
                sTblName As String, sFldName As String, sTranNo As String) As String
Dim lCtrlNoOk As Boolean
Dim strNewNo As String
    
    If IDExisting(rs, sFldName, sTblName, sTranNo) Then
        
        strNewNo = Trim(GetValueFrTable("cValue", "SYSTEM_OPTION", " cCode='" & Trim(sRefCode) & "'"))
        
        If Mid(strNewNo, 4, 2) <> Format(Date, "yy") Then
            strNewNo = Left(strNewNo, 3) + Format(Date, "yy") + "0000"
        Else
            strNewNo = Left(strNewNo, 3) + Format(Date, "yy") + Format(Right(strNewNo, 4) + 1, "0000")
        End If
        
        Do Until lCtrlNoOk
            If IDExisting(rs, sFldName, sTblName, strNewNo) Then
                strNewNo = Left(strNewNo, 5) + Format(CStr((CInt(Right(strNewNo, 4)) + 1)), "0000")
            Else
                lCtrlNoOk = True
            End If
        Loop
        
        ChkCtrlNo = strNewNo
        FrmName.strNo = strNewNo
    Else
        ChkCtrlNo = sTranNo
    End If
    
End Function


'Update system generated number
Public Sub UpdateControlNo(sAutoNo As String, sTranNo As String, sRefCode As String)
On Error GoTo TheSource
Dim sCtrlNo As String

    sCtrlNo = Trim(GetValueFrTable("cValue", "SYSTEM_OPTION", " cCode='" & Trim(sRefCode) & "'"))
    
    If Left(sCtrlNo, 3) + Format(Date, "yy") + "0000" = Trim(sTranNo) Then
        sSQL = "UPDATE SYSTEM_OPTION SET cValue = '" & Trim(sTranNo) & "' WHERE cCode = '" & Trim(sRefCode) & "' AND cCompanyID = '" & COID & "'"
        cn.Execute sSQL
        Exit Sub
    End If
    
    If (Left(sTranNo, 5) & (Format(Right(sTranNo, 4), "0000"))) <> Trim(sAutoNo) Then Exit Sub
    sSQL = "UPDATE SYSTEM_OPTION SET cValue = '" & Trim(sTranNo) & "' WHERE cCode = '" & Trim(sRefCode) & "' AND cCompanyID = '" & COID & "'"
    cn.Execute sSQL

TheSource:
    Exit Sub
End Sub


'=====
'
'For NNP and NPP only
'
'=====

Public Function GetAutoCtrlNo1(rs As ADODB.Recordset, sRefCode As String, _
                sTblName As String, sFldName As String) As String
                
Dim lCtrlNoOk As Boolean
Dim sAutoNo As String
    sAutoNo = Trim(GetValueFrTable("cValue", "SYSTEM_OPTION", " cCode='" & Trim(sRefCode) & "'"))
    
    Do Until lCtrlNoOk
        If IDExisting(rs, sFldName, sTblName, sAutoNo) Then
            sAutoNo = Left(sAutoNo, 5) + Format(CStr((CInt(Mid(sAutoNo, 6, 3)) - 1)), "000")
        Else
            lCtrlNoOk = True
        End If
    Loop
    
    lCtrlNoOk = False
    
    If Mid(sAutoNo, 1, 4) <> Format(Date, "yyyy") Then
        sAutoNo = Format(Date, "yyyy") + "-999"
    Else
        'sAutoNo = Format(Date, "yyyy") + "-" + Format(Mid(sAutoNo, 6, 3) + 1, "000")
    End If
        
    Do Until lCtrlNoOk
        If IDExisting(rs, sFldName, sTblName, sAutoNo) Then
            sAutoNo = Left(sAutoNo, 5) + Format(CStr((CInt(Mid(sAutoNo, 6, 3)) - 1)), "000")
        Else
            lCtrlNoOk = True
        End If
    Loop
    
    GetAutoCtrlNo1 = sAutoNo
    FrmName.strNo = sAutoNo
    
End Function


Public Function ChkCtrlNo1(rs As ADODB.Recordset, sRefCode As String, _
                sTblName As String, sFldName As String, sTranNo As String) As String
Dim lCtrlNoOk As Boolean
Dim strNewNo As String
    
    If IDExisting(rs, sFldName, sTblName, sTranNo) Then
        
        strNewNo = Trim(GetValueFrTable("cValue", "SYSTEM_OPTION", " cCode='" & Trim(sRefCode) & "'"))
        
        If Mid(strNewNo, 1, 4) <> Format(Date, "yyyy") Then
            strNewNo = Format(Date, "yyyy") + "-000"
        Else
            strNewNo = Format(Date, "yyyy") + "-" + Format(Mid(strNewNo, 6, 3) - 1, "000")
        End If
        
        Do Until lCtrlNoOk
            If IDExisting(rs, sFldName, sTblName, strNewNo) Then
                strNewNo = Left(strNewNo, 5) + Format(CStr((CInt(Mid(strNewNo, 6, 3)) - 1)), "000")
            Else
                lCtrlNoOk = True
            End If
        Loop
        
        ChkCtrlNo1 = strNewNo
        FrmName.strNo = strNewNo
    Else
        ChkCtrlNo1 = sTranNo
    End If
    
End Function


'Update system generated number
Public Sub UpdateControlNo1(sAutoNo As String, sTranNo As String, sRefCode As String)
On Error GoTo TheSource
Dim sCtrlNo As String

    sCtrlNo = Trim(GetValueFrTable("cValue", "SYSTEM_OPTION", " cCode='" & Trim(sRefCode) & "'"))
    
    If Format(Date, "yyyy") + "-000" = Trim(sTranNo) Then
        sSQL = "UPDATE SYSTEM_OPTION SET cValue = '" & Trim(sTranNo) & "' WHERE cCode = '" & Trim(sRefCode) & "' AND cCompanyID = '" & COID & "'"
        cn.Execute sSQL
        Exit Sub
    End If
    
    If (Left(sTranNo, 5) & (Format(Mid(sTranNo, 6, 3), "000"))) <> Trim(sAutoNo) Then Exit Sub
    sSQL = "UPDATE SYSTEM_OPTION SET cValue = '" & Trim(sTranNo) & "' WHERE cCode = '" & Trim(sRefCode) & "' AND cCompanyID = '" & COID & "'"
    cn.Execute sSQL

TheSource:
    Exit Sub
End Sub

'===================

Public Function GetAutoCtrlNo2(rs As ADODB.Recordset, sRefCode As String, _
                sTblName As String, sFldName As String) As String
                
Dim lCtrlNoOk As Boolean
Dim sAutoNo As String
    sAutoNo = Trim(GetValueFrTable("cValue", "SYSTEM_OPTION", " cCode='" & Trim(sRefCode) & "'"))
    
    Do Until lCtrlNoOk
        If IDExisting(rs, sFldName, sTblName, sAutoNo) Then
            sAutoNo = Left(sAutoNo, 5) + Format(CStr((CInt(Mid(sAutoNo, 6, 3)) + 1)), "000")
        Else
            lCtrlNoOk = True
        End If
    Loop
    
    lCtrlNoOk = False
    
    If Mid(sAutoNo, 1, 4) <> Format(Date, "yyyy") Then
        sAutoNo = Format(Date, "yyyy") + "-001"
    Else
        'sAutoNo = Format(Date, "yyyy") + "-" + Format(Mid(sAutoNo, 6, 3) + 1, "000")
    End If
        
    Do Until lCtrlNoOk
        If IDExisting(rs, sFldName, sTblName, sAutoNo) Then
            sAutoNo = Left(sAutoNo, 5) + Format(CStr((CInt(Mid(sAutoNo, 6, 3)) + 1)), "000")
        Else
            lCtrlNoOk = True
        End If
    Loop
    
    GetAutoCtrlNo2 = sAutoNo
    FrmName.strNo = sAutoNo
    
End Function


Public Function ChkCtrlNo2(rs As ADODB.Recordset, sRefCode As String, _
                sTblName As String, sFldName As String, sTranNo As String) As String
Dim lCtrlNoOk As Boolean
Dim strNewNo As String
    
    If IDExisting(rs, sFldName, sTblName, sTranNo) Then
        
        strNewNo = Trim(GetValueFrTable("cValue", "SYSTEM_OPTION", " cCode='" & Trim(sRefCode) & "'"))
        
        If Mid(strNewNo, 1, 4) <> Format(Date, "yyyy") Then
            strNewNo = Format(Date, "yyyy") + "-000"
        Else
            strNewNo = Format(Date, "yyyy") + "-" + Format(Mid(strNewNo, 6, 3) - 1, "000")
        End If
        
        Do Until lCtrlNoOk
            If IDExisting(rs, sFldName, sTblName, strNewNo) Then
                strNewNo = Left(strNewNo, 5) + Format(CStr((CInt(Mid(strNewNo, 6, 3)) - 1)), "000")
            Else
                lCtrlNoOk = True
            End If
        Loop
        
        ChkCtrlNo2 = strNewNo
        FrmName.strNo = strNewNo
    Else
        ChkCtrlNo2 = sTranNo
    End If
    
End Function


'Update system generated number
Public Sub UpdateControlNo2(sAutoNo As String, sTranNo As String, sRefCode As String)
On Error GoTo TheSource
Dim sCtrlNo As String

    
    sCtrlNo = Trim(GetValueFrTable("cValue", "SYSTEM_OPTION", " cCode='" & Trim(sRefCode) & "'"))
    
    If Format(Date, "yyyy") + "-000" = Trim(sTranNo) Then
        sSQL = "UPDATE SYSTEM_OPTION SET cValue = '" & Trim(sTranNo) & "' WHERE cCode = '" & Trim(sRefCode) & "' AND cCompanyID = '" & COID & "'"
        cn.Execute sSQL
        Exit Sub
    End If
    
    If (Left(sTranNo, 5) & (Format(Mid(sTranNo, 6, 3), "000"))) <> Trim(sAutoNo) Then Exit Sub
    sSQL = "UPDATE SYSTEM_OPTION SET cValue = '" & Trim(sTranNo) & "' WHERE cCode = '" & Trim(sRefCode) & "' AND cCompanyID = '" & COID & "'"
    cn.Execute sSQL

TheSource:
    Exit Sub
End Sub


'-new codes

Public Function Generate_AutoNumber(cCode As String, Optional cType As String = "ALL", Optional txtbx As ITGTextBox) As String
    
    Dim GenNum As String
    
    Dim sPrefix As String
    Dim sDay As String
    Dim sMonth As String
    Dim sYear As String
    Dim sSeparator As String
    Dim sNumeric As String
    Dim sSuffix As String

    
    GetConst_AutoNum_Param sPrefix, sDay, sMonth, sYear, sSeparator, sNumeric, sSuffix
    
    On Error Resume Next
    If rs.State = adStateOpen Then rs.Close
    sSQL = "select * from AUTONUM where cModuleID ='" & cCode & "' and cType ='" & cType & "' and lInactive=0 order by cType Desc"
    rs.Open sSQL, cn, adOpenStatic, adLockOptimistic
    If rs.RecordCount = 0 Then
        If rs.State = adStateOpen Then rs.Close
        sSQL = "select * from AUTONUM where cModuleID ='" & cCode & "' and cType ='ALL' and lInactive=0 order by cType Desc"
        rs.Open sSQL, cn, adOpenStatic, adLockOptimistic
        If rs.RecordCount = 0 Then
            Generate_AutoNumber = ""
            txtbx.Locked = False
            Exit Function
        End If
    End If
        
    Update_AutoNumber cCode, cType, gblQty
        
    'reset numbering when current year is not updated
    If Year(Now) <> CurrentYear Or Month(Now) <> CurrentMonth Then
        rs!nValue = 1
        cn.Execute "update autonum set nValue=1"
    End If
    nCurrentCounter = rs!nValue
        
    GenNum = rs!cFormat
    
    'New Code - To resolve AutoNum concern
    AutonumFormat = GenNum
    gblsNumeric = sNumeric
    nNumStart = InStr(1, AutonumFormat, sNumeric)
    nNumLen = Len(AutonumFormat) - nNumStart
    
    
    txtbx.Locked = rs!lLocked
    txtbx.Font.Bold = rs!lBold
    txtbx.Visible = rs!lVisible
    
    Dim i, j As Integer
    Dim k, l As String
    'Day
    i = InStr(1, GenNum, sDay)
    j = InStrRev(GenNum, sDay)
    If (i = 0 Or j = 0) Then
        k = ""
    Else
        k = Mid(GenNum, i, j - i + 1)
        l = Format(Right(Day(Now), Len(k)), Replace(k, sDay, "0"))
        GenNum = Replace(GenNum, k, l)
    End If
    
    'Month
    i = InStr(1, GenNum, sMonth)
    j = InStrRev(GenNum, sMonth)
    If (i = 0 Or j = 0) Then
        k = ""
    Else
        k = Mid(GenNum, i, j - i + 1)
        l = Format(Right(Month(Now), Len(k)), Replace(k, sMonth, "0"))
        GenNum = Replace(GenNum, k, l)
    End If
    
    'Year
    i = InStr(1, GenNum, sYear)
    j = InStrRev(GenNum, sYear)
    If (i = 0 Or j = 0) Then
        k = ""
    Else
        k = Mid(GenNum, i, j - i + 1)
        l = Format(Right(Year(Now), Len(k)), Replace(k, sYear, "0"))
        GenNum = Replace(GenNum, k, l)
    End If
    
    'Prefix
    i = InStr(1, GenNum, sPrefix)
    j = InStrRev(GenNum, sPrefix)
    If (i = 0 Or j = 0) Then
        k = ""
    Else
        k = Mid(GenNum, i, j - i + 1)
        GenNum = Replace(GenNum, k, IIf(rs!cPrefix = ".", SecTerminalID, rs!cPrefix)) 'rs!cPrefix)
    End If
    
    'Numeric
    i = InStr(1, GenNum, sNumeric)
    j = InStrRev(GenNum, sNumeric)
    If (i = 0 Or j = 0) Then
        k = ""
    Else
        k = Mid(GenNum, i, j - i + 1)
        l = Format(rs!nValue, Replace(Mid(GenNum, i, j - i + 1), sNumeric, "0"))
        GenNum = Replace(GenNum, k, l)
    End If
    
    'Suffix
    If rs!cSuffix <> "" Then
        i = InStr(1, GenNum, sSuffix)
        j = InStrRev(GenNum, sSuffix)
        If (i = 0 Or j = 0) Then
            k = ""
        Else
            k = Mid(GenNum, i, j - i + 1)
            GenNum = Replace(GenNum, k, rs!cSuffix)
        End If
    End If
        
    rs.Update
    rs.Close
    Set rs = Nothing
        
    Generate_AutoNumber = GenNum
    
'To ensure that the Transaction Number is in Genuine Value
'    GenuineTranNo = GenNum
'    CounterfeitTranNo = GenNum
    
    If err.Number = 91 Then
        err.Clear
    End If
End Function

Public Sub Update_AutoNumber(cCode As String, Optional cType As String = "ALL", Optional nQty As Integer = 1)
'    If GenuineTranNo = CounterfeitTranNo Then
        If Year(Now) <> CurrentYear Or Month(Now) <> CurrentMonth Then
            'change of year: reseting autonumbering
            CurrentYear = Year(Now)
            CurrentMonth = Month(Now)
            cn.Execute "update PARAMETER_USER set cValue = '" & CurrentYear & "' where cParamName ='cSystemCurrentYear'"
            cn.Execute "update PARAMETER_USER set cValue = '" & CurrentMonth & "' where cParamName ='cSystemCurrentMonth'"
            
            If rs.State = adStateOpen Then rs.Close
            sSQL = "select * from AUTONUM WHERE cModuleID='" & cCode & "' and cType='" & cType & "' and lInactive=0"
            rs.Open sSQL, cn, adOpenStatic, adLockReadOnly
            
            If rs.RecordCount <> 0 Then
                sSQL = "UPDATE AUTONUM SET nValue= " & nCurrentCounter & " WHERE cModuleID='" & cCode & "' and cType='" & cType & "' and lInactive=0"
                cn.Execute sSQL
            Else
                sSQL = "UPDATE AUTONUM SET nValue= " & nCurrentCounter & " WHERE cModuleID='" & cCode & "' and cType='ALL' and lInactive=0"
                cn.Execute sSQL
            End If
        
        Else

            If rs.State = adStateOpen Then rs.Close
            sSQL = "select * from AUTONUM WHERE cModuleID='" & cCode & "' and cType='" & cType & "' and lInactive=0"
            rs.Open sSQL, cn, adOpenStatic, adLockReadOnly

            If rs.RecordCount <> 0 Then
                sSQL = "UPDATE AUTONUM SET nValue= nValue +  " & nQty & " WHERE cModuleID='" & cCode & "' and cType='" & cType & "' and lInactive=0"
                cn.Execute sSQL
            Else
                sSQL = "UPDATE AUTONUM SET nValue= nValue +  " & nQty & " WHERE cModuleID='" & cCode & "' and cType='ALL' and lInactive=0"
                cn.Execute sSQL
            End If
        End If
'    End If
End Sub

Public Sub GetConst_AutoNum_Param(ByRef sPrefix As String, ByRef sDay As String, ByRef sMonth As String, _
    ByRef sYear As String, ByRef sSeparator As String, ByRef sNumeric As String, ByRef sSuffix As String)
    
    sPrefix = "@"
    sDay = "D"
    sMonth = "M"
    sYear = "Y"
    sSeparator = "-"
    sNumeric = "#"
    sSuffix = "*"
    
'    Dim I As Integer
'    With rs
'        If .State = adStateOpen Then .Close
'        sSQL = "select cParamName,LEFT(cDesc,CHARINDEX(' ',cDesc)) as cDesc,cValue " & _
'                     "   from PARAMETER_USER where cType = 'NumSETTING' order by nOrder"
'        .Open sSQL, cn, adOpenStatic, adLockReadOnly
'
'        For I = 0 To rs.RecordCount - 1
'            If Trim(rs!cDesc) = "Prefix" Then
'                sPrefix = rs!cValue
'            ElseIf Trim(rs!cDesc) = "Day" Then
'                sDay = rs!cValue
'            ElseIf Trim(rs!cDesc) = "Month" Then
'                sMonth = rs!cValue
'            ElseIf Trim(rs!cDesc) = "Year" Then
'                sYear = rs!cValue
'            ElseIf Trim(rs!cDesc) = "Separator" Then
'                sSeparator = rs!cValue
'            ElseIf Trim(rs!cDesc) = "Numeric" Then
'                sNumeric = rs!cValue
'            ElseIf Trim(rs!cDesc) = "Suffix" Then
'                sSuffix = rs!cValue
'            End If
'
'            rs.MoveNext
'        Next I
'
'    End With
End Sub

Public Sub Validate_AutoNumber(ByRef sRS As Recordset, sField As String, sTable As String, sFilter As String, _
    cCode As String, Optional cType As String = "ALL")
    
    Dim LoopCtrl As Boolean
    LoopCtrl = False
    
    On Error Resume Next
    
    Do Until LoopCtrl = True
        DoEvents
        If (IDExisting(sRS, sField, sTable, sFilter) = True) Then
            Update_AutoNumber "SALESTICK"
            sRS.Fields("cTranNo") = Generate_AutoNumber("SALESTICK")
            sFilter = sRS.Fields("cTranNo")
        Else
            LoopCtrl = True
        End If
    Loop
    
    
    
'    CounterfeitTranNo = sFilter
    
End Sub

