VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAutoNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public oForm As New clsFormSetup


Public Function Generate_CtrlNo2(sMenuName As String, Optional ByVal sDate As String) As String
'Dim oRecordset As New clsRecordset
'Dim oForm As New clsFormSetup
Dim sFormat As String, sNumerals As String, sAlphabet As String, sReset As String, sDateVal As String
Dim nNum As Double, sCtrlNo As String
On Error Resume Next
'On Error GoTo errordesc
    sFormat = GetValueFrTable("cNumberFormat", "system_autonumber", "cCode ='" & sMenuName & "'", True)
    sReset = GetValueFrTable("cResetCtr", "system_autonumber", "cCode ='" & sMenuName & "'", True)
    sAlphabet = GetValueFrTable("cAlphaVal", "system_autonumber", "cCode ='" & sMenuName & "'", True)
    nNum = GetValueFrTable("nCtr", "system_autonumber", "cCode ='" & sMenuName & "'", True)
    sDateVal = GetValueFrTable("cDate", "system_autonumber", "cCode ='" & sMenuName & "'", True)
    If sDateVal = "" Then sDateVal = sDate
    Select Case sReset
        Case ""
            nNum = nNum + 1
        Case "NONE"
            nNum = nNum + 1
        Case "DAILY"
            If Format(sDate, "MMDDYYYY") = Format(sDateVal, "MMDDYYYY") Then
                nNum = nNum + 1
            Else
                nNum = 1
            End If
        Case "MONTHLY"
            If Format(sDate, "MMYYYY") = Format(sDateVal, "MMYYYY") Then
                nNum = nNum + 1
            Else
                nNum = 1
            End If
        Case "ANNUALY"
            If Format(sDate, "YYYY") = Format(sDateVal, "YYYY") Then
                nNum = nNum + 1
            Else
                nNum = 1
            End If
    End Select

    sCtrlNo = Replace(sFormat, GetStringVal(sFormat, "#"), Format(nNum, FormatNumeric(nNum, Len(GetStringVal(sFormat, "#")))))
    sCtrlNo = Replace(sCtrlNo, GetStringVal(sCtrlNo, "@"), sAlphabet)
    sCtrlNo = Replace(sCtrlNo, GetStringVal(sCtrlNo, "m"), Format(sDate, GetStringVal(sCtrlNo, "m")))
    sCtrlNo = Replace(sCtrlNo, GetStringVal(sCtrlNo, "d"), Format(sDate, GetStringVal(sCtrlNo, "d")))
    sCtrlNo = Replace(sCtrlNo, GetStringVal(sCtrlNo, "y"), Format(sDate, GetStringVal(sCtrlNo, "y")))
    Generate_CtrlNo2 = sCtrlNo
    FrmName.strNo = sCtrlNo
    'Set oRecordset = Nothing
    'Set oForm = Nothing
'errordesc:
'   MsgBox Err.Description
End Function

Private Function GetStringVal(sVal As String, char As String) As String
Dim i As Integer
Dim sChar As String
    For i = 1 To Len(sVal)
        If char = Mid(sVal, i, 1) Then
            sChar = sChar & Mid(sVal, i, 1)
        End If
    Next i
    GetStringVal = sChar
End Function

Private Function FormatNumeric(n As Double, nLen As Double) As String
Dim i As Integer
Dim sVal As String
    For i = 1 To nLen
        sVal = sVal & "0"
    Next i
    FormatNumeric = sVal
End Function
'----------------------------------------------------------------------------------------------------------------------------------------------
Public Sub Update_CtrlNo2(sTranNo As String, sMenuName As String, sTranDate As String)
On Error GoTo ErrHandler
Dim sFormat As String
Dim sSQL As String, nPos As Double, nVal As Double
    sFormat = GetValueFrTable("cNumberFormat", "system_autonumber", "cCode ='" & sMenuName & "'", True)
    nPos = InStr(sFormat, GetStringVal(sFormat, "#"))
    nVal = CDbl(Mid(sTranNo, nPos))
    
    sSQL = "EXEC rsp_Update_AutoNumber '" & sMenuName & "' ," & nVal & ",'" & sTranDate & "'"
    cn.Execute sSQL
    
ErrHandler:
If Err.Number <> 0 Then
    ErrorLog Err.Number, Err.Description, "Update Autonumber"
End If
End Sub


Public Function Check_CtrlNo2(rs As ADODB.Recordset, sTblName As String, _
sFldName As String, sTranNo As String, sMenuName As String, sDate As String) As String
'Dim oRecordset As New clsRecordset

Dim lCtrlNoOk As Boolean
Dim sFormat As String, sDateVal As String, sReset  As String, sAlphaVal  As String
Dim nNum As Double
Dim strNewNo As String
On Error Resume Next

        strNewNo = sTranNo
        sFormat = GetValueFrTable("cNumberFormat", "system_autonumber", "cCode ='" & sMenuName & "'", True)
        sReset = GetValueFrTable("cResetCtr", "system_autonumber", "cCode ='" & sMenuName & "'", True)
        sAlphaVal = GetValueFrTable("cAlphaVal", "system_autonumber", "cCode ='" & sMenuName & "'", True)
        sDateVal = GetValueFrTable("cDate", "system_autonumber", "cCode ='" & sMenuName & "'", True)
        nNum = GetValueFrTable("nCtr", "system_autonumber", "cCode ='" & sMenuName & "'", True)
        If sDateVal = "" Then sDateVal = sDate
        Do Until lCtrlNoOk
            If IDExisting(rs, sFldName, sTblName, strNewNo) Then
                Select Case sReset
                    Case ""
                        nNum = nNum + 1
                    Case "NONE"
                        nNum = nNum + 1
                    Case "DAILY"
                        If Format(sDate, "MMDDYYYY") = Format(sDateVal, "MMDDYYYY") Then
                            nNum = nNum + 1
                        Else
                            nNum = 1
                        End If
                    Case "MONTHLY"
                        If Format(sDate, "MMYYYY") = Format(sDateVal, "MMYYYY") Then
                            nNum = nNum + 1
                        Else
                            nNum = 1
                        End If
                    Case "ANNUALY"
                        If Format(sDate, "YYYY") = Format(sDateVal, "YYYY") Then
                            nNum = nNum + 1
                        Else
                            nNum = 1
                        End If
                End Select
                strNewNo = Replace(sFormat, GetStringVal(sFormat, "#"), Format(nNum, FormatNumeric(nNum, Len(GetStringVal(sFormat, "#")))))
                strNewNo = Replace(strNewNo, GetStringVal(strNewNo, "@"), sAlphaVal)
                strNewNo = Replace(strNewNo, GetStringVal(strNewNo, "m"), Format(sDate, GetStringVal(strNewNo, "m")))
                strNewNo = Replace(strNewNo, GetStringVal(strNewNo, "d"), Format(sDate, GetStringVal(strNewNo, "d")))
                strNewNo = Replace(strNewNo, GetStringVal(strNewNo, "y"), Format(sDate, GetStringVal(strNewNo, "y")))
            Else
                lCtrlNoOk = True
            End If
        Loop
        
    Check_CtrlNo2 = strNewNo
    FrmName.strNo = strNewNo
    
    'Set oRecordset = Nothing
    'Set oForm = Nothing
End Function



