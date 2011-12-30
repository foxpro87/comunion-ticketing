Attribute VB_Name = "modValidation"
Option Explicit

Public Function ValidKeys(KeyIn As Integer, ValidateString As String, Editable As Boolean) As Integer

'ex:  KeyAscii = ValidKeys(KeyAscii, "01234567890.", True)
Dim ValidateList As String
Dim KeyOut As Integer

    If KeyIn = 13 Then Exit Function
    
    If Editable = True Then
        ValidateList = UCase(ValidateString) & Chr(8)
    Else
        ValidateList = UCase(ValidateString)
    End If
    
    If InStr(1, ValidateList, UCase(Chr(KeyIn)), 1) > 0 Then
        KeyOut = KeyIn
    Else
        KeyOut = 0
        Beep
    End If
    
    ValidKeys = KeyOut
End Function

Public Function InvalidKeys(KeyAscii As Integer, varStr As String) As Integer
    If varStr Like ("*" & Chr(KeyAscii) & "*") Then
        InvalidKeys = 0
    Else: InvalidKeys = KeyAscii
    End If
End Function

'Sends the cursor to the next tab stop
Public Sub SendKeysTab(Key As Integer)
    If Key = 13 Then SendKeys "{Tab}"
End Sub

'Date grid validation module (boolean function)
'   x -- not mandatory field
'   d -- mandatory date field
'   n -- mandatory numeric field
'   s -- mandatory string field
'   b -- mandatory bit field
'NOT FUNCTIONING (Unfinished concept)
Public Function GridValidationOK(dtg As DataGrid, sDtg As String) As Boolean
Dim sMandatory As String
    
    GridValidationOK = True
    
    i = 0
    
    Do While Not i = dtg.Columns.Count
        sMandatory = Mid(sDtg, (i + 1), 1)
        Select Case UCase(sMandatory)
            Case "D"
                If Not IsDate(dtg.Columns(i).Value) Then
                    GridValidationOK = False
                    MsgBox dtg.Columns(i).Caption & " is mandatory. Null value is not allowed.", vbCritical, "ComUnion"
                    dtg.Col = i
                    dtg.Columns(i).Value = Date
                    Exit Function
                End If
            Case "N"
                If dtg.Columns(i).Value = 0 Then
                    GridValidationOK = False
                    MsgBox dtg.Columns(i).Caption & " is mandatory. Null or zero value is not allowed.", vbCritical, "ComUnion"
                    dtg.Col = i
                    dtg.Columns(i).Value = 0
                    Exit Function
                End If
            Case "S"
                If Trim(dtg.Columns(i).Text) = "" Then
                    GridValidationOK = False
                    MsgBox dtg.Columns(i).Caption & " is mandatory. Null value is not allowed.", vbCritical, "ComUnion"
                    dtg.Col = i
                    dtg.Columns(i).Value = ""
                    Exit Function
                End If
            'Case "B"
        End Select
        i = i + 1
    Loop
    
End Function

'This sub validates mandatory fields
Public Function ValidateMandatoryOK() As Boolean
Dim cboName As String
Dim objCtl1 As Control

    ValidateMandatoryOK = True
    
    For Each objCtl In FrmName
        If (TypeName(objCtl) = "ITGTextBox") Then
            If Trim(objCtl.Text) = "" And (objCtl.Mandatory) Then
                ValidateMandatoryOK = False
                MsgBox "Field '" & Trim(objCtl.Label) & "' is mandatory. Null value is not allowed", vbInformation, "ComUnion"
                If UCase(Mid(objCtl.Tag, 1, 3)) = "CBO" Then
                    cboName = Trim(objCtl.Tag)
                    For Each objCtl1 In FrmName
                        If (TypeName(objCtl1) = "ComboBox") Then
                            If objCtl1.Name = cboName Then
                                objCtl1.SetFocus
                                Exit Function
                            End If
                        End If
                    Next
                End If
                If objCtl.Visible Then objCtl.SetFocus
                Exit Function
            End If
        ElseIf TypeName(objCtl) = "ITGDateBox" Then
            If Trim(objCtl.Text) = "__/__/____" And (objCtl.Mandatory) Then
                ValidateMandatoryOK = False
                MsgBox "Date field " & Trim(objCtl.Tag) & " is mandatory. Null value is not allowed", vbInformation, "ComUnion"
                If objCtl.Visible Then objCtl.SetFocus
                Exit Function
            End If
        End If
    Next

End Function

'Audit trail function
Function UpdateLogFile(cModule As String, cTranNo As String, cEvent As String)
On Error Resume Next
    Dim LogCmd As ADODB.Command
    Set LogCmd = New ADODB.Command
    With LogCmd
        Set .ActiveConnection = cn
        .CommandText = "AuditTrail"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
    End With
    LogCmd.Parameters("@cModule") = LTrim(RTrim(cModule))
    LogCmd.Parameters("@cTranNo") = LTrim(RTrim(cTranNo))
    LogCmd.Parameters("@cEvent") = LTrim(RTrim(cEvent))
    LogCmd.Parameters("@cUser") = LTrim(RTrim(SecUserID))
    LogCmd.Parameters("@cCompanyID") = LTrim(RTrim(COID))
    LogCmd.Execute
    Set LogCmd.ActiveConnection = Nothing
End Function

'Gl Activity
Public Function UpdateGLA(cEvent As String, cModule As String, cTranNo As String)
On Error Resume Next
Dim LogCmd As ADODB.Command
Set LogCmd = New ADODB.Command
    With LogCmd
        Set .ActiveConnection = cn
        If cEvent = "Update" Then
            .CommandText = "USP_UpdateGL"
        Else: .CommandText = "USP_DeleteGL"
        End If
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
    End With
    LogCmd.Parameters("@cCompanyID") = LTrim(RTrim(COID))
    LogCmd.Parameters("@cModule") = LTrim(RTrim(cModule))
    LogCmd.Parameters("@cTranNo") = LTrim(RTrim(cTranNo))
    
    LogCmd.Execute
    Set LogCmd.ActiveConnection = Nothing
End Function

'FixedLength
Public Function FixedLength(KeyAscii As Integer, sText As String, nLength As Long) As Integer
    If KeyAscii = vbKeyReturn Then
        FixedLength = KeyAscii
    ElseIf KeyAscii = vbKeyDelete Then
        FixedLength = KeyAscii
    ElseIf KeyAscii = vbKeyBack Then
        FixedLength = KeyAscii
    ElseIf Len(sText) = nLength Then
        Beep
        FixedLength = 0
    Else
        FixedLength = KeyAscii
    End If
End Function

'Passing string value to variable 'sFilterString'
Public Sub PassFilterStringValue(str As String)
    If Trim(str) <> "" Then
        sFilterString = Trim(str)
    Else
        sFilterString = ""
    End If
End Sub

Public Sub InitEmailTime()
    Dim i As Integer
    Dim TimeCount As Integer
    If rs.State = 1 Then rs.Close
    rs.Open "select (case when ltimer = 0 then dTime when ltimer = 1 then convert (varchar(5), dTimer) end) as dTime from MailSetup where cCompanyID = '" & COID & "'", cn, adOpenStatic, adLockReadOnly
    ReDim TimeCollection(1 To rs.RecordCount) As String
    ReDim IsReportSend(1 To rs.RecordCount) As Boolean
    For i = 1 To rs.RecordCount
        TimeCollection(i) = rs!dTime
        rs.MoveNext
    Next i
    rs.Close
End Sub

Public Function FileExists(sFullPath As String) As Boolean
    Dim oFile As New Scripting.FileSystemObject
    FileExists = oFile.FileExists(sFullPath)
End Function

Public Sub CreateFile(cPath As String)
    Dim oFile As New Scripting.FileSystemObject
    Dim sTMP As String
    sTMP = Mid(cPath, 1, InStrRev(cPath, "\") - 1)
    If oFile.FolderExists(sTMP) = False Then
        MakeDirPath sTMP
    End If
    oFile.CreateTextFile cPath, True
End Sub

Sub MakeDirPath(dirname As String)
    Dim i As Long, path As String
    
    Do
        i = InStr(i + 1, dirname & "\", "\")
        path = Left$(dirname, i - 1)
        ' don't try to create a root directory
        If Right$(path, 1) <> ":" And Dir$(path, vbDirectory) = "" Then
            ' make this subdirectory if it doesn't exist
            ' (exits if any error)
            MkDir path
        End If
    Loop Until i >= Len(dirname)

End Sub

Public Function OverRideTransaction(cModule As String, cTranNo As String) As Boolean
    frmOverride.cModule = cModule
    frmOverride.cTranNo = cTranNo
    frmOverride.Show vbModal
    OverRideTransaction = frmOverride.lApproved
End Function

