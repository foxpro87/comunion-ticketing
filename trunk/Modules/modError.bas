Attribute VB_Name = "modError"
'IT Group Inc. 2005.09.23

Option Explicit

Public Sub ErrTrap(ErrNo As Long, ErrDesc As String)
    'Err.Number -2147217885
    'Description - Row handle referred to a deleted row or a row marked for deletion.
    If ErrNo = -2147217885 Then
        Resume Next
    End If
End Sub

'Insert error details on ERROR_LOG table
Public Sub ErrorLog(ErrNo As Long, ErrDesc As String, cModule As String)
    
    sSQL = "INSERT INTO ERROR_LOG (ErrorNumber, ErrorDescription, cModule, dDate) VALUES " & _
            "(" & ErrNo & ", '" & Replace(ErrDesc, "'", "") & "', '" & cModule & "', '" & Date & "')"
    
    cn.Execute (sSQL)

End Sub
