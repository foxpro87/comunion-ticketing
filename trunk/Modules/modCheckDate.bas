Attribute VB_Name = "modCheckDate"
Public Function GetSplitDate(MInLine As String) As String()
    Dim i As Integer, j As Integer, FoundIt As Boolean, MKey(50) As String
         i = 1: FoundIt = False: j = 1
         Do While i <= Len(MInLine)
            If Mid(MInLine, i, 1) = "/" Then
                j = j + 1
            Else
               MKey(j) = MKey(j) + Mid(MInLine, i, 1)
            End If
            i = i + 1
         Loop
    GetSplitDate = MKey()
End Function

Public Function ChkDate(sDate As String, DateFormatIndian As Boolean) As Boolean
    Dim StrDate As String, StrMonth As String
    Dim StrChkDate() As String, strError As String
    Dim i As Integer, j As Integer, FoundIt As Boolean
    Dim iYear
    ChkDate = True: strError = ""
    StrChkDate = GetSplitDate(sDate)
    i = 1: FoundIt = False: j = 1
    If Len(Trim(sDate)) > 10 Then
        ChkDate = False
        GoTo error
    End If
    Do While i <= Len(sDate)
       If Mid(sDate, i, 1) = "/" Then
           j = j + 1
       End If
       i = i + 1
    Loop
    If j <> 3 Then
        ChkDate = False
        GoTo error
    End If
    If DateFormatIndian Then
        StrDate = StrChkDate(1)
        StrMonth = StrChkDate(2)
        iYear = StrChkDate(3)
    Else
        StrMonth = StrChkDate(1)
        StrDate = StrChkDate(2)
        iYear = StrChkDate(3)
    End If
   If Trim(StrDate) <> "" Or Trim(StrDate) <> "" Or Trim(iYear) <> "" Then
        If IsNumeric(StrDate) Or IsNumeric(StrDate) Or IsNumeric(iYear) Then
        Else
            ChkDate = False
            GoTo error
        End If
   Else
        ChkDate = False
        GoTo error
   End If
        If StrMonth <= 0 Or StrMonth > 12 Then
            strError = strError & "Enter Months Between 1 TO 12" & vbCrLf
            ChkDate = False
        End If
        If StrMonth = 1 Or StrMonth = 3 Or StrMonth = 5 Or StrMonth = 7 Or StrMonth = 8 Or StrMonth = 10 Or StrMonth = 12 Then
            If StrDate <= 0 Or StrDate > 31 Then
                strError = strError & "Enter Days Between 1 TO 31" & vbCrLf
                ChkDate = False
            End If
        Else
            If StrMonth = 4 Or StrMonth = 6 Or StrMonth = 9 Or StrMonth = 11 Then
                     If StrDate <= 0 Or StrDate > 30 Then
                         strError = strError & "Enter Days Between 1 TO 30" & vbCrLf
                         ChkDate = False
                     End If
            Else
                 If StrMonth = 2 And (iYear Mod 4 = 0) Then
                     If StrDate <= 0 Or StrDate > 28 Then
                         strError = strError & "Enter Days Between 1 TO 28" & vbCrLf
                         ChkDate = False
                     End If
                  Else
                     If StrDate <= 0 Or StrDate > 29 Then
                         strError = strError & "Enter Days Between 1 TO 29" & vbCrLf
                         ChkDate = False
                     End If
                 End If
            End If
        End If
                
error:
        If Not ChkDate Then
            If DateFormatIndian Then
                MsgBox (strError & "Date should be in DD/MM/YYYY")
            Else
                MsgBox (strError & "Date should be in MM/DD/YYYY")
            End If
            Exit Function
        End If
        
End Function

