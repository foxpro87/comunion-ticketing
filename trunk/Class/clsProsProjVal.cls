VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProsProjVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function CheckProspectiveCode(cExpType As String, ProjectID As String) As Boolean
Dim cProjectNo As String
cProjectNo = ""
CheckProspectiveCode = True
    ' -- Checking for Prospective
        If Not IsNull(ProjectID) Or Trim(ProjectID) <> "" Then
            If cExpType = "Marketing" Then
                If (GetValueFrTable("cLeadID", "PMS_LEAD", "cLeadID = '" & Mid(ProjectID, 1, 8) & "'") = "") And cExpType = "Marketing" Then
                    CheckProspectiveCode = False
                    MsgBox "Project ID does not exists!", vbInformation, "ComUnion"
                    Exit Function
                End If
                ' Checking of status of Prospective Project
                If GetValueFrTable("cStatus", "PMS_LEAD", "cLeadID = '" & Mid(ProjectID, 1, 8) & "'") = "Active" Then
                        If Date > CDate(GetValueFrTable("dExpiry", "PMS_LEAD", "cLeadID = '" & Mid(ProjectID, 1, 8) & "'")) Then
                                CheckProspectiveCode = False
                                MsgBox "Prospective Project Code has expired." & vbNewLine & "Please check the expiry date!", vbInformation, "ComUnion"
                                Exit Function
                        Else
                                If GetExpenseAmount(ProjectID) > GetProbableServiceAmount(ProjectID) Then
                                        CheckProspectiveCode = False
                                        MsgBox "Total Expense exceeds 2% of Probable Service Fee. " & vbNewLine & "Please Check the Total Expense of Prospective Project Code " & ProjectID, vbInformation, "ComUnion"
                                        Exit Function
                                End If
                        End If
                ElseIf GetValueFrTable("cStatus", "PMS_LEAD", "cLeadID = '" & Mid(ProjectID, 1, 8) & "'") = "Won" Then
                        cProjectNo = GetValueFrTable("cProjNo", "PMS_LEAD", "cLeadID = '" & Mid(ProjectID, 1, 8) & "'")
                        If GetValueFrTable("cStatus", "PMS_Project", "cProjNo = '" & Trim(cProjectNo) & "'") = "Active" Then
                        
                            If Date > DateAdd("d", 30, CDate(GetValueFrTable("dDate", "PMS_LEAD a LEFT JOIN PMS_Project b ON a.cLeadID = substring(b.cLeadID,1,8)  AND a.cCompanyID = b.cCompanyID", "a.cCompanyID = '" & COID & "' and a.cLeadID = '" & Mid(ProjectID, 1, 8) & "'", True))) Then
                                CheckProspectiveCode = False
                                MsgBox "Prospective Project Code already Won. " & vbNewLine & "Please charged to approved project Code", vbInformation, "ComUnion"
                                Exit Function
                            End If
                        ElseIf GetValueFrTable("cStatus", "PMS_Project", "cProjNo = '" & Trim(cProjectNo) & "'") = "Float" Then
                            If Date > DateAdd("d", 30, CDate(GetValueFrTable("dDate", "PMS_LEAD a LEFT JOIN PMS_Project b ON a.cLeadID = substring(b.cLeadID,1,8)  AND a.cCompanyID = b.cCompanyID", "a.cCompanyID = '" & COID & "' and a.cLeadID = '" & Mid(ProjectID, 1, 8) & "'", True))) Then
                                CheckProspectiveCode = False
                                MsgBox "Prospective Project Code Already Won. " & vbNewLine & "System allows charging between " & GetValueFrTable("dDate", "PMS_LEAD a LEFT JOIN PMS_Project b ON a.cLeadID = substring(b.cLeadID,1,8)  AND a.cCompanyID = b.cCompanyID", "a.cCompanyID = '" & COID & "' and a.cLeadID = '" & Mid(ProjectID, 1, 8) & "'", True) & " plus 30 days.", vbInformation, "ComUnion"
                                Exit Function
                            End If
                        Else
                            CheckProspectiveCode = False
                            MsgBox "Prospective Project Code already won.But Project ID is not Active.", vbInformation, "ComUnion"
                            Exit Function
                        End If
                Else
                    CheckProspectiveCode = False
                    MsgBox "Prospective Project Code is not active!", vbInformation, "ComUnion"
                    Exit Function
                End If
            Else
                CheckProspectiveCode = False
                MsgBox "Please change the type to 'Marketing'.", vbInformation, "ComUnion"
                Exit Function
            End If
        End If
End Function

Public Function CheckProjectCode(cExpType As String, ProjectID As String) As Boolean
CheckProjectCode = True
        If Not IsNull(ProjectID) Or Trim(ProjectID) <> "" Then
            If cExpType = "Project" Then
                If GetValueFrTable("cProjectID", "PMS_PROJECT_SERVICES", "cProjectID = '" & ProjectID & "'") = "" And cExpType = "Project" Then
                    CheckProjectCode = False
                    MsgBox "Project ID does not exists!", vbInformation, "ComUnion"
                    Exit Function
                End If
                If GetValueFrTable("cStatus", "PMS_Project_Services a left join PMS_Project b on a.cCompanyID = b.cCompanyID and a.cProjNo = b.cProjNo", "a.cCompanyID = '" & COID & "' and a.cProjectID = '" & ProjectID & "'", True) <> "Active" Then
                    CheckProjectCode = False
                    MsgBox "Project ID is not active!", vbInformation, "ComUnion"
                    Exit Function
                Else
                        If GetServiceFeeAmount(ProjectID) <= 0 Then
                             CheckProjectCode = False
                             MsgBox "System not allows zero Service fee!", vbInformation, "ComUnion"
                             Exit Function
                        End If
                
                End If
            Else
                    CheckProjectCode = False
                    MsgBox "Please change the type to 'Project'.", vbInformation, "ComUnion"
                    Exit Function
            End If
        End If
End Function

