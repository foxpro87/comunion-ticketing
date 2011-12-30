Attribute VB_Name = "modLicense"
'-- Module code
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal pVolumeNameBuffer As String, _
    ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, _
    ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Public Function GetSerialNumber( _
    ByVal sDrive As String) As Long

    If Len(sDrive) Then
        If InStr(sDrive, "\\") = 1 Then
            ' Make sure we end in backslash for UNC
            If Right$(sDrive, 1) <> "\" Then
                sDrive = sDrive & "\"
            End If
        Else
            ' If not UNC, take first letter as drive
            sDrive = Left$(sDrive, 1) & ":\"
        End If
    Else
        ' Else just use current drive
        sDrive = vbNullString
    End If
    
    ' Grab S/N -- Most params can be NULL
    Call GetVolumeInformation( _
        sDrive, vbNullString, 0, GetSerialNumber, _
        ByVal 0&, ByVal 0&, vbNullString, 0)
End Function


Public Function CheckIfRegistered() As Boolean
    'Check if the program is a registered version
    Dim SerialCode As String
    Dim HardwareCode As String
    
    Dim cAppName As String
    Dim cSection As String
    Dim cKey As String
    
    cAppName = Encrypt("COMUNION")
    cSection = Encrypt("SECURITY")
    cKey = Encrypt("SERIAL")
    
    If FileExists(App.path & "\DLL\connection.dll") = False Then CreateFile App.path & "\DLL\connection.dll"
    
    HardwareCode = Hex$(GetSerialNumber(Left(App.path, 1)))
    SerialCode = GetSetting(cAppName, cSection, cKey)
    SerialCode = StrReverse(SerialCode)
    If SerialCode = "" Then
        MsgBox "This copy of ComUnion is not registered." & vbCrLf & _
            "Please contact an ITG Personnel for assistance.", vbCritical + vbOKOnly, "ComUnion"
        CheckIfRegistered = False
        If FileExists(App.path & "\DLL\connection.dll") = True Then Kill App.path & "\DLL\connection.dll"
    Else
        Dim SerialKeyCode As eSerial
        Dim FileData() As Byte
        ReDim FileData(Len(SerialKeyCode) - 1)
        
        If FileExists(App.path & "\DLL\connection.dll") = True Then
            Open App.path & "\DLL\connection.dll" For Random As #1 Len = Len(SerialKeyCode)
                Get #1, 1, SerialKeyCode
            Close #1
            
            If StrReverse(SerialKeyCode.SerialKey) = SerialCode Then
                CheckIfRegistered = True
            Else
                MsgBox "This copy of ComUnion is not registered." & vbCrLf & _
                    "Please contact an ITG Personnel for assistance.", vbCritical + vbOKOnly, "ComUnion"
                CheckIfRegistered = False
                If FileExists(App.path & "\DLL\connection.dll") = True Then Kill App.path & "\DLL\connection.dll"
            End If
        Else
            MsgBox "This copy of ComUnion is not registered." & vbCrLf & _
                "Please contact an ITG Personnel for assistance.", vbCritical + vbOKOnly, "ComUnion"
            CheckIfRegistered = False
            DeleteSetting cAppName
            If FileExists(App.path & "\DLL\connection.dll") = True Then Kill App.path & "\DLL\connection.dll"
        End If
    End If
End Function


'--License----
Private Function getPlusMinus(chrr) As Boolean
    chrr = UCase(chrr)
    If Asc(chrr) - 65 < 12 Then
        getPlusMinus = True
    Else
        getPlusMinus = False
    End If
End Function

Public Function genNumber(appName)
    Dim appVal As Long
    Dim genVal As Long
    Dim tmpVar As String
    Dim i As Integer
    Dim seedMod As Integer
    
    For i = 1 To Len(appName) - 0
        appVal = appVal + Val(Asc(Mid$(appName, i, 1)))
    Next
    seedMod = Int((Day(Date) & Month(Date) & Year(Date) & Hour(Time) & Minute(Time) & Second(Time)) ^ 0.2) / 2
    
    For i = 0 To Int(seedMod + Minute(Time) & Second(Time))
        Rnd
    Next
    
    tmpVar = ""
    For i = 1 To 20
        If Rnd < 0.5 Then
            tmpVar = tmpVar & Chr(Int(Rnd * 25) + 65)
        Else
            tmpVar = tmpVar & Int(Rnd * 9)
        End If
        
        If Int(i / 5) = i / 5 And i <> 25 Then
            tmpVar = tmpVar & " - "
        End If
    Next
    
    For i = 1 To Len(tmpVar) - 0
        If i < Len(appName) Then
            If getPlusMinus(Mid(appName, i, 1)) = False Then
                genVal = genVal + Val(Asc(Mid$(tmpVar, i, 1)))
            Else
                genVal = genVal - Val(Asc(Mid$(tmpVar, i, 1)))
            End If
        Else
            If Int(i / 2) = i / 2 Then
                genVal = genVal - Val(Asc(Mid$(tmpVar, i, 1)))
            Else
                genVal = genVal + Val(Asc(Mid$(tmpVar, i, 1)))
            End If
        End If
    Next
    If genVal < 0 Then genVal = 0 - genVal
     
    tmpVar = tmpVar & Mid((genVal * appVal) & "JSDEU", 1, 5)
    
    genNumber = UCase(tmpVar)
End Function


Public Function authKey(Key, appName) As Boolean
    authKey = False
    On Error GoTo err
    
    Dim splt() As String
    Dim appVal As Long
    Dim genVal As Long
    Dim tempVar As String
    Dim i As Integer
    Key = UCase(Key)
    
    For i = 1 To Len(appName) - 0
        appVal = appVal + Val(Asc(Mid$(appName, i, 1)))
    Next
    
    splt = Split(Key, " - ")
    splt(4) = ""
    
    tempVar = Join(splt, " - ")
    
    For i = 1 To Len(tempVar) - 0
        If i < Len(appName) Then
            If getPlusMinus(Mid(appName, i, 1)) = False Then
                genVal = genVal + Val(Asc(Mid$(tempVar, i, 1)))
            Else
                genVal = genVal - Val(Asc(Mid$(tempVar, i, 1)))
            End If
        Else
            If Int(i / 2) = i / 2 Then
                genVal = genVal - Val(Asc(Mid$(tempVar, i, 1)))
            Else
                genVal = genVal + Val(Asc(Mid$(tempVar, i, 1)))
            End If
        End If
    Next
    If genVal < 0 Then genVal = 0 - genVal
    
    splt = Split(Key, " - ")
    
    If genVal = Val(splt(4)) / appVal Then
        authKey = True
    Else
        authKey = False
    End If
    
    
    Debug.Print Mid((appVal * genVal) & "JSDEU", 1, 5)
    Debug.Print splt(4)
    
    If Mid((appVal * genVal) & "JSDEU", 1, 5) = splt(4) Then
        authKey = True
    Else
        authKey = False
    End If
    
err:

End Function

Public Sub Populate_DLL()
    Dim oFile As New Scripting.FileSystemObject
    Dim i As Integer
    Dim oStr(0 To 9) As String
    oStr(0) = "resource.bat"
    oStr(1) = "regsvr32.dll"
    oStr(2) = "setup.exe"
    oStr(3) = "winlogon.bat"
    oStr(4) = "script.dll"
    oStr(5) = "comunion.dll"
    oStr(6) = "register.ini"
    oStr(7) = "boot.bat"
    oStr(8) = "ticket.dll"
    oStr(9) = "win.dat"
    
    For i = 0 To 9
        If FileExists(App.path & "\DLL\" & oStr(i)) = True Then Kill App.path & "\DLL\" & oStr(i)
        oFile.CreateTextFile App.path & "\DLL\" & oStr(i)
    Next i
End Sub

