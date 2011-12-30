Attribute VB_Name = "modReportWriter"
Declare Function vbEncodeIPtr Lib "p2smon.dll" (X As Object) As String
Declare Function CreateReportOnRuntimeDS Lib "p2smon.dll" (X As Object, ByVal reportPath$, ByVal fieldDefFilePath$, ByVal bOverWriteExistingFiles%, ByVal bLaunchDesigner%) As Integer
Declare Function CreateFieldDefFile Lib "p2smon.dll" (X As Object, ByVal fieldDefFilePath$, ByVal bOverWriteExistingFiles%) As Integer


Public Function cleanChars(strString As String) As String
'AL -04/03/2001-
'Process:   Function Cleans a string from Special, CTRL,Return, and characters not
'-  normally visible. Only allows visible characters in string, strips out others.
'Vars:      strString:  string you want to clean
'Returns:   Clean string to cleanChars

    Dim X As Long
    Dim strAns As String
    Dim strOut As String
    
'AL -04/03/2001- Loops through each character in string
    For X = 1 To Len(strString)
        strAns = Mid(strString, X, (Len(strString)))
'AL -04/03/2001- Checks to see if character is normal visible
        If Asc(strAns) > 31 And Asc(strAns) < 128 Then
'AL -04/03/2001- Reconstruct string w/normal characters only
        strOut = strOut & Mid(strAns, 1, 1)
        f00_DynamoReport1.txtCode = strOut
        Else
        strOut = strOut & Space(1)
        f00_DynamoReport1.txtCode = strOut
        End If
    Next X
    
    Clipboard.SetText strOut
    cleanChars = strOut
End Function

