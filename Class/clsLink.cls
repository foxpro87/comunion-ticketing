VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub LinkForm(frm As Form, ModName As String, RefCode As String)
    With frm
        If .Mode = 1 Then
            MsgBox "Sytem detected that '" & Trim(ModName) & "' is in Add/Edit mode." & _
                    vbCr & "Please save/undo the current transaction.", vbCritical + vbOKOnly, "Comunion"
            FrmName.SetFocus
            Exit Sub
        Else
            .Show
            .ZOrder
            .ShowForm (RefCode)
        End If
    End With
End Sub
    
