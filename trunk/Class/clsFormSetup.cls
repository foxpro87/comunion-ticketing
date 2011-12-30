VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub FormLocking(bval As Boolean)
    For Each objCtl In FrmName
        If TypeName(objCtl) = "ITGTextBox" Then
            objCtl.Locked = bval
            objCtl.LabelForeColor = &H0&
        ElseIf TypeName(objCtl) = "TextBox" Then
            objCtl.Locked = bval
        ElseIf TypeName(objCtl) = "ITGCommandButton" Then
            objCtl.Enabled = Not bval
        ElseIf TypeName(objCtl) = "CommandButton" Then
            objCtl.Enabled = Not bval
        ElseIf TypeName(objCtl) = "ComboBox" Then
            objCtl.Enabled = Not bval
        ElseIf TypeName(objCtl) = "MaskEdBox" Then
            objCtl.Enabled = Not bval
        ElseIf TypeName(objCtl) = "ITGDateBox" Then
            objCtl.Enabled = Not bval
        ElseIf TypeName(objCtl) = "ITGCheckBox" Then
            objCtl.Enabled = Not bval
        ElseIf TypeName(objCtl) = "OptionButton" Then
            objCtl.Enabled = Not bval
        ElseIf TypeName(objCtl) = "DataGrid" Then
            objCtl.AllowUpdate = Not bval
        ElseIf TypeName(objCtl) = "RichTextBox" Then
            objCtl.Locked = bval
        ElseIf TypeName(objCtl) = "ComunionButton" Then
            objCtl.Enabled = Not bval
        ElseIf TypeName(objCtl) = "ITGCombobox" Then
            objCtl.Enabled = Not bval
        End If
    Next
End Sub

Public Sub FormSearch(bval As Boolean)
    For Each objCtl In FrmName
        If TypeName(objCtl) = "ITGTextBox" Then
            If objCtl.Mandatory = True Then
                objCtl.Locked = Not bval
                objCtl.LabelForeColor = &HC00000
            Else
                objCtl.Locked = bval
            End If
        ElseIf TypeName(objCtl) = "TextBox" Then
            objCtl.Locked = bval
        ElseIf TypeName(objCtl) = "ITGCommandButton" Then
            objCtl.Enabled = Not bval
        ElseIf TypeName(objCtl) = "CommandButton" Then
            objCtl.Enabled = Not bval
        ElseIf TypeName(objCtl) = "ComboBox" Then
            objCtl.Enabled = Not bval
        ElseIf TypeName(objCtl) = "MaskEdBox" Then
            objCtl.Enabled = Not bval
        ElseIf TypeName(objCtl) = "ITGDateBox" Then
            objCtl.Enabled = Not bval
        ElseIf TypeName(objCtl) = "ITGCheckBox" Then
            objCtl.Enabled = Not bval
        ElseIf TypeName(objCtl) = "OptionButton" Then
            objCtl.Enabled = Not bval
        ElseIf TypeName(objCtl) = "DataGrid" Then
            objCtl.AllowUpdate = Not bval
        ElseIf TypeName(objCtl) = "RichTextBox" Then
            objCtl.Locked = bval
'        ElseIf TypeName(objCtl) = "ITGCombobox" Then
'            objCtl.Locked = bval
        End If
    Next
End Sub

Public Sub TextClearing()
    For Each objCtl In FrmName
        If TypeName(objCtl) = "ITGTextBox" Then
            objCtl.Text = ""
        ElseIf TypeName(objCtl) = "TextBox" Then
            objCtl.Text = ""
        ElseIf TypeName(objCtl) = "RichTextBox" Then
            objCtl.Text = ""
        ElseIf TypeName(objCtl) = "MaskEdBox" Then
            objCtl.Mask = "99/99/9999"
        ElseIf TypeName(objCtl) = "ITGDateBox" Then
            objCtl.Text = "__/__/____"
        ElseIf TypeName(objCtl) = "ITGCheckBox" Then
            objCtl.Value = Unchecked
        ElseIf TypeName(objCtl) = "DataGrid" Then
            objCtl.Refresh
        ElseIf TypeName(objCtl) = "ITGCombobox" Then
            objCtl.Text = ""
        ElseIf TypeName(objCtl) = "Image" Then
            If objCtl.Name <> "imgLogo" Then
                objCtl.Picture = Nothing
            End If
        End If
    Next
End Sub

Public Sub FormTheme(Theme1 As Integer)
    For Each objCtl In FrmName
        If Theme1 = 1 Then
            If TypeName(objCtl) = "ITGTextBox" Then
                objCtl.BackColor = &HF7F2E6
                objCtl.LabelBackColor = &HFBF4F4
            End If
            If TypeName(objCtl) = "ITGTab" Then
                objCtl.ActiveTabBackEndColor = &HFBFDFB
                objCtl.ActiveTabBackStartColor = &HB49799
                objCtl.TabStripBackColor = &H80000014
                objCtl.InActiveTabBackStartColor = &HFFFFFF
            End If
            If TypeName(objCtl) = "ComunionFrames" Then
                objCtl.ThemeColor = 1
                objCtl.FillColor = &HFBF4F4
            End If
            If TypeName(objCtl) = "ComunionLabel" Then
                objCtl.BackColor = &HFBF4F4
            End If
            If TypeName(objCtl) = "ITGDateBox" Then
                objCtl.BackColor = &HF7F2E6
            End If
            If TypeName(objCtl) = "ITGCombobox" Then
                objCtl.BackColor = &HF7F2E6
                objCtl.CaptionBackColor = &HFBF4F4
            End If
            If TypeName(objCtl) = "ComboBox" Then
                objCtl.BackColor = &HF7F2E6
            End If
            If TypeName(objCtl) = "DataGrid" Then
                objCtl.BackColor = &HF7F2E6
            End If
            If TypeName(objCtl) = "ITGCheckBox" Then
                objCtl.BackColor = &HFBF4F4
            End If
            If TypeName(objCtl) = "Label" Then
                objCtl.BackColor = &HFBF4F4
            End If
            If TypeName(objCtl) = "Form" Then
                objCtl.BackColor = &HFBF4F4
            End If
            FrmName.BackColor = &H80000014
        ElseIf Theme1 = 2 Then
            If TypeName(objCtl) = "ITGTextBox" Then
                objCtl.BackColor = &HC0FFFF
                objCtl.LabelBackColor = &HFFFFFF
            End If
            If TypeName(objCtl) = "ITGTab" Then
                objCtl.ActiveTabBackEndColor = &HFFFFFF
                objCtl.ActiveTabBackStartColor = &HA5EFFF
                objCtl.TabStripBackColor = &H80000018
                objCtl.InActiveTabBackStartColor = &HC0FFFF
            End If
            If TypeName(objCtl) = "ComunionFrames" Then
                objCtl.ThemeColor = 4
                objCtl.FillColor = &HFFFFFF
            End If
            If TypeName(objCtl) = "ComunionLabel" Then
                objCtl.BackColor = &HFFFFFF
            End If
            If TypeName(objCtl) = "ITGDateBox" Then
                objCtl.BackColor = &HC0FFFF
            End If
            If TypeName(objCtl) = "ITGCombobox" Then
                objCtl.BackColor = &HC0FFFF
                objCtl.CaptionBackColor = &HFFFFFF
            End If
            If TypeName(objCtl) = "ComboBox" Then
                objCtl.BackColor = &HC0FFFF
            End If
            If TypeName(objCtl) = "DataGrid" Then
                objCtl.BackColor = &HC0FFFF
            End If
            If TypeName(objCtl) = "ITGCheckBox" Then
                objCtl.BackColor = &HFFFFFF
            End If
            If TypeName(objCtl) = "Label" Then
                objCtl.BackColor = &HFFFFFF
            End If
            If TypeName(objCtl) = "Form" Then
                objCtl.BackColor = &HFFFFFF
            End If
            FrmName.BackColor = &H80000018
        ElseIf Theme1 = 3 Then
            If TypeName(objCtl) = "ITGTextBox" Then
                objCtl.BackColor = &HE1F9F0
                objCtl.LabelBackColor = &HFFFFFF
            End If
            If TypeName(objCtl) = "ITGTab" Then
                objCtl.ActiveTabBackEndColor = &HE1F9F7
                objCtl.ActiveTabBackStartColor = &H8FC5B5
                objCtl.TabStripBackColor = &H80000014
                objCtl.InActiveTabBackStartColor = &HFFFFFF
            End If
            If TypeName(objCtl) = "ComunionFrames" Then
                objCtl.ThemeColor = 2
                objCtl.FillColor = &HFFFFFF
            End If
            If TypeName(objCtl) = "ComunionLabel" Then
                objCtl.BackColor = &HFFFFFF
            End If
            If TypeName(objCtl) = "ITGDateBox" Then
                objCtl.BackColor = &HE1F9F0
            End If
            If TypeName(objCtl) = "ITGCombobox" Then
                objCtl.BackColor = &HE1F9F0
                objCtl.CaptionBackColor = &HFFFFFF
            End If
            If TypeName(objCtl) = "ComboBox" Then
                objCtl.BackColor = &HE1F9F0
            End If
            If TypeName(objCtl) = "DataGrid" Then
                objCtl.BackColor = &HE1F9F0
            End If
            If TypeName(objCtl) = "ITGCheckBox" Then
                objCtl.BackColor = &HFFFFFF
            End If
            If TypeName(objCtl) = "Label" Then
                objCtl.BackColor = &HFFFFFF
            End If
            If TypeName(objCtl) = "Form" Then
                objCtl.BackColor = &HFFFFFF
            End If
        End If
    Next
End Sub

Public Sub LockPrimary(objTxt As ITGTextBox)
    objTxt.Locked = True
End Sub

Public Sub ClrRequired(ClrCode As Long)
    For Each objCtl In FrmName
        If TypeName(objCtl) = "ITGTextBox" Then
            If objCtl.Required = True Then
                objCtl.LabelForeColor = ClrCode
            End If
        End If
    Next
End Sub

'Public Property Get FrmName() As Object
'    Set FrmName = mFrmName
'End Property
'
'Public Property Set FrmName(ByVal sFrmName As Object)
'   Set mFrmName = sFrmName
'End Property
