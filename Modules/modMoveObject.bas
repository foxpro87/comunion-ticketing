Attribute VB_Name = "modMoveObject"
'IT Group Inc. 2005.09.23

Option Explicit

Public Sub MoveCombo(cbo As ComboBox, dtg As DataGrid, gcol As MSDataGridLib.Column)
On Error GoTo error_handler
    
    cbo.Visible = True
    
    Set gcol = dtg.Columns(dtg.Col)
    
    ' Move the ComboBox
    cbo.Move dtg.Left + gcol.Left, _
        dtg.Top + dtg.RowTop(dtg.Row), gcol.Width
    cbo.ZOrder
    cbo.Text = gcol.Text
    cbo.SetFocus
    Exit Sub

error_handler:
    Resume Next
End Sub

Public Sub MoveText(txt As TextBox, dtg As DataGrid, gcol As MSDataGridLib.Column)
On Error GoTo error_handler
Dim nTop
    
    txt.Visible = True
    
    Set gcol = dtg.Columns(dtg.Col)
    
    If (dtg.Height - (dtg.Top + dtg.RowTop(dtg.Row))) > (txt.Height - dtg.Top) Then
        nTop = (dtg.Top + dtg.RowTop(dtg.Row) + dtg.RowHeight)
    Else: nTop = (dtg.Top + dtg.RowTop(dtg.Row)) - txt.Height
    End If
    
    ' Move the textbox
    txt.Move dtg.Left + gcol.Left, nTop
    txt.ZOrder
    txt.Text = gcol.Text
    txt.SetFocus
    Exit Sub

error_handler:
    Resume Next
End Sub

'Combo box value upon startup/initialization
Public Sub ComboLoadValue(cbo As ComboBox, strValue As String)
On Error GoTo error_handler
    If Trim(strValue) <> "" Then
        cbo = Trim(strValue)
    Else
        cbo.ListIndex = 0
    End If
error_handler:
    If Err.Number = 383 Then
        Resume Next
    End If
End Sub

'Sorts recordset
Public Sub SortGrid(dtg As DataGrid, ColIndex As Integer, rsT As Recordset)
Dim strDTFld As String
    
    strDTFld = dtg.Columns(ColIndex).DataField
    
    If rsT.Sort = strDTFld & " ASC" Then
        rsT.Sort = strDTFld & " DESC"
    Else: rsT.Sort = strDTFld & " ASC"
    End If

End Sub

Public Sub FormCenter(frm As Form, Optional frmMDI As Variant)
'*************************
'Jon Fonacier
'*************************
    'Syntax:
    'Center Form1, [MDIForm1]
    Dim ret As Boolean
    ret = IsMissing(frmMDI)

    If ret = False Then
        frm.Top = (frmMDI.ScaleHeight - frm.Height) / 2
        frm.Left = (frmMDI.ScaleWidth - frm.Width) / 2
    Else
        frm.Top = (Screen.Height - frm.Height) / 2
        frm.Left = (Screen.Width - frm.Width) / 2
    End If
End Sub

'Show waiting form
Public Sub FormWaitShow(Msg As String, Optional bval As Boolean)
On Error GoTo TheSource
    frmWait.ITGLabel1.Caption = Msg '"Loading data. . ."
    frmWait.ITGLabel3.Visible = bval
    frmWait.ITGLabel4.Visible = bval
    frmWait.Height = 2200
    Load frmWait
    frmWait.Show
    DoEvents
TheSource:
    Resume Next
End Sub
