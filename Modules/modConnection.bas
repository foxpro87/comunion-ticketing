Attribute VB_Name = "modConnection"
'IT Group Inc. 2005.09.23

Option Explicit

Public Property Get cn() As ADODB.Connection
    Set cn = glbCN
End Property

Public Property Set cn(ByVal vNewValue As ADODB.Connection)
    Set glbCN = vNewValue
End Property

Public Property Get cnShape() As ADODB.Connection
    Set cnShape = glbCNShape
End Property

Public Property Set cnShape(ByVal vNewValue As ADODB.Connection)
    Set glbCNShape = vNewValue
End Property

Public Sub OpenConnection(strConnection As String)
    Set cn = New ADODB.Connection
    With cn
        .CursorLocation = adUseClient
        .ConnectionString = strConnection
        .Open
    End With
End Sub

Public Sub OpenShapeConnection(strConnection As String)
    Set cnShape = New ADODB.Connection
    With cnShape
        .CursorLocation = adUseClient
        .ConnectionString = strConnection
        .Open
    End With
End Sub
