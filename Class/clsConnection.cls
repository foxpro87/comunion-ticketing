VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'IT Group Inc. 2005.09.23

Option Explicit

Public Sub OpenNewConnection(cConn As ADODB.Connection)
    
    Set cConn = Nothing
    Set cConn = New ADODB.Connection
    cConn.CursorLocation = adUseClient
    

    cConn.ConnectionString = "driver={" & sDBDriver & "};" & _
    "server=" & sServer & ";uid=sa;pwd=" & sDBPassword & ";database=" & sDBname

    cConn.Open

End Sub


