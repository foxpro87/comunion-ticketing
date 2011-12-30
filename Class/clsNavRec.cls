VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNavRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'IT Group Inc. 2005.09.23

Option Explicit

'=========================================================
'Procudure Name:  MoveFirst
'Synopsis:        Move Record at Top
'e.g              MoveFirst RS-(your Recordset)
'========================================================
Public Sub MoveFirst(ByRef oPayRec As ADODB.Recordset)
On Error GoTo ErrorHandler
    
    If oPayRec.RecordCount > 0 Then oPayRec.MoveFirst

ErrorHandler:
    If Err.Number = -2147217842 Then
        Resume Next
    End If

End Sub

'========================================================
'Procudure Name: MovePrevious
'Synopsis:       Move Previous Record
'e.g             MovePrevious RS-(your Recordset)
'========================================================
Public Sub MovePrevious(ByRef oPayRec As ADODB.Recordset)
On Error GoTo ErrorHandler
    
    If oPayRec.RecordCount > 0 Then
        If Not oPayRec.BOF Then oPayRec.MovePrevious
        If oPayRec.BOF Then
            Beep
            oPayRec.MoveFirst
        End If
    End If

ErrorHandler:
    If Err.Number = -2147217842 Then
        Resume Next
    End If

End Sub

'========================================================
'Procudure Name: MoveNext
'Synopsis:       Move MoveNext Record
'e.g             MoveNext RS-(your Recordset)
'========================================================
Public Sub MoveNext(ByRef oPayRec As ADODB.Recordset)
On Error GoTo ErrorHandler
    
    If oPayRec.RecordCount > 0 Then
        If Not oPayRec.EOF Then oPayRec.MoveNext
        If oPayRec.EOF Then
            Beep
            oPayRec.MoveLast
        End If
    End If

ErrorHandler:
    If Err.Number = -2147217842 Then
        Resume Next
    End If

End Sub

'=========================================================
'Procudure Name: Movelast
'Synopsis:       Move Movelast Record
'e.g             Movelast RS-(your Recordset)
'========================================================
Public Sub MoveLast(ByRef oPayRec As ADODB.Recordset)
On Error GoTo ErrorHandler
    
    If oPayRec.RecordCount > 0 Then oPayRec.MoveLast

ErrorHandler:
    If Err.Number = -2147217842 Then
        Resume Next
    End If

End Sub
