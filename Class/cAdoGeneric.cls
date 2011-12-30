VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAdoGeneric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'IT Group Inc. 2005.09.23

Option Explicit
'******************************************************************************
'** Module.........: CAdoGeneric
'** Description....: Dbase Methodes and properties
'**                  Reference to ADO 2.5
'**
'**
'**
'** Version........: 1.0.0.A
'** Modifications..:
'** 1.0.1 - xxx , dd-mmm-yyyy
'**
'** Property             Data Type     Description
'** ------------------   ---------     --------------------------------------
'** Dbname               String        DataBase Name
'** ServerName           String        Server Name (SQL-Server)
'** Password             String        Password for ODBC and SQl-Server
'** USerID               string        The User for ODBC & SQL-Server
'**
'** Method(Public)       Description
'** ------------------   --------------------------------------
'** RunSQL               Execute SQL without recordset return
'** GetBatchRecord       Get a recordset for Later Read/Write (Batch)
'** SaveBatchRecord      Update recordset from (GetBatchRecord) in BatchMode
'** ReadOnlyRecord       Get a recordset for Read Only (Fast)
'** GetRecord            Get a Recordset for Read/Write
'** SaveRecord           Save a recordset
'** SetConnectType       The Conection Type (Access, ODBC, SQL-Server)
'**
'******************************************************************************

Private Const m_CLASSNAME = "CadoGeneric"
Private Const vbObjectError = 513
Private Const ERR_UPDATE_COLLISION = -2147217864
Private m_sDbname As String
Private m_sConnect As String
Private AdoConnect As ADODB.Connection
Private LocRecset As ADODB.Recordset
Private ADOBatch As ADODB.Recordset
Private m_ecType As Integer

Enum ecType
    [Access] = 0
    [SQLServer] = 1
    [ODBC] = 2
End Enum
Private mvarPassword As String
Private mvarUserID As String
Private mvarServerName As String

'**************************
'****    Properties    ****
'**************************

Public Property Let ServerName(ByVal vData As String)
'******************************************************************************
'** Description....: The Server Name SQL-Server
'** Author, date...: '  21-Jun-2000.
'******************************************************************************
    mvarServerName = vData
End Property

Public Property Get ServerName() As String
'******************************************************************************
'** Description....: Get the Server Name
'** Author, date...: '  21-Jun-2000.
'******************************************************************************
    ServerName = mvarServerName
End Property

Public Property Let Password(ByVal vData As String)
'******************************************************************************
'** Description....: Assigne le Password de Connection ODBC ou SQL
'** Author, date...: '  17-Jun-2000.
'******************************************************************************
    mvarPassword = vData
End Property

Public Property Get Password() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Password
    Password = mvarPassword
End Property

Public Property Let UserID(ByVal vData As String)
'******************************************************************************
'** Description....: Assigne le UserID pour Connection ODBC ou SQL
'** Author, date...: '  17-Jun-2000.
'******************************************************************************
    mvarUserID = vData
End Property

Public Property Get UserID() As String
'******************************************************************************
'** Description....: Retourne la valeur de cette proriété
'** Author, date...: '  17-Jun-2000.
'******************************************************************************
    UserID = mvarUserID
End Property

Public Property Let Dbname(ByVal sData As String)
'******************************************************************************
'** Description....: Assigne une valeur a cette proriété
'** Author, date...: ' , 12-Aug-1999.
'******************************************************************************
m_sDbname = sData
End Property

Public Property Get Dbname() As String
'******************************************************************************
'** Description....: Retourne la valeur de cette proriété
'** Author, date...: ' , 12-Aug-1999.
'******************************************************************************
Dbname = m_sDbname
End Property



'******************************
'****    Methodes Public   ****
'******************************

Private Sub Class_Initialize()

Set AdoConnect = New ADODB.Connection

End Sub

Private Sub Class_Terminate()
'*YL Do we need to close the connection ?
If Not AdoConnect Is Nothing Then
    Set AdoConnect = Nothing
End If

End Sub

Public Sub SetConnectType(ByVal vData As ecType)
'******************************************************************************
'** SubRoutine.....: SetConnectType
'**
'** Description....: Choose the Connection Type
'**
'** Author, date...: ' , 21-Jun-2000.
'**
'** Modifications..:
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'** vData               Enum       R     Access, ODBC, or SQL-Server connection
'******************************************************************************
    m_ecType = vData
End Sub

Public Function ReadOnlyRecord(ByVal strQuery As String) As ADODB.Recordset
'******************************************************************************
'** Function.......: ReadOnlyRecord
'**
'** Description....: Get a recordset for read only
'**
'**
'** Author, date...: ', 17-nov-1999.
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'** strQuery            String     R     The query
'******************************************************************************
On Error GoTo ErrorSection

Set LocRecset = New ADODB.Recordset

'**YL Connect and read Record
With LocRecset
    .ActiveConnection = RebuildConnection
    '*YL adUseClient cause need the RecordCount property
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockReadOnly
    .Source = strQuery
    .Open
End With

Set ReadOnlyRecord = LocRecset
Set LocRecset.ActiveConnection = Nothing

'********************
'Exit Point
'********************
ExitPoint:

Exit Function
'********************
'Error Section
'********************
ErrorSection:
If AdoConnect.Errors.Count > 0 Then
    ShowError AdoConnect.Errors(0).Number, AdoConnect.Errors(0).Description, "ReadOnlyRecord", m_CLASSNAME, vbLogEventTypeError
Else
    Select Case Err.Number
        Case Else
        ShowError Err.Number, Err.Description, "ReadOnlyRecord", m_CLASSNAME, vbLogEventTypeError
    End Select
End If
Resume ExitPoint

End Function

Public Sub RunSQL(ByVal strSQL As String)
'******************************************************************************
'** SubRoutine.....: RunSQL
'**
'** Description....: Execute a query only
'**
'**
'** Author, date...: ' , 17-Aug-1999.
'**
'** Modifications..:
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'** strSQL              String     R     The query
'******************************************************************************
On Error GoTo ErrorSection

With AdoConnect
    .ConnectionString = RebuildConnection
    .Open
    .Execute (strSQL)
End With

'********************
' B) Exit Point
'********************
ExitPoint:
AdoConnect.Close
Exit Sub
'********************
' C) Error Section
'********************
ErrorSection:
If AdoConnect.Errors.Count > 0 Then
    ShowError AdoConnect.Errors(0).Number, AdoConnect.Errors(0).Description, "RunSQL", m_CLASSNAME, vbLogEventTypeError
Else
    Select Case Err.Number
        Case Else
        ShowError Err.Number, Err.Description, "RunSQL", m_CLASSNAME, vbLogEventTypeError
    End Select
End If
Resume ExitPoint

End Sub


Public Function GetBatchRecord(ByVal strQuery As String) As ADODB.Recordset
'******************************************************************************
'** Function.......: GetBatchRecord
'**
'** Description....: Get a recordset for
'**                  Batchmode Read/Write operation
'**
'** Author, date...: ', 16-nov-1999.
'**
'** Modifications..:
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'**  strQuery           String     R      The Query
'**  GetBacthRecord     Recordset  W      The recordset result
'******************************************************************************
On Error GoTo ErrorSection

Set ADOBatch = New ADODB.Recordset
'**YL Open Recordset for batch mode
With ADOBatch
    .ActiveConnection = RebuildConnection
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset      'Need the RecordCount Property
    .LockType = adLockBatchOptimistic
    .Source = strQuery
    .Open
End With
Set GetBatchRecord = ADOBatch
Set ADOBatch.ActiveConnection = Nothing

'********************
'Exit Point
'********************
ExitPoint:
Exit Function
'********************
'Error Section
'********************
ErrorSection:
Select Case Err.Number
    Case Else
    ShowError Err.Number, Err.Description, "GetBatchRecord", m_CLASSNAME, vbLogEventTypeError
End Select
Resume ExitPoint
Set GetBatchRecord = Nothing
End Function

Public Function GetRecord(ByVal strQuery As String) As ADODB.Recordset
'******************************************************************************
'** Function.......: GetRecord
'**
'** Description....: Put a query into recordset for
'**                  Read/Write operation
'**
'** Author, date...: ', 16-nov-1999.
'**
'** Modifications..:
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'**  strQuery           String     R      The Query
'**  GetRecord          Recordset  W      The Recordset result
'******************************************************************************
On Error GoTo ErrorSection

Set LocRecset = New ADODB.Recordset
'**YL Open Recordset
With LocRecset
    .ActiveConnection = RebuildConnection
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Source = strQuery
    .Open
End With
Set GetRecord = LocRecset
'********************
'Exit Point
'********************
ExitPoint:
Set LocRecset.ActiveConnection = Nothing

Exit Function
'********************
'Error Section
'********************
ErrorSection:
Select Case Err.Number
    Case Else
    ShowError Err.Number, Err.Description, "GetRecord", m_CLASSNAME, vbLogEventTypeError
End Select
Resume ExitPoint
Set GetRecord = Nothing
End Function

Public Function SaveBatchRecord(ByVal AdoRS As ADODB.Recordset) As Boolean
'******************************************************************************
'** SubRoutine.....: SaveBatchRecord
'**
'** Description....: Save Recordset in Batch Mode
'**                  If no Error then Return TRUE
'**
'** Author, date...: ' , 17-nov-1999.
'**
'** Modifications..:
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'** AdoRS               Recordset   R    The Recordset to save
'** SaveBatchRecord     Boolean     W    TRUE = No Error
'******************************************************************************
On Error GoTo ErrorSection
'**YL Save recordset Marshal Mode
'*YL No error default
SaveBatchRecord = True

With AdoRS
    .MarshalOptions = adMarshalModifiedOnly
    .ActiveConnection = RebuildConnection
    .UpdateBatch
End With
Set AdoRS.ActiveConnection = Nothing

'********************
'Exit Point
'********************
ExitPoint:
Exit Function
'********************
'Error Section
'********************
ErrorSection:
Select Case Err.Number
    Case ERR_UPDATE_COLLISION
        '**YL Resync problem so try again
        Err.Clear
        With AdoRS
            .Resync adAffectAllChapters, adResyncUnderlyingValues
            .UpdateBatch
        End With
        Resume ExitPoint
    Case Else
    ShowError Err.Number, Err.Description, "SaveBatchRecord", m_CLASSNAME, vbLogEventTypeError
End Select
SaveBatchRecord = False
'** Error so cancel the batch
AdoRS.CancelBatch
Resume ExitPoint

End Function

Public Function SaveRecord(ByVal AdoRS As ADODB.Recordset) As Boolean
'******************************************************************************
'** SubRoutine.....: SaveRecord
'**
'** Description....: Save Recordset
'**                  If no Error then Return TRUE
'**
'** Author, date...: ', 17-nov-1999.
'**
'** Modifications..:
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'** AdoRS               Recordset   R    The Recordset to save
'** SaveRecord          Boolean     W    TRUE -> Ok else error
'******************************************************************************
On Error GoTo ErrorSection
'*YL No error default
SaveRecord = True

With AdoConnect
    .ConnectionString = RebuildConnection
    .Open
End With
Set AdoRS.ActiveConnection = AdoConnect

'** We start transaction
AdoConnect.BeginTrans
AdoRS.Update
'** Commit the transaction
AdoConnect.CommitTrans
'********************
'Exit Point
'********************
ExitPoint:
Set AdoRS.ActiveConnection = Nothing
AdoConnect.Close
Exit Function
'********************
'Error Section
'********************
ErrorSection:
If AdoConnect.Errors.Count > 0 Then
    ShowError AdoConnect.Errors(0).Number, AdoConnect.Errors(0).Description, "SaveRecord", m_CLASSNAME, vbLogEventTypeError
Else
    Select Case Err.Number
        Case Else
        ShowError Err.Number, Err.Description, "SaveRecord", m_CLASSNAME, vbLogEventTypeError
    End Select
End If
SaveRecord = False
AdoConnect.RollbackTrans
End Function

'******************************
'****    Methodes Private  ****
'******************************

Private Function RebuildConnection() As String
'******************************************************************************
'** Function.......: RebuildConnection
'**
'** Description....: En fonction du Type Connection desire
'**                  on construit la connection
'**
'** Author, date...: ' , 17-Jun-2000.
'**
'** Modifications..:
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'** RebuildConnection   String     W     La connection ADO
'******************************************************************************
On Error GoTo ErrorSection

Dim szLocalString As String

Select Case m_ecType
    Case ecType.Access
        '** If you have MDAC 2.5
        szLocalString = "Provider=Microsoft.jet.OLEDB.4.0;Data source=" & Dbname
        '** If you have MDAC 2.1
        'szLocalString = "Provider=Microsoft.jet.OLEDB.3.51;Data source=" & Dbname
    Case ecType.ODBC
        szLocalString = "Provider=MSDASQL.1;Password=" & Password & _
              ";Persist Security Info=True;User ID=" & UserID & ";Data Source=" & Dbname
    Case ecType.SQLServer
                        
        szLocalString = "Provider=SQLOLEDB.1;Password=" & Password & ";Persist Security Info=True;User ID=" & UserID & ";Initial Catalog=" & Dbname & ";Data Source=" & ServerName

End Select

RebuildConnection = szLocalString
'********************
'Exit Point
'********************
ExitPoint:
Exit Function
'********************
'Error Section
'********************
ErrorSection:
Select Case Err.Number
    Case Else
    ShowError Err.Number, Err.Description, "RebuildConnection", m_CLASSNAME, vbLogEventTypeError
End Select
Resume ExitPoint

End Function

'********************************
'****    Error's Handling    ****
'********************************

Private Sub ShowError(ErrorNumber As Long, ErrorMsg As String _
                      , ErrorModule As String, ErrorForm As String _
                     , LogEventType As Long, Optional ErrorInfo As Variant)
'******************************************************************************
'** Module.........: ShowError
'** Description....: This routine is used to show the current
'**                  error Message and LOG the error to a file.
'**
'** Author, date...: ', 16-nov-1999.
'**
'** Name                Type     Acces   Description
'** ------------------  -------  ------  --------------------------------------
'**  ErrorNumber         Long      R      Error Number
'**  ErrorMsg            String    R      Error Message
'**  ErrorModule         String    R      Module name where the error occured
'**  ErrorForm           String    R      Form Name where the error occured
'**  LogEventType        Long      R      Log event type (vbLogEventTypeError ,
'**                                       vbLogEventTypeWarning , vbLogEventTypeInformation)
'**  ErrorInfo           Variant   R      Additional error Information to Display
'**
'******************************************************************************
On Error GoTo ErrorSection
Dim ErrorTitle As String
Dim ErrorMessage As String

ErrorTitle = "ERROR - " & ErrorNumber & " - " & ErrorModule & " - " & ErrorForm
ErrorMessage = "ERROR  " & ErrorNumber & " - " & ErrorMsg

If Not IsMissing(ErrorInfo) Then
    ErrorMessage = ErrorMessage & vbCrLf & ErrorInfo
End If
App.LogEvent ErrorTitle & ": " & ErrorMessage, LogEventType
ExitPoint:
Exit Sub
ErrorSection:
Resume ExitPoint

End Sub







