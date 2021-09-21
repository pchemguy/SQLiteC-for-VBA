VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LiteADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Simplified ADODB routines used by SQLiteDB"
'@Folder "SQLiteDB"
'@ModuleDescription "Minimalistic ADODB wrapper used by SQLiteDB"
'@PredeclaredId
'@Exposed
Option Explicit

Implements ILiteADO

Private Type TLiteADO
    MainDB As String '''' Absolute file pathname to the main database.
    AdoConnection As ADODB.Connection
    AdoCommand As ADODB.Command
End Type
Private this As TLiteADO


'@DefaultMember
Public Function Create(ByVal Database As String, _
              Optional ByVal AllowNonExistent As Boolean = False) As ILiteADO
Attribute Create.VB_UserMemId = 0
    Dim Instance As LiteADO
    Set Instance = New LiteADO
    Instance.Init Database, AllowNonExistent
    Set Create = Instance
End Function


'''' Args:
''''   Database (string):
''''     Name of the database to be opened as the "main" SQLite database.
''''
''''   AllowNonExistent (boolean, optional, False):
''''     If True, Database may refer to a non-existent database, which will be
''''     created.
''''
'''' Note:
''''   VerifyOrGetDefaultPath should be called with two arguments max. If
''''   the second argument is not Empty, the third argument will be ignored.
''''   The first argument is allowed to refer to a non-existent file only if
''''   the third argument is True.
''''
Friend Sub Init(ByVal Database As String, _
       Optional ByVal AllowNonExistent As Boolean = False)
    If Database = ":memory:" Then
        this.MainDB = Database
    ElseIf AllowNonExistent Then
        this.MainDB = VerifyOrGetDefaultPath(Database, , AllowNonExistent)
    Else
        this.MainDB = VerifyOrGetDefaultPath(Database, Array("db", "sqlite"))
    End If
    Set this.AdoCommand = New ADODB.Command
    With this.AdoCommand
        .CommandType = adCmdText
        .CommandText = SQLiteSQLDbInfo.Create.Engine.Version
        .Prepared = True
        .ActiveConnection = ConnectionString()
        .ActiveConnection.CursorLocation = adUseClient
        Set this.AdoConnection = .ActiveConnection
    End With
End Sub


'''' Args:
''''   DatabaseName (string, optional, vbNullString):
''''     If provided, assume that it will be attached as an additional database.
''''     Only exsistent databases are allowed to be attached. A new persistent or
''''     in-memory database should be created as the "main" database via the
''''     factory method.
''''
'@Description "Constructs default SQLiteODBC connection string"
Private Function ConnectionString( _
            Optional ByVal DatabaseName As String = vbNullString) As String
Attribute ConnectionString.VB_Description = "Constructs default SQLiteODBC connection string"
    Dim Driver As String
    Driver = "SQLite3 ODBC Driver"

    Dim Database As String
    If Len(DatabaseName) > 0 Then
        Database = VerifyOrGetDefaultPath(DatabaseName, Array("db", "sqlite"))
    Else
        Database = this.MainDB
    End If

    Dim Options As String
    Options = "SyncPragma=NORMAL;FKSupport=True;"
        
    ConnectionString = "Driver=" & Driver & ";Database=" & Database & ";" & Options
End Function


'''' Opens, disconnects, and returns an ADODB.Recordset. If SQLQuery is provided,
'''' this.AdoCommand attribute is updated before initializing the Recordset.
''''
'''' Empty SQLQuery can be used, for example, with previously set parametrized query.
'''' Parameter values should be bound to this.AdoCommand before calling this method.
''''
'@Description "Returns disconnected Recordset"
Private Function GetAdoRecordset(Optional ByVal SQLQuery As String = vbNullString) As ADODB.Recordset
Attribute GetAdoRecordset.VB_Description = "Returns disconnected Recordset"
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = New ADODB.Recordset
    If Len(SQLQuery) > 0 Then this.AdoCommand.CommandText = SQLQuery
    With AdoRecordset
        Set .Source = this.AdoCommand
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .Open
        Set .ActiveConnection = Nothing
    End With

    Set GetAdoRecordset = AdoRecordset
End Function


'@Description "Returns disconnected Recordset"
Private Function GetScalar(Optional ByVal SQLQuery As String = vbNullString) As Variant
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = New ADODB.Recordset
    If Len(SQLQuery) > 0 Then this.AdoCommand.CommandText = SQLQuery
    With AdoRecordset
        Set .Source = this.AdoCommand
        .MaxRecords = 1
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Open
        Set .ActiveConnection = Nothing
    End With
    
    If AdoRecordset.RecordCount = 0 Then
        GetScalar = Empty
    Else
        GetScalar = AdoRecordset.Fields(0).Value
    End If
End Function


'@Description "Executes database modifying statement (UPDATE, DELETE, INSERT)"
Private Function ExecuteNonQuery(ByVal SQLQuery As String) As Long
Attribute ExecuteNonQuery.VB_Description = "Executes database modifying statement (UPDATE, DELETE, INSERT)"
    Dim RecordsAffected As Long
    With this.AdoCommand
        .CommandText = SQLQuery
        .Execute RecordsAffected, Options:=adExecuteNoRecords
    End With
    ExecuteNonQuery = RecordsAffected
End Function


'============================================================================='
'============================= ILiteADO INTERFACE ============================'
'============================================================================='


Private Property Get ILiteADO_MainDB() As String
    ILiteADO_MainDB = this.MainDB
End Property

Private Property Get ILiteADO_ConnectionString() As String
    ILiteADO_ConnectionString = ConnectionString(this.MainDB)
End Property

Private Property Get ILiteADO_AdoConnection() As ADODB.Connection
    Set ILiteADO_AdoConnection = this.AdoConnection
End Property

Private Property Get ILiteADO_AdoCommand() As ADODB.Command
    Set ILiteADO_AdoCommand = this.AdoCommand
End Property

Private Function ILiteADO_GetAdoRecordset(Optional ByVal SQLQuery As String = vbNullString) As ADODB.Recordset
    Set ILiteADO_GetAdoRecordset = GetAdoRecordset(SQLQuery)
End Function

Private Function ILiteADO_ExecuteNonQuery(ByVal SQLQuery As String) As Long
    ILiteADO_ExecuteNonQuery = ExecuteNonQuery(SQLQuery)
End Function

Private Function ILiteADO_GetScalar(Optional ByVal SQLQuery As String = vbNullString) As Variant
    ILiteADO_GetScalar = GetScalar(SQLQuery)
End Function
