VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MinADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Simplified ADODB routines used by SQLiteDB"
'@Folder "SQLiteDB"
'@ModuleDescription "Simplified ADODB routines used by SQLiteDB"
'@PredeclaredId
'@Exposed
Option Explicit

Private Type TMinADO
    MainDB As String '''' Absolute file pathname to the main database.
    AdoConnection As ADODB.Connection
    AdoCommand As ADODB.Command
End Type
Private this As TMinADO


'@DefaultMember
Public Function Create(ByVal Database As String, _
              Optional ByVal AllowNonExistent As Boolean = False) As MinADO
Attribute Create.VB_UserMemId = 0
    Dim Instance As MinADO
    Set Instance = New MinADO
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


Public Property Get MainDB() As String
    MainDB = this.MainDB
End Property


Public Property Get AdoConnection() As ADODB.Connection
    Set AdoConnection = this.AdoConnection
End Property


Public Property Get AdoCommand() As ADODB.Command
    Set AdoCommand = this.AdoCommand
End Property


'''' Args:
''''   DatabaseName (string, optional, vbNullString):
''''     If provided, assume that it will be attached as an additional database.
''''     Only exsistent databases are allowed to be attached. A new persistent or
''''     in-memory database should be created as the "main" database via the
''''     factory method.
''''
'@Description "Constructs default SQLiteODBC connection string"
Public Function ConnectionString( _
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
Public Function GetAdoRecordset(Optional ByVal SQLQuery As String = vbNullString) As ADODB.Recordset
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


'@Description "Executes database modifying statement (UPDATE, DELETE, INSERT)"
Public Sub ExecuteNonQuery(ByVal SQLQuery As String)
Attribute ExecuteNonQuery.VB_Description = "Executes database modifying statement (UPDATE, DELETE, INSERT)"
    this.AdoCommand.CommandText = SQLQuery
    this.AdoCommand.Execute Options:=adExecuteNoRecords
End Sub
