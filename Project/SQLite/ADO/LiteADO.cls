VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LiteADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Minimalistic ADODB wrapper used by SQLiteDB"
'@Folder "SQLite.ADO"
'@ModuleDescription "Minimalistic ADODB wrapper used by SQLiteDB"
'@PredeclaredId
'@Exposed
'@IgnoreModule IndexedDefaultMemberAccess, FunctionReturnValueDiscarded
Option Explicit

Implements ILiteADO

Private Type TLiteADO
    MainDB As String '''' Absolute file pathname to the main database.
    ConnectionString As String
    AdoConnection As ADODB.Connection
    AdoCommand As ADODB.Command
End Type
Private this As TLiteADO


'@DefaultMember
Public Function Create(ByVal Database As String, _
              Optional ByVal AllowNonExistent As Variant = False) As ILiteADO
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
''''   AllowNonExistent (Variant, optional, False):
''''     If True, Database may refer to a non-existent database, which will be
''''     created. If Empty, path resolution will be skipped.
''''
'''' Note:
''''   VerifyOrGetDefaultPath should be called with two arguments max. If
''''   the second argument is not Empty, the third argument will be ignored.
''''   The first argument is allowed to refer to a non-existent file only if
''''   the third argument is True.
''''
Friend Sub Init(ByVal Database As String, _
       Optional ByVal AllowNonExistent As Variant = False)
    Set this.AdoCommand = New ADODB.Command
    With this.AdoCommand
        .CommandType = adCmdText
        .CommandText = LiteMetaSQL.Version
        .Prepared = True
        .ActiveConnection = ConnectionString(Database, AllowNonExistent)
        .ActiveConnection.CursorLocation = adUseClient
        Set this.AdoConnection = .ActiveConnection
    End With
End Sub


'@Description "Generate ILiteADO instance from and existing ADODB.Connection."
Public Function FromConnection(ByVal db As ADODB.Connection) As ILiteADO
Attribute FromConnection.VB_Description = "Generate ILiteADO instance from and existing ADODB.Connection."
    Dim Instance As LiteADO
    Set Instance = New LiteADO
    Instance.InitFromConnection db
    Set FromConnection = Instance
End Function


'@Description "Generate ILiteADO instance from and existing ADODB.Connection."
Friend Sub InitFromConnection(ByVal db As ADODB.Connection)
Attribute InitFromConnection.VB_Description = "Generate ILiteADO instance from and existing ADODB.Connection."
    Guard.NullReference db
    
    Set this.AdoConnection = db
    Set this.AdoCommand = New ADODB.Command
    With this.AdoCommand
        .CommandType = adCmdText
        .Prepared = True
        Set .ActiveConnection = db
    End With
    this.MainDB = GetScalar(SQLlib.AttachedDbPaths("main"))
    this.AdoCommand.CommandText = LiteMetaSQL.Version
End Sub


Public Property Get AdoConnection() As ADODB.Connection
    Set AdoConnection = this.AdoConnection
End Property


Public Property Get AdoCommand() As ADODB.Command
    Set AdoCommand = this.AdoCommand
End Property


'''' Args:
''''   DatabaseName (string):
''''     File pathname of the main database
''''   AllowNonExistent (boolean, optional, false):
''''     If false, creation of new database will not be allowed.
''''
'@Description "Constructs default SQLiteODBC connection string"
Private Function ConnectionString(ByVal Database As String, _
                         Optional ByVal AllowNonExistent As Variant = False _
                         ) As String
Attribute ConnectionString.VB_Description = "Constructs default SQLiteODBC connection string"
    Dim PathCheck As LiteFSCheck
    Set PathCheck = LiteFSCheck(Database, AllowNonExistent)
    If PathCheck.ErrNumber <> 0 Then PathCheck.Raise
    Debug.Assert (Len(PathCheck.DatabasePathName) > 0 And Database <> ":blank:") Or _
                 (Len(PathCheck.DatabasePathName) = 0 And Database = ":blank:")
    
    Dim FilePathName As String
    FilePathName = PathCheck.DatabasePathName
    this.MainDB = FilePathName
    
    Dim Driver As String
    Driver = "SQLite3 ODBC Driver"

    Dim Options As String
    Options = "SyncPragma=NORMAL;FKSupport=True;"
    If Not (AllowNonExistent Or PathCheck.IsSpecial) Then _
        Options = Options & "NoCreat=True;"
        
    this.ConnectionString = "Driver=" & Driver & ";Database=" & FilePathName & ";" & Options
    ConnectionString = this.ConnectionString
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
        .LockType = adLockBatchOptimistic
        .Open
        Set .ActiveConnection = Nothing
    End With

    Set GetAdoRecordset = AdoRecordset
End Function


'@Description "Returns disconnected Recordset"
Private Function GetScalar(Optional ByVal SQLQuery As String = vbNullString) As Variant
Attribute GetScalar.VB_Description = "Returns disconnected Recordset"
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
    
    '@Ignore ValueRequired: false positive
    If AdoRecordset.RecordCount = 0 Then
        GetScalar = Empty
    Else
        '@Ignore IndexedDefaultMemberAccess
        GetScalar = AdoRecordset.Fields(0).Value
    End If
End Function


'@Description "Executes database modifying statement (UPDATE, DELETE, INSERT)"
Private Function ExecuteNonQuery(ByVal SQLQuery As String) As Long
Attribute ExecuteNonQuery.VB_Description = "Executes database modifying statement (UPDATE, DELETE, INSERT)"
    Dim RecordsAffected As Long: RecordsAffected = 0 '''' RD Workaround
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
    ILiteADO_ConnectionString = this.ConnectionString
End Property

Private Function ILiteADO_GetScalar(ByVal SQLQuery As String, _
                           Optional ByVal QueryParameters As Variant = Empty _
                           ) As Variant
    ILiteADO_GetScalar = GetScalar(SQLQuery)
End Function

Private Function ILiteADO_GetAdoRecordset(ByVal SQLQuery As String, _
                                 Optional ByVal QueryParameters As Variant = Empty _
                                 ) As ADODB.Recordset
    Set ILiteADO_GetAdoRecordset = GetAdoRecordset(SQLQuery)
End Function

Private Function ILiteADO_ExecuteNonQuery(ByVal SQLQuery As String, _
                                 Optional ByVal QueryParameters As Variant = Empty _
                                 ) As Long
    ILiteADO_ExecuteNonQuery = ExecuteNonQuery(SQLQuery)
End Function

Private Sub ILiteADO_Begin(Optional ByVal Mode As SQLiteTxnType = SQLITE_TXN_DEFERRED)
    If Mode = SQLITE_TXN_DEFERRED Then
        this.AdoConnection.BeginTrans
    Else
        Dim TxMode As String
        Select Case Mode
            Case SQLITE_TXN_IMMEDIATE
                TxMode = " IMMEDIATE"
            Case SQLITE_TXN_EXCLUSIVE
                TxMode = " EXCLUSIVE"
            Case Else
                Err.Raise ErrNo.InvalidParameterErr, "LiteADO", _
                    "Invalid transaction mode: <" & CStr(Mode) & ">"
        End Select
        ExecuteNonQuery "BEGIN" & TxMode
    End If
End Sub

Private Sub ILiteADO_Commit()
    this.AdoConnection.CommitTrans
End Sub

Private Sub ILiteADO_SavePoint(ByVal SavePointName As String)
    Guard.EmptyString SavePointName
    ExecuteNonQuery "SAVEPOINT " & SavePointName
End Sub

Private Sub ILiteADO_ReleasePoint(ByVal SavePointName As String)
    Guard.EmptyString SavePointName
    ExecuteNonQuery "RELEASE " & SavePointName
End Sub

Private Sub ILiteADO_Rollback(Optional ByVal SavePointName As String = vbNullString)
    If Len(SavePointName) = 0 Then
        this.AdoConnection.RollbackTrans
    Else
        ExecuteNonQuery "ROLLBACK " & SavePointName
    End If
End Sub
