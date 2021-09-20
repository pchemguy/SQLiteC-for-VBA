VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLiteDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Provides SQLite specific functionality"
'@Folder "SQLiteDB"
'@ModuleDescription "Provides SQLite specific functionality"
'@PredeclaredId
'@Exposed
'@IgnoreModule IndexedRecursiveDefaultMemberAccess, FunctionReturnValueDiscarded
'@IgnoreModule IndexedDefaultMemberAccess, ImplicitDefaultMemberAccess
'@IgnoreModule ValueRequired, ProcedureNotUsed
Option Explicit


Private Type TSQLiteDB
    MainDB As String '''' Absolute file pathname to the main database.
    AdoConnection As ADODB.Connection
    AdoCommand As ADODB.Command
    SQLInfo As SQLiteSQLDbInfo
End Type
Private this As TSQLiteDB


'''' @ClassMethodStrict
'''' This method should only be used on the default instance
''''
'@DefaultMember
'@Description "Default factory"
Public Function Create(ByVal Database As String, _
              Optional ByVal AllowNonExistent As Boolean = False) As SQLiteDB
Attribute Create.VB_Description = "Default factory"
Attribute Create.VB_UserMemId = 0
    Dim Instance As SQLiteDB
    Set Instance = New SQLiteDB
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
'@Description "Default constructor"
Public Sub Init(ByVal Database As String, _
       Optional ByVal AllowNonExistent As Boolean = False)
Attribute Init.VB_Description = "Default constructor"
    Set this.SQLInfo = SQLiteSQLDbInfo.Create
    If Database = ":memory:" Then
        this.MainDB = Database
    ElseIf AllowNonExistent Then
        this.MainDB = VerifyOrGetDefaultPath(Database, , AllowNonExistent)
    Else
        this.MainDB = VerifyOrGetDefaultPath(Database, Array("db", "sqlite"))
    End If
    Set this.AdoCommand = New ADODB.Command
    With AdoCommand
        .CommandType = adCmdText
        .CommandText = this.SQLInfo.Engine.Version
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


Public Property Get SQLInfo() As SQLiteSQLDbInfo
    Set SQLInfo = this.SQLInfo
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


'''' Open an ADODB.Recordset, disconnects it, and returns it. If SQLQuery is provided,
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


'''' Returns:
''''   String:
''''     SQL code, containing CREATE statements for all non-system database
''''     object. This code can be executed as a "non-query" against a blank
''''     "main" database to clone the source database schema.
''''
'@Description "Retrieves full SQLite database schema"
Public Function GetDbSchemaSQL(Optional ByVal SchemaAlias As String = "main") As String
Attribute GetDbSchemaSQL.VB_Description = "Retrieves full SQLite database schema"
    Dim DbSchema As Variant
    DbSchema = GetAdoRecordset(this.SQLInfo.GetDbSchema(SchemaAlias)).GetRows
    DbSchema = ArrayLib.GetRow(DbSchema, 0)
    GetDbSchemaSQL = Join(DbSchema, ";" & vbNewLine) & ";"
End Function


'''' Returns:
''''   String:
''''     SQL code, containing CREATE statements for all non-system database
''''     object, except for triggers. This code can be executed as a "non-query"
''''     against a blank "main" database to clone the source database schema.
''''
'@Description "Retrieves SQLite database schema, skipping triggers"
Public Function GetDbSchemaNoTriggersSQL(Optional ByVal SchemaAlias As String = "main") As String
Attribute GetDbSchemaNoTriggersSQL.VB_Description = "Retrieves SQLite database schema, skipping triggers"
    Dim DbSchema As Variant
    DbSchema = GetAdoRecordset(this.SQLInfo.DbSchemaNoTriggers(SchemaAlias)).GetRows
    DbSchema = ArrayLib.GetRow(DbSchema, 0)
    GetDbSchemaNoTriggersSQL = Join(DbSchema, ";" & vbNewLine) & ";"
End Function


'''' Returns:
''''   String:
''''     SQL code, containing CREATE statements for triggers. This code can be
''''     executed as a "non-query" against the "main" database to clone the
''''     source database trigger schema.
''''
'@Description "Retrieves SQLite database trigger schema"
Public Function GetTriggersSQL(Optional ByVal SchemaAlias As String = "main") As String
Attribute GetTriggersSQL.VB_Description = "Retrieves SQLite database trigger schema"
    Dim DbSchema As Variant
    DbSchema = GetAdoRecordset(this.SQLInfo.Triggers(SchemaAlias)).GetRows
    DbSchema = ArrayLib.GetRow(DbSchema, 1)
    GetTriggersSQL = Join(DbSchema, ";" & vbNewLine) & ";"
End Function


'''' Returns:
''''   Dictionary, containing table schema information:
''''     .Item("Name"):
''''       1D array of table names ordered by ROWID
''''     .Item("SQL"):
''''       1D array of corresponding CREATE TABLE statements
''''
'    Note:
''''   ADODB.Recordset.GetRows returns a column-wise 2D array
''''
'@Description "Retrieves table schema"
Public Property Get TableSchema() As Scripting.Dictionary
Attribute TableSchema.VB_Description = "Retrieves table schema"
    Dim TablesInfo As Variant
    TablesInfo = GetAdoRecordset(this.SQLInfo.Tables).GetRows

    Dim SchemaDic As Scripting.Dictionary
    Set SchemaDic = New Scripting.Dictionary
    SchemaDic.CompareMode = TextCompare
    
    Const NAME_ROW_INDEX As Long = 0
    Const SQL_ROW_INDEX As Long = 1
    SchemaDic("Name") = ArrayLib.GetRow(TablesInfo, NAME_ROW_INDEX)
    SchemaDic("SQL") = ArrayLib.GetRow(TablesInfo, SQL_ROW_INDEX)
    
    Set TableSchema = SchemaDic
End Property


'''' Returns:
''''   1D array, containing table names
''''
'@Description "Retrieves table names"
Public Function GetTableList(Optional ByVal SchemaAlias As String = "main") As Variant
Attribute GetTableList.VB_Description = "Retrieves table names"
    Dim SQLQuery As String
    SQLQuery = this.SQLInfo.Tables(SchemaAlias)
    GetTableList = ArrayLib.GetRow(GetAdoRecordset(SQLQuery).GetRows, 0)
End Function


'''' Retrieves full fields metadata from an SQLite database table.
''''
'''' SQLite's pragrma function does not return complete field metadata information.
'''' This routine uses a Pragma function to get the list of fields with core
'''' metadata. Then it retrieves table schema and parses the field clauses to
'''' collect the remaining metadata. The returned recordset is modified with
'''' missing metadata.
''''
'''' Returns:
''''   ADODB.Recordset, containing detailed field information.
''''
'@Description "Retrieves full fields metadata from an SQLite database table."
Public Function GetTableColumnsEx(ByVal TableName As String) As ADODB.Recordset
Attribute GetTableColumnsEx.VB_Description = "Retrieves full fields metadata from an SQLite database table."
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = GetAdoRecordset(this.SQLInfo.TableColumnsEx(TableName))
    Dim TableSQL As String
    TableSQL = GetAdoRecordset(this.SQLInfo.TableSQL(TableName)).Fields("sql")
    Debug.Print TableSQL
    Dim CreateTableLines As Variant
    CreateTableLines = Split(TableSQL, vbLf)

    Dim re As RegExp
    Set re = New RegExp
    re.Pattern = "\w*"
    Dim Matches As MatchCollection
     
    Dim FieldSQL As String
    Dim ClauseBeg As Long
    Dim ClauseEnd As Long
    Dim ColumnIndex As Long
    With AdoRecordset
        .MoveFirst
        For ColumnIndex = 1 To .RecordCount
            FieldSQL = CreateTableLines(ColumnIndex)
            If InStr(1, FieldSQL, " UNIQUE ", vbBinaryCompare) Then
                .Fields("unique") = 1
            Else
                .Fields("unique") = Null
            End If
            ClauseBeg = InStr(1, FieldSQL, " CHECK(", vbBinaryCompare)
            If ClauseBeg > 0 Then
                ClauseBeg = ClauseBeg + Len(" CHECK(")
                ClauseEnd = InStrRev(FieldSQL, ")")
                .Fields("check") = Mid$(FieldSQL, ClauseBeg, ClauseEnd - ClauseBeg)
            End If
            ClauseBeg = InStr(1, FieldSQL, " COLLATE ", vbBinaryCompare)
            If ClauseBeg > 0 Then
                ClauseBeg = ClauseBeg + Len(" COLLATE ")
                Set Matches = re.Execute(Mid$(FieldSQL, ClauseBeg))
                .Fields("collate") = Matches(0).Value
            End If
            .MoveNext
        Next ColumnIndex
        .MoveFirst
    End With
    
    Set GetTableColumnsEx = AdoRecordset
End Function


'''' Returns:
''''   Dictionary, containing index schema (no auto indices) information:
''''     .Item("Name"):
''''       1D array of index names ordered by ROWID
''''     .Item("SQL"):
''''       1D array of corresponding CREATE INDEX statements
''''     .Item("Unique"):
''''       1D array of corresponding "unique" flags
''''
'@Description "Retrieves index schema"
Public Property Get IndexSchema() As Scripting.Dictionary
Attribute IndexSchema.VB_Description = "Retrieves index schema"
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = GetAdoRecordset(this.SQLInfo.Indices)
    Dim IndicesInfo As Variant
    IndicesInfo = AdoRecordset.GetRows

    Dim FieldMap As Scripting.Dictionary
    Set FieldMap = New Scripting.Dictionary
    FieldMap.CompareMode = TextCompare
    Dim FieldIndex As Long
    
    For FieldIndex = 0 To AdoRecordset.Fields.Count - 1
        FieldMap(AdoRecordset.Fields(FieldIndex).Name) = FieldIndex
    Next FieldIndex
    
    Dim SchemaDic As Scripting.Dictionary
    Set SchemaDic = New Scripting.Dictionary
    SchemaDic.CompareMode = TextCompare
    
    With ArrayLib
        SchemaDic("Name") = .GetRow(IndicesInfo, FieldMap("idx_name"))
        SchemaDic("SQL") = .GetRow(IndicesInfo, FieldMap("sql"))
        SchemaDic("Unique") = .GetRow(IndicesInfo, FieldMap("unique"))
    End With
    
    Set IndexSchema = SchemaDic
End Property


'''' This routines queries the database to get the list of databases attached
'''' to the current connection. For each database, two pragmas are generated:
'''' one sets sync mode to 'NORMAL', the other sets journal mode according to
'''' the provided argument. 'WAL' mode is more efficient and should be used
'''' in most circumstances unless ACID transactions across multiple attached
'''' databases are required, in which case 'DELETE' or 'TRUNCATE' should be
'''' used.
''''
'@Description "Sets NORMAL sync and journal mode to WAL or DELETE for all attached dbs"
Public Sub ToggleJournalMode(Optional ByVal JournalMode As String = "WAL")
Attribute ToggleJournalMode.VB_Description = "Sets NORMAL sync and journal mode to WAL or DELETE for all attached dbs"
    Dim Databases As Variant
    Databases = GetAdoRecordset(this.SQLInfo.Databases).GetRows
    Databases = ArrayLib.GetRow(Databases, 0)
    
    Dim DbIndex As Long
    For DbIndex = LBound(Databases) To UBound(Databases)
        Databases(DbIndex) = _
            "PRAGMA " & Databases(DbIndex) & ".synchronous='NORMAL';" & vbNewLine & _
            "PRAGMA " & Databases(DbIndex) & ".journal_mode='" & JournalMode & "'"
    Next DbIndex
    
    Dim SQLQuery As String
    SQLQuery = Join(Databases, ";" & vbNewLine) & ";"
    ExecuteNonQuery SQLQuery
End Sub


'@Description "Queries journal mode for the given database"
Public Function JournalModeGet(Optional ByVal SchemaAlias As String = "main") As String
Attribute JournalModeGet.VB_Description = "Queries journal mode for the given database"
    Dim SQLQuery As String
    SQLQuery = "PRAGMA " & SchemaAlias & ".journal_mode"
    JournalModeGet = GetAdoRecordset(SQLQuery).Fields("journal_mode")
End Function


'''' Args:
''''   JournalMode (string, optional, "WAL"):
''''     New journal mode. 'WAL' mode is more efficient and should be used in
''''     most circumstances unless ACID transactions across multiple attached
''''     databases are required, in which case 'DELETE' or 'TRUNCATE' should be
''''     used.
''''
''''   SchemaAlias (string, optional, "main"):
''''     Database alias, for which journal mode should be set. If "ALL" is provided,
''''     the mode will be set for all attached databases.
''''
''''   This routine also sets sync mode to 'NORMAL'.
''''
'@Description "Sets journal mode for the given database"
Public Sub JournalModeSet(Optional ByVal JournalMode As String = "WAL", _
                          Optional ByVal SchemaAlias As String = "main")
Attribute JournalModeSet.VB_Description = "Sets journal mode for the given database"
    Const JOURNAL_MODES As String = "DELETE | TRUNCATE | PERSIST | MEMORY | WAL | OFF"
    Dim JournalModeUC As String
    JournalModeUC = UCase$(JournalMode)
    Dim CheckResult As Boolean
    CheckResult = CBool(InStr("| " & JOURNAL_MODES & " |", "| " & JournalModeUC & " |"))
    Guard.Expression CheckResult, "SQLiteDB/JournalModeSet", _
        "Invalid journal mode: '" & JournalMode & "'. Valid values are: " _
        & JOURNAL_MODES
    
    Dim SQLQuery As String
    Dim Databases As Variant
    If UCase$(SchemaAlias) <> "ALL" Then
        Databases = Array(SchemaAlias)
    Else
        SQLQuery = "SELECT name FROM pragma_database_list"
        Databases = ArrayLib.GetRow(GetAdoRecordset(SQLQuery).GetRows, 0)
    End If
    
    Dim DbIndex As Long
    
    '''' The driver SQLite ODBC has a bug. It tries to do some statement
    '''' interpretation and fails, when multiple "PRAGMA" statements are inlcuded.
    For DbIndex = LBound(Databases) To UBound(Databases)
        If JournalModeUC = "WAL" Then
            SQLQuery = "PRAGMA [" & Databases(DbIndex) & "].synchronous = 'NORMAL'"
            ExecuteNonQuery SQLQuery
        End If
        SQLQuery = "PRAGMA [" & Databases(DbIndex) & "].journal_mode = '" & JournalMode & "'"
        ExecuteNonQuery SQLQuery
    Next DbIndex
End Sub


'''' Executes provided SQLQuery and prints returned Recordset as a table to
'''' 'immediate'. If OutputRange is provided, the returned Recordset is also
'''' placed on the referenced Excel Worksheet via the QueryTable feature.
''''
'@Description "'Debug.Print' for an SQL query"
Public Sub DebugPrintRecordset(ByVal SQLQuery As String, _
                      Optional ByVal OutputRange As Excel.Range = Nothing)
    Guard.EmptyString SQLQuery
    
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = GetAdoRecordset(SQLQuery)
    
    If Not OutputRange Is Nothing Then
        ADOlib.RecordsetToQT AdoRecordset, OutputRange
    End If
    
    Dim FieldIndex As Long
    Dim FieldNames() As String
    ReDim FieldNames(1 To AdoRecordset.Fields.Count)
    For FieldIndex = LBound(FieldNames) To UBound(FieldNames)
        FieldNames(FieldIndex) = AdoRecordset.Fields(FieldIndex - 1).Name
    Next FieldIndex
    
    Dim Table As String
    Table = Join(FieldNames, vbTab) & vbNewLine & _
            AdoRecordset.GetString(, , vbTab, vbNewLine)
    Debug.Print Table
End Sub


'@Description "Attaches SQLite database to existing connection"
Public Sub AttachDatabase(ByVal DatabaseName As String, Optional ByVal DatabaseAlias As String = vbNullString)
Attribute AttachDatabase.VB_Description = "Attaches SQLite database to existing connection"
    ExecuteNonQuery SQLlib.Attach(DatabaseName, DatabaseAlias)
End Sub


'@Description "Detaches SQLite database from existing connection"
Public Sub DetachDatabase(ByVal DatabaseAlias As String)
Attribute DetachDatabase.VB_Description = "Detaches SQLite database from existing connection"
    ExecuteNonQuery SQLlib.Detach(DatabaseAlias)
End Sub


'@Description "Defer foreing keys"
Public Sub DeferForeignKeys()
Attribute DeferForeignKeys.VB_Description = "Defer foreing keys"
    ExecuteNonQuery SQLlib.DeferForeignKeys
End Sub


'@Description "Runs SQLite database access checks."
Public Function CheckAccess() As Boolean

End Function


'''' The database to be checked should be attached as the "main" and only
'''' database via the SQLiteDB factory. While integrity/FK check should run
'''' against the specified schema, if provided, I had issues with getting
'''' expected behavior. Prefer to run checks when a single database being
'''' checked is attached as "main".
'''' 1) Verify that existing file is a database and exactly one db is attached.
'''' 2) Run integrity check and verify that it returns "ok"
'''' 3) Run foreign key check and verify that nothing is returned.
''''
'''' Examples:
''''   >>> ?SQLiteDB("").CheckDB
''''   -- Connection check passed
''''   -- Integrity check passed for: '<Path>\SQLiteDB.db'
''''   True
''''
''''   >>> ?SQLiteDB("ICfailFKCfail.db").CheckDB
''''   -- Connection check passed
''''   Error:
''''       message: Source database failed integrity check
''''
''''   >>> ?SQLiteDB("ICokFKCfail.db").CheckDB
''''   -- Connection check passed
''''   Error:
''''       message: Source database failed foreign key check
''''
'@Description "Runs SQLite database integrity checks."
Public Function CheckIntegrity() As Boolean
    Const SQLITE_NOTADB_MESSAGE As String = "file is not a database (26)"
    
    Dim SQLQuery As String
    Dim CheckResult As Boolean
    Dim AdoRecordset As ADODB.Recordset

    '''' 1) Verify that existing file is a database and only one db is atttached
    SQLQuery = this.SQLInfo.Databases
    Dim AttachedDbCount
    Dim ErrNumber As Long
    ErrNumber = 0
    Dim ErrSource As String
    Dim ErrDescription As String
    
    On Error Resume Next
    AttachedDbCount = GetAdoRecordset(SQLQuery).RecordCount
    With Err
        If .Number <> 0 Then
            ErrNumber = .Number
            ErrSource = .Source
            ErrDescription = .Description
        End If
    End With
    On Error GoTo 0
    
    If ErrNumber <> 0 Then
        If (ErrNumber = ErrNo.AdoInvalidFileFormatErr) _
                And (ErrDescription = SQLITE_NOTADB_MESSAGE) Then
            Err.Raise ErrNo.AdoInvalidFileFormatErr, "SQLiteDB", _
                      "File is not a database"
        Else
            With Err
                .Raise ErrNumber, ErrSource, ErrDescription
            End With
        End If
    End If
    If AttachedDbCount <> 1 Then
        Err.Raise ErrNo.CustomErr, "SQLiteDB", _
                  "Please have exactly one database attached before the check."
    End If
    
    '''' 2) "PRAGMA integrity_check"
    SQLQuery = this.SQLInfo.CheckIntegrity
    Set AdoRecordset = GetAdoRecordset(SQLQuery)
    CheckResult = (AdoRecordset.Fields("integrity_check") = "ok")
    Guard.ExpressionErr CheckResult, ErrNo.IntegrityCheckErr, _
                        "SQLiteDB", "Integrity check failed"
    
    '''' 3) "PRAGMA foreign_key_check"
    SQLQuery = this.SQLInfo.CheckFKs
    Set AdoRecordset = GetAdoRecordset(SQLQuery)
    CheckResult = (AdoRecordset.RecordCount = 0)
    Guard.ExpressionErr CheckResult, ErrNo.ConsistencyCheckErr, _
                        "SQLiteDB", "Foreign key check failed"

    Debug.Print "-- Integrity check passed for: '" & this.MainDB & "'"
    CheckIntegrity = True
End Function


'''' The database to be checked should be attached as the "main" and only
'''' database via the SQLiteDB factory. While integrity/FK check should run
'''' against the specified schema, if provided, I had issues with getting
'''' expected behavior. Prefer to run checks when a single database being
'''' checked is attached as "main". This routine first runs base checks:
'''' 1) Attach an in-memory database and verify that the database count is 2;
'''' 2) Remove in-memory database and verify that the database count is 1
'''' 3) Run integrity check and verify that it returns "ok"
'''' 4) Run foreign key check and verify that nothing is returned.
''''
'''' Examples:
''''   >>> ?SQLiteDB("").CheckDB
''''   -- Connection check passed
''''   -- Integrity check passed for: '<Path>\SQLiteDB.db'
''''   True
''''
''''   >>> ?SQLiteDB("ICfailFKCfail.db").CheckDB
''''   -- Connection check passed
''''   Error:
''''       message: Source database failed integrity check
''''
''''   >>> ?SQLiteDB("ICokFKCfail.db").CheckDB
''''   -- Connection check passed
''''   Error:
''''       message: Source database failed foreign key check
''''
'@Ignore FunctionReturnValueAlwaysDiscarded
'@Description "Runs SQLite database integrity checks."
Public Function CheckDb() As Boolean
Attribute CheckDb.VB_Description = "Runs SQLite database integrity checks."
    Dim SQLQuery As String
    Dim CheckResult As Boolean
    Dim AdoRecordset As ADODB.Recordset
    
    '''' 1) Attach in-memory db (baseline check)
    AttachDatabase ":memory:", "memory"
    SQLQuery = this.SQLInfo.Databases
    Set AdoRecordset = GetAdoRecordset(SQLQuery)
    CheckResult = (AdoRecordset.RecordCount = 2) And _
                  (AdoRecordset.Fields("name") = "main")
    Set AdoRecordset = Nothing
    
    '''' 2) Detach in-memory db
    DetachDatabase "memory"
    SQLQuery = this.SQLInfo.Databases
    Set AdoRecordset = GetAdoRecordset(SQLQuery)
    CheckResult = CheckResult And (AdoRecordset.RecordCount = 1) And _
                                  (AdoRecordset.Fields("name") = "main")
    Set AdoRecordset = Nothing
    
    Guard.Expression CheckResult, "SQLiteDB/CheckDb", "Failed connection check"
    Debug.Print "-- Connection check passed"
    
    '''' 3) "PRAGMA integrity_check"
    SQLQuery = this.SQLInfo.CheckIntegrity
    Set AdoRecordset = GetAdoRecordset(SQLQuery)
    CheckResult = (AdoRecordset.Fields("integrity_check") = "ok")
    Guard.ExpressionErr CheckResult, ErrNo.IntegrityCheckErr, _
                        "SQLiteDB/CloneDb", "Integrity check failed"
    Set AdoRecordset = Nothing
    
    '''' 4) "PRAGMA foreign_key_check"
    SQLQuery = this.SQLInfo.CheckFKs
    Set AdoRecordset = GetAdoRecordset(SQLQuery)
    CheckResult = (AdoRecordset.RecordCount = 0)
    Guard.ExpressionErr CheckResult, ErrNo.ConsistencyCheckErr, _
                        "SQLiteDB/CloneDb", "Foreign key check failed"
    Set AdoRecordset = Nothing

    Debug.Print "-- Integrity check passed for: '" & this.MainDB & "'"
    CheckDb = True
End Function


'''' @ClassMethodStrict
'''' This method should only be used on the default instance
''''
'''' Clones an SQLite database.
''''
'''' Triggers cannot be disabled in SQLite, so trigger schema should be cloned
'''' separately after the all data is transfered.
'''' Defer_foreign_keys does not behave expectedly (FK violation during data
'''' transfer step). Prefer disabling foreign keys for the duration of process.
'''' Present implmentation does not clone the ROWID column for tables with
'''' separate hidden ROWID column. The use of such column should be avoided
'''' anyway, as its value can change at any time. In practice, either
'''' INTEGER PRIMARY KEY AUTOINCREMENT should be defined to serve as ROWID or
'''' the table should be defined as "WITHOUT ROWID".
''''
'''' Args:
''''   DstDbName (string):
''''     Name of the new clone database, referring to either ":memory:" or
''''     non-existent file to be created. Initial steps:
''''     1) Run integrity checks.
''''     2) Attach the destination database as the "main" to a new SQLiteDB
''''        instance use "True" as the second argument to the factory to enable
''''        file creation.
''''
''''   SrcDbName (string):
''''     Name of the database to be cloned referring to an existing file or an
''''     empty string, if fallback checks can pick it up.
''''     3) Attach as an additional database with alias "source" to the SQLiteDB
''''        instance from step (2).
''''     4) Retrieve source schema without triggers and trigger schema.
''''     5) Retrieve source table list.
''''
''''  6) Set journal mode to 'WAL'; Disable foreign keys; Start transaction; Execute schema;
''''     Commit transaction
''''  7) Start transaction; Clone data; Commit transaction
''''  8) Start transaction; Execute trigger schema; Commit transaction; Enable foreign keys
''''  9) Verify transfer and run integrity check on the destination database.
''''
'@Description "Clones SQLite database."
'@Ignore FunctionReturnValueAlwaysDiscarded
Public Function CloneDb(ByVal DstDbName As String, ByVal SrcDbName As String) As SQLiteDB
Attribute CloneDb.VB_Description = "Clones SQLite database."
    Dim SQLQuery As String
    Dim AdoRecordset As ADODB.Recordset
    Dim CheckResult As Boolean
    
    '''' 1) Check source integrity
    ''''      SQLiteDB reference is not saved, so the db is released at check exit.
    SQLiteDB(SrcDbName).CheckDb
    
    '''' 2) Attach destination db
    Dim DbManager As SQLiteDB
    Set DbManager = SQLiteDB(DstDbName, True)
    Debug.Assert Not DbManager Is Nothing
    Debug.Print "-- Destination db is attached"
        
    '''' 3) Attach source db
    DbManager.AttachDatabase SrcDbName, "source"
    SQLQuery = this.SQLInfo.Databases
    Set AdoRecordset = DbManager.GetAdoRecordset(SQLQuery)
    CheckResult = (AdoRecordset.RecordCount = 2)
    Set AdoRecordset = Nothing
    Guard.Expression CheckResult, "SQLiteDB/CloneDb", "Attach source db failed"
    Debug.Print "-- Source db is attached"
    
    '''' 4) Get schema without triggers and trigger schema
    Dim SchemaNoTriggersSQL As String
    SchemaNoTriggersSQL = DbManager.GetDbSchemaNoTriggersSQL("source")
    Debug.Assert Len(SchemaNoTriggersSQL) > 0
    Dim TriggerSchemaSQL As String
    TriggerSchemaSQL = DbManager.GetTriggersSQL("source")
    Debug.Print "-- Source schema is retrieved"
    
    '''' 5) Get table list
    Dim TableList As Variant
    TableList = DbManager.GetTableList("source")
    Debug.Assert Not IsEmpty(TableList)
    Debug.Print "-- Source table list is retrieved"
    
    '''' 6) Clone schema without triggers
    DbManager.ExecuteNonQuery SQLlib.FKStatus(False)
    DbManager.JournalModeSet "WAL", "main"
    DbManager.AdoConnection.BeginTrans
    DbManager.ExecuteNonQuery SchemaNoTriggersSQL
    DbManager.AdoConnection.CommitTrans
    Debug.Assert UBound(TableList) = UBound(DbManager.GetTableList)
    Debug.Print "-- Schema without triggers is cloned"
    
    '''' 7) Clone data
    Dim TableName As Variant
    DbManager.AdoConnection.BeginTrans
    For Each TableName In TableList
        SQLQuery = SQLlib.CopyTableData("source", TableName)
        DbManager.ExecuteNonQuery SQLQuery
    Next TableName
    DbManager.AdoConnection.CommitTrans
    Debug.Print "-- Data is cloned"
    
    '''' 8) Clone trigger schema
    DbManager.AdoConnection.BeginTrans
    DbManager.ExecuteNonQuery TriggerSchemaSQL
    DbManager.AdoConnection.CommitTrans
    DbManager.ExecuteNonQuery SQLlib.FKStatus(True)
    Debug.Print "-- Triggers are cloned"
    
    '''' 9) Verify transfer and target db integrity
    ' TODO:
    '   Transfer verification (e.g., compare row counts for each table)
    DbManager.DetachDatabase "source"
    DbManager.ExecuteNonQuery "ANALYZE"
    DbManager.CheckDb
    
    Set CloneDb = DbManager
End Function
