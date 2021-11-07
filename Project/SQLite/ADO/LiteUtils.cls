VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LiteUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Provides shortcuts for common operations."
'@Folder "SQLite.ADO"
'@ModuleDescription "Provides shortcuts for common operations."
'@PredeclaredId
'@IgnoreModule ValueRequired: False positive with ADO
'@IgnoreModule IndexedDefaultMemberAccess, FunctionReturnValueDiscarded
'@IgnoreModule ProcedureNotUsed
Option Explicit

Private Type TLiteUtils
    Exec As ILiteADO
    Meta As LiteMetaSQLSchema
End Type
Private this As TLiteUtils


'@DefaultMember
Public Function Create(ByVal ExecADO As ILiteADO) As LiteUtils
Attribute Create.VB_UserMemId = 0
    Dim Instance As LiteUtils
    Set Instance = New LiteUtils
    Instance.Init ExecADO
    Set Create = Instance
End Function


Friend Sub Init(ByVal ExecADO As ILiteADO)
    Guard.NullReference ExecADO
    Guard.EmptyString ExecADO.MainDB
    Set this.Exec = ExecADO
    Set this.Meta = LiteMetaSQLSchema.Create()
End Sub


'@Description "Attaches SQLite database to existing connection"
Public Sub AttachDatabase(ByVal DatabaseName As String, _
                 Optional ByVal DatabaseAlias As String = vbNullString)
Attribute AttachDatabase.VB_Description = "Attaches SQLite database to existing connection"
    this.Exec.ExecuteNonQuery SQLlib.Attach(DatabaseName, DatabaseAlias)
End Sub


'@Description "Detaches SQLite database from existing connection"
Public Sub DetachDatabase(ByVal DatabaseAlias As String)
Attribute DetachDatabase.VB_Description = "Detaches SQLite database from existing connection"
    this.Exec.ExecuteNonQuery SQLlib.Detach(DatabaseAlias)
End Sub


'@Description "Defer foreing keys"
Public Sub DeferForeignKeys()
Attribute DeferForeignKeys.VB_Description = "Defer foreing keys"
    this.Exec.ExecuteNonQuery SQLlib.DeferForeignKeys
End Sub


'''' Executes provided SQLQuery and prints returned Recordset as a table to
'''' 'immediate'. If OutputRange is provided, the returned Recordset is also
'''' placed on the referenced Excel Worksheet via the QueryTable feature.
''''
'@Description "'Debug.Print' for an SQL query"
Public Sub DebugPrintRecordset(ByVal SQLQuery As String, _
                      Optional ByVal OutputRange As Excel.Range = Nothing)
Attribute DebugPrintRecordset.VB_Description = "'Debug.Print' for an SQL query"
    Guard.EmptyString SQLQuery
    
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = this.Exec.GetAdoRecordset(SQLQuery)
    
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
    Databases = this.Exec.GetAdoRecordset(this.Meta.Databases).GetRows
    Databases = ArrayLib.GetRow(Databases, 0)
    
    Dim DbIndex As Long
    For DbIndex = LBound(Databases) To UBound(Databases)
        Databases(DbIndex) = _
            "PRAGMA " & Databases(DbIndex) & ".synchronous='NORMAL';" & vbNewLine & _
            "PRAGMA " & Databases(DbIndex) & ".journal_mode='" & JournalMode & "'"
    Next DbIndex
    
    Dim SQLQuery As String
    SQLQuery = Join(Databases, ";" & vbNewLine) & ";"
    this.Exec.ExecuteNonQuery SQLQuery
End Sub


'@Description "Queries journal mode for the given database"
Public Function JournalModeGet(Optional ByVal SchemaAlias As String = "main") As String
Attribute JournalModeGet.VB_Description = "Queries journal mode for the given database"
    Dim SQLQuery As String
    SQLQuery = "PRAGMA " & SchemaAlias & ".journal_mode"
    JournalModeGet = this.Exec.GetScalar(SQLQuery)
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
        Databases = ArrayLib.GetRow(this.Exec.GetAdoRecordset(SQLQuery).GetRows, 0)
    End If
    
    Dim DbIndex As Long
    
    '''' The driver SQLite ODBC has a bug. It tries to do some statement
    '''' interpretation and fails, when multiple "PRAGMA" statements are inlcuded.
    For DbIndex = LBound(Databases) To UBound(Databases)
        If JournalModeUC = "WAL" Then
            SQLQuery = "PRAGMA [" & Databases(DbIndex) & "].synchronous = 'NORMAL'"
            this.Exec.ExecuteNonQuery SQLQuery
        End If
        SQLQuery = "PRAGMA [" & Databases(DbIndex) & "].journal_mode = '" & JournalMode & "'"
        this.Exec.ExecuteNonQuery SQLQuery
    Next DbIndex
End Sub


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
Public Function CloneDb(ByVal DstDbName As String, ByVal SrcDbName As String) As ILiteADO
Attribute CloneDb.VB_Description = "Clones SQLite database."
    Dim SQLQuery As String
    Dim AdoRecordset As ADODB.Recordset
    Dim CheckResult As Boolean
    
    '''' 1) Check source integrity
    ''''      SQLiteDB reference is not saved, so the db is released at check exit.
    LiteACID(LiteADO(SrcDbName)).IntegrityADODB
    
    '''' 2) Attach destination db
    Dim dbm As LiteDB
    Set dbm = LiteDB(DstDbName, True)
    Debug.Assert Not dbm Is Nothing
    Debug.Print "-- Destination db is attached"
        
    '''' 3) Attach source db
    dbm.Util.AttachDatabase SrcDbName, "source"
    SQLQuery = this.Meta.Databases
    Set AdoRecordset = dbm.Exec.GetAdoRecordset(SQLQuery)
    CheckResult = (AdoRecordset.RecordCount = 2)
    Set AdoRecordset = Nothing
    Guard.Expression CheckResult, "LiteUtils/CloneDb", "Attach source db failed"
    Debug.Print "-- Source db is attached"
    
    '''' 4) Get schema without triggers and trigger schema
    Dim SchemaNoTriggersSQL As String
    SchemaNoTriggersSQL = dbm.Meta.GetDbSchemaNoTriggersSQL("source")
    Debug.Assert Len(SchemaNoTriggersSQL) > 0
    Dim TriggerSchemaSQL As String
    TriggerSchemaSQL = dbm.Meta.GetTriggersSQL("source")
    Debug.Print "-- Source schema is retrieved"
    
    '''' 5) Get table list
    Dim TableList As Variant
    TableList = dbm.Meta.GetTableList("source")
    Debug.Assert Not IsEmpty(TableList)
    Debug.Print "-- Source table list is retrieved"
    
    '''' 6) Clone schema without triggers
    dbm.Exec.ExecuteNonQuery SQLlib.FKStatus(False)
    dbm.Util.JournalModeSet "WAL", "main"
    dbm.ExecCI.AdoConnection.BeginTrans
    dbm.Exec.ExecuteNonQuery SchemaNoTriggersSQL
    dbm.ExecCI.AdoConnection.CommitTrans
    Debug.Assert UBound(TableList) = UBound(dbm.Meta.GetTableList)
    Debug.Print "-- Schema without triggers is cloned"
    
    '''' 7) Clone data
    Dim TableName As Variant
    dbm.ExecCI.AdoConnection.BeginTrans
    For Each TableName In TableList
        SQLQuery = SQLlib.CopyTableData("source", TableName)
        dbm.Exec.ExecuteNonQuery SQLQuery
    Next TableName
    dbm.ExecCI.AdoConnection.CommitTrans
    Debug.Print "-- Data is cloned"
    
    '''' 8) Clone trigger schema
    dbm.ExecCI.AdoConnection.BeginTrans
    dbm.Exec.ExecuteNonQuery TriggerSchemaSQL
    dbm.ExecCI.AdoConnection.CommitTrans
    dbm.Exec.ExecuteNonQuery SQLlib.FKStatus(True)
    Debug.Print "-- Triggers are cloned"
    
    '''' 9) Verify transfer and target db integrity
    ' TODO:
    '   Transfer verification (e.g., compare row counts for each table)
    dbm.Util.DetachDatabase "source"
    dbm.Exec.ExecuteNonQuery "ANALYZE"
    '@Ignore IndexedDefaultMemberAccess
    LiteACID(dbm.Exec).IntegrityADODB
    
    Set CloneDb = dbm.Exec
End Function
