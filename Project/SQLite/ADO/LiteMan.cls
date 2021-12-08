VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LiteMan"
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

#If VBA7 Then
Private Declare PtrSafe Function SQLGetInstalledDrivers Lib "ODBCCP32" ( _
    ByVal lpszBuf As String, ByVal cbBufMax As Long, ByRef pcbBufOut As Long) As Long
#Else
Private Declare Function SQLGetInstalledDrivers Lib "ODBCCP32" ( _
    ByVal lpszBuf As String, ByVal cbBufMax As Long, ByRef pcbBufOut As Long) As Long
#End If

Private Type TLiteMan
    ExecADO As ILiteADO
    MetaADO As LiteMetaADO
    MetaSQL As LiteMetaSQL
End Type
Private this As TLiteMan


'''' Args:
''''   DbPathName (string):
''''     String describing the target database:
''''       * relative or absolute file pathname
''''       * ':memory:' - anonymous in-memory database
''''       * ':temp:'/':tmp:' - file db in the Temp folder with random name
''''       * ':blank:' - anonymous file-based db
''''
''''   AllowNonExistent (variant, optional, Empty):
''''     Controls additional path preprocessing and creatability
''''       * Empty                  - db file must exist, no path resolution
''''       * Not (Empty or Boolean) - new db file is ok, no path resolution
''''       * True/False             - new db file is based on the value
''''                                  and path resolution is enabled
''''
''''   ConnOptions (Variant, optional, Empty):
''''     Either a string or dictionary of ODBC options. If a string is provided,
''''     it is used as is. If a dictionary is provided, the containing options
''''     are added to / override the default options.
''''
'@DefaultMember
Public Function Create(ByVal Database As String, _
              Optional ByVal AllowNonExistent As Variant = True, _
              Optional ByVal ConnOptions As Variant = Empty) As LiteMan
Attribute Create.VB_UserMemId = 0
    Dim Instance As LiteMan
    Set Instance = New LiteMan
    Instance.Init Database, AllowNonExistent, ConnOptions
    Set Create = Instance
End Function

Friend Sub Init(ByVal Database As String, _
       Optional ByVal AllowNonExistent As Variant = True, _
       Optional ByVal ConnOptions As Variant = Empty)
    With this
        Set .ExecADO = NewDB(Database, AllowNonExistent, ConnOptions)
        Set .MetaADO = LiteMetaADO(.ExecADO)
        Set .MetaSQL = .MetaADO.MetaSQL
    End With
End Sub

Friend Function NewDB(ByVal Database As String, _
                      ByVal AllowNonExistent As Variant, _
             Optional ByVal ConnOptions As Variant = Empty) As ILiteADO
    Dim PathCheck As LiteFSCheck
    Set PathCheck = LiteFSCheck(Database, AllowNonExistent)
    With PathCheck
        If .ErrNumber <> 0 Then .Raise
        Dim FilePathName As String
        FilePathName = .DatabasePathName
        Debug.Assert (Len(FilePathName) > 0 And Database <> ":blank:") _
                  Or (Len(FilePathName) = 0 And Database = ":blank:")
        Set NewDB = LiteADO(.DatabasePathName, AllowNonExistent Or .IsSpecial, _
                            ConnOptions)
    End With
End Function

Public Property Get ExecADO() As ILiteADO
    Set ExecADO = this.ExecADO
End Property

Public Property Get MetaADO() As LiteMetaADO
    Set MetaADO = this.MetaADO
End Property

Public Property Get MetaSQL() As LiteMetaSQL
    Set MetaSQL = this.MetaADO.MetaSQL
End Property

'@Description "Attaches SQLite database to existing connection"
Public Sub AttachDatabase(ByVal DbName As String, _
                 Optional ByVal DbAlias As String = vbNullString)
Attribute AttachDatabase.VB_Description = "Attaches SQLite database to existing connection"
    Dim SQLQuery As String
    SQLQuery = SQLlib.Attach(NewDB(DbName, False).MainDB, DbAlias)
    this.ExecADO.ExecuteNonQuery SQLQuery
End Sub

'@Description "Detaches SQLite database from existing connection"
Public Sub DetachDatabase(ByVal DatabaseAlias As String)
Attribute DetachDatabase.VB_Description = "Detaches SQLite database from existing connection"
    this.ExecADO.ExecuteNonQuery SQLlib.Detach(DatabaseAlias)
End Sub

'@Description "Defer foreing keys"
Public Sub DeferForeignKeys()
Attribute DeferForeignKeys.VB_Description = "Defer foreing keys"
    this.ExecADO.ExecuteNonQuery SQLlib.DeferForeignKeys
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
    Set AdoRecordset = this.ExecADO.GetAdoRecordset(SQLQuery)
    
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


'@Description "Queries journal mode for the given database"
Public Function JournalModeGet(Optional ByVal SchemaAlias As String = "main") As String
Attribute JournalModeGet.VB_Description = "Queries journal mode for the given database"
    Dim SQLQuery As String
    SQLQuery = "PRAGMA " & SchemaAlias & ".journal_mode"
    JournalModeGet = this.ExecADO.GetScalar(SQLQuery)
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
        Databases = ArrayLib.GetRow(this.ExecADO.GetAdoRecordset(SQLQuery).GetRows, 0)
    End If
    
    Dim DbIndex As Long
    
    '''' The SQLite ODBC driver appears to have a bug. It tries to do some statement
    '''' interpretation and fails, when multiple "PRAGMA" statements are inlcuded.
    For DbIndex = LBound(Databases) To UBound(Databases)
        If JournalModeUC = "WAL" Then
            SQLQuery = "PRAGMA [" & Databases(DbIndex) & "].synchronous = 'NORMAL'"
            this.ExecADO.ExecuteNonQuery SQLQuery
        End If
        SQLQuery = "PRAGMA [" & Databases(DbIndex) & "].journal_mode = '" & JournalMode & "'"
        this.ExecADO.ExecuteNonQuery SQLQuery
    Next DbIndex
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
Public Sub JournalModeToggle(Optional ByVal JournalMode As String = "WAL")
Attribute JournalModeToggle.VB_Description = "Sets NORMAL sync and journal mode to WAL or DELETE for all attached dbs"
    Dim Databases As Variant
    Databases = this.ExecADO.GetAdoRecordset(this.MetaSQL.Databases).GetRows
    Databases = ArrayLib.GetRow(Databases, 0)
    
    Dim DbIndex As Long
    For DbIndex = LBound(Databases) To UBound(Databases)
        Databases(DbIndex) = _
            "PRAGMA " & Databases(DbIndex) & ".synchronous='NORMAL';" & vbNewLine & _
            "PRAGMA " & Databases(DbIndex) & ".journal_mode='" & JournalMode & "'"
    Next DbIndex
    
    Dim SQLQuery As String
    SQLQuery = Join(Databases, ";" & vbNewLine) & ";"
    this.ExecADO.ExecuteNonQuery SQLQuery
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
''''   6) Set journal mode to 'WAL'; Disable foreign keys; Start transaction; Execute schema;
''''      Commit transaction
''''   7) Start transaction; Clone data; Commit transaction
''''   8) Start transaction; Execute trigger schema; Commit transaction; Enable foreign keys
''''   9) Verify transfer and run integrity check on the destination database.
''''
'''' Returns:
''''   Database manager for the newly created database clone.
''''
'@Description "Clones SQLite database."
Public Function CloneDb(ByVal DstDbName As String, ByVal SrcDbName As String) As LiteMan
Attribute CloneDb.VB_Description = "Clones SQLite database."
    Dim SQLQuery As String
    Dim AdoRecordset As ADODB.Recordset
    Dim CheckResult As Boolean
    
    '''' 1) Check source integrity
    ''''      DB reference is not saved, so the db is released at check exit.
    LiteACID(NewDB(SrcDbName, False)).IntegrityADODB
    
    '''' 2) Attach destination db
    Dim dbm As LiteMan
    Set dbm = LiteMan(DstDbName, True)
    Debug.Assert Not dbm Is Nothing
    Debug.Print "-- Destination db is attached"
    Dim ExecADO As ILiteADO
    Set ExecADO = dbm.ExecADO
    Dim MetaADO As LiteMetaADO
    Set MetaADO = dbm.MetaADO
        
    '''' 3) Attach source db
    dbm.AttachDatabase SrcDbName, "source"
    SQLQuery = dbm.MetaSQL.Databases
    Set AdoRecordset = ExecADO.GetAdoRecordset(SQLQuery)
    CheckResult = (AdoRecordset.RecordCount = 2)
    Set AdoRecordset = Nothing
    Guard.Expression CheckResult, "LiteMan/CloneDb", "Attach source db failed"
    Debug.Print "-- Source db is attached"
    
    With MetaADO
        '''' 4) Get schema without triggers and trigger schema
        Dim SchemaNoTriggersSQL As String
        SchemaNoTriggersSQL = .GetDbSchemaNoTriggersSQL("source")
        Debug.Assert Len(SchemaNoTriggersSQL) > 0
        Dim TriggerSchemaSQL As String
        TriggerSchemaSQL = .GetTriggersSQL("source")
        Debug.Print "-- Source schema is retrieved"
        
        '''' 5) Get table list
        Dim TableList As Variant
        TableList = .GetTableList("source")
        Debug.Assert Not IsEmpty(TableList)
        Debug.Print "-- Source table list is retrieved"
    End With
    
    With ExecADO
        '''' 6) Clone schema without triggers
        .ExecuteNonQuery SQLlib.FKStatus(False)
        dbm.JournalModeSet "WAL", "main"
        .Begin
        .ExecuteNonQuery SchemaNoTriggersSQL
        .Commit
        Debug.Assert UBound(TableList) = UBound(MetaADO.GetTableList)
        Debug.Print "-- Schema without triggers is cloned"
    
        '''' 7) Clone data
        Dim TableName As Variant
        .Begin
        For Each TableName In TableList
            SQLQuery = SQLlib.CopyTableData("source", TableName)
            .ExecuteNonQuery SQLQuery
        Next TableName
        .Commit
        Debug.Print "-- Data is cloned"
    
        '''' 8) Clone trigger schema
        .Begin
        .ExecuteNonQuery TriggerSchemaSQL
        .Commit
        .ExecuteNonQuery SQLlib.FKStatus(True)
        Debug.Print "-- Triggers are cloned"
    End With
    
    '''' 9) Verify transfer and target db integrity
    ' TODO:
    '   Transfer verification (e.g., compare row counts for each table)
    dbm.DetachDatabase "source"
    ExecADO.ExecuteNonQuery "ANALYZE"
    '@Ignore IndexedDefaultMemberAccess
    LiteACID(ExecADO).IntegrityADODB
    
    Set CloneDb = dbm
End Function


'''' This function attempts to confirm that the standard registry key for the
'''' SQLite3ODBC driver is present and that the file driver exists. No attempt
'''' is made to verify its usability.
''''
'''' Attempt to determine environment (native X32onX32 or X64onX64) or X32onX64.
'''' If successfull, try retrieving SQLite3ODBC driver file pathname from the
'''' standard registry key (adjusted to the type of environment, if necessary).
'''' If successful, adjust path to the type of environment, if necessary, and
'''' check if file driver exists. If successful, return true, or false otherwise.
''''
'@Description "Checks if SQLite3ODBC diver is available."
Public Function SQLite3ODBCDriverCheck() As Boolean
Attribute SQLite3ODBCDriverCheck.VB_Description = "Checks if SQLite3ODBC diver is available."
    Const SQLITE3_ODBC_NAME As String = "SQLite3 ODBC Driver"
    
    '''' Check if SQLGetInstalledDrivers contains the standard SQLite3ODBC driver
    '''' description. Fail if not found.
    Dim Buffer As String
    Buffer = String(2000, vbNullChar)
    Dim ActualSize As Long: ActualSize = 0 '''' RD ByRef workaround
    Dim Result As Boolean
    Result = SQLGetInstalledDrivers(Buffer, Len(Buffer) * 2, ActualSize)
    Debug.Assert Result = True
    Result = InStr(Replace(Left$(Buffer, ActualSize - 1), vbNullChar, vbLf), _
                   SQLITE3_ODBC_NAME)
    If Not Result Then GoTo DRIVER_NOT_FOUND:
    
    Dim ODBCINSTPrefix As String
    Dim EnvArch As EnvArchEnum
    EnvArch = GetEnvX32X64Type()
    Select Case EnvArch
        Case ENVARCH_NATIVE
            ODBCINSTPrefix = "HKLM\SOFTWARE\ODBC\ODBCINST.INI\"
        Case ENVARCH_32ON64
            ODBCINSTPrefix = "HKLM\SOFTWARE\WOW6432Node\ODBC\ODBCINST.INI\"
        Case ENVARCH_NOTSUP
            Logger.Logg "Failed to determine Win/Office architecture or " & _
                        "unsupported.", , DEBUGLEVEL_ERROR
            SQLite3ODBCDriverCheck = False
            Exit Function
    End Select
    
    '''' Query standard ODBCINST.INI registry keys
    '@Ignore SelfAssignedDeclaration
    Dim wsh As New IWshRuntimeLibrary.WshShell
    Dim SQLite3ODBCDriverInstalled As Boolean
    Dim RegPath As String
    RegPath = ODBCINSTPrefix & "ODBC Drivers\" & SQLITE3_ODBC_NAME
    On Error Resume Next
        SQLite3ODBCDriverInstalled = (wsh.RegRead(RegPath) = "Installed")
        If Not SQLite3ODBCDriverInstalled Then GoTo DRIVER_NOT_FOUND:
    On Error GoTo 0
    RegPath = ODBCINSTPrefix & SQLITE3_ODBC_NAME & "\Driver"
    Dim SQLite3ODBCDriverPath As String
    On Error Resume Next
        SQLite3ODBCDriverPath = wsh.RegRead(RegPath)
        If Len(SQLite3ODBCDriverPath) = 0 Then GoTo DRIVER_NOT_FOUND:
    On Error GoTo 0
    Const SYSTEM_NATIVE As String = "System32"
    Const SYSTEM_32ON64 As String = "SysWOW64"
    If EnvArch = ENVARCH_32ON64 Then
        SQLite3ODBCDriverPath = _
            Replace(SQLite3ODBCDriverPath, SYSTEM_NATIVE, SYSTEM_32ON64)
    End If
    
    '''' Check if driver file exists
    '@Ignore SelfAssignedDeclaration
    Dim fso As New IWshRuntimeLibrary.FileSystemObject
    If Not fso.FileExists(SQLite3ODBCDriverPath) Then GoTo DRIVER_NOT_FOUND:
    
    Logger.Logg "SQLite3ODBC driver appears to be available.", , DEBUGLEVEL_INFO
    SQLite3ODBCDriverCheck = True
    Exit Function
    
DRIVER_NOT_FOUND:
    Logger.Logg "Failed to verify SQLite3ODBC driver availability", , DEBUGLEVEL_ERROR
    SQLite3ODBCDriverCheck = False
    Exit Function
End Function
