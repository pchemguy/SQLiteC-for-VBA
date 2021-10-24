VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LiteMetaADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Retrieves metadata from an SQLite database"
'@Folder "SQLite.SQLiteDB"
'@ModuleDescription "Retrieves metadata from an SQLite database"
'@PredeclaredId
'@Exposed
'@IgnoreModule IndexedRecursiveDefaultMemberAccess, FunctionReturnValueDiscarded
'@IgnoreModule IndexedDefaultMemberAccess, ImplicitDefaultMemberAccess
'@IgnoreModule ProcedureNotUsed, ValueRequired
Option Explicit

Private Type TLiteMetaADO
    ADOTool As ILiteADO
    MetaSQL As LiteMetaSQL
End Type
Private this As TLiteMetaADO


Public Function Create(ByVal ADOTool As ILiteADO) As LiteMetaADO
    Dim Instance As LiteMetaADO
    Set Instance = New LiteMetaADO
    Instance.Init ADOTool
    Set Create = Instance
End Function


Friend Sub Init(ByVal ADOTool As ILiteADO)
    Guard.NullReference ADOTool
    Guard.NullReference ADOTool.AdoConnection
    Guard.EmptyString ADOTool.MainDB
    
    Set this.ADOTool = ADOTool
    Set this.MetaSQL = LiteMetaSQL.Create()
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
    DbSchema = this.ADOTool.GetAdoRecordset(this.MetaSQL.GetDbSchema(SchemaAlias)).GetRows
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
    DbSchema = this.ADOTool.GetAdoRecordset(this.MetaSQL.DbSchemaNoTriggers(SchemaAlias)).GetRows
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
    DbSchema = this.ADOTool.GetAdoRecordset(this.MetaSQL.Triggers(SchemaAlias)).GetRows
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
    TablesInfo = this.ADOTool.GetAdoRecordset(this.MetaSQL.Tables).GetRows

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
    SQLQuery = this.MetaSQL.Tables(SchemaAlias)
    GetTableList = ArrayLib.GetRow(this.ADOTool.GetAdoRecordset(SQLQuery).GetRows, 0)
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
    Set AdoRecordset = this.ADOTool.GetAdoRecordset(this.MetaSQL.TableColumnsEx(TableName))
    Dim TableSQL As String
    TableSQL = this.ADOTool.GetAdoRecordset(this.MetaSQL.TableSQL(TableName)).Fields("sql")
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
    Set AdoRecordset = this.ADOTool.GetAdoRecordset(this.MetaSQL.Indices)
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
    Databases = this.ADOTool.GetAdoRecordset(this.MetaSQL.Databases).GetRows
    Databases = ArrayLib.GetRow(Databases, 0)
    
    Dim DbIndex As Long
    For DbIndex = LBound(Databases) To UBound(Databases)
        Databases(DbIndex) = _
            "PRAGMA " & Databases(DbIndex) & ".synchronous='NORMAL';" & vbNewLine & _
            "PRAGMA " & Databases(DbIndex) & ".journal_mode='" & JournalMode & "'"
    Next DbIndex
    
    Dim SQLQuery As String
    SQLQuery = Join(Databases, ";" & vbNewLine) & ";"
    this.ADOTool.ExecuteNonQuery SQLQuery
End Sub


'@Description "Queries journal mode for the given database"
Public Function JournalModeGet(Optional ByVal SchemaAlias As String = "main") As String
Attribute JournalModeGet.VB_Description = "Queries journal mode for the given database"
    Dim SQLQuery As String
    SQLQuery = "PRAGMA " & SchemaAlias & ".journal_mode"
    JournalModeGet = this.ADOTool.GetScalar(SQLQuery)
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
        Databases = ArrayLib.GetRow(this.ADOTool.GetAdoRecordset(SQLQuery).GetRows, 0)
    End If
    
    Dim DbIndex As Long
    
    '''' The driver SQLite ODBC has a bug. It tries to do some statement
    '''' interpretation and fails, when multiple "PRAGMA" statements are inlcuded.
    For DbIndex = LBound(Databases) To UBound(Databases)
        If JournalModeUC = "WAL" Then
            SQLQuery = "PRAGMA [" & Databases(DbIndex) & "].synchronous = 'NORMAL'"
            this.ADOTool.ExecuteNonQuery SQLQuery
        End If
        SQLQuery = "PRAGMA [" & Databases(DbIndex) & "].journal_mode = '" & JournalMode & "'"
        this.ADOTool.ExecuteNonQuery SQLQuery
    Next DbIndex
End Sub
