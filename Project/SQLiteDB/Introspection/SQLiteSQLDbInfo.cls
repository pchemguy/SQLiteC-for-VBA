VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLiteSQLDbInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "SQL queries for retrieving SQLite database metadata."
'@Folder "SQLiteDB.Introspection"
'@ModuleDescription "SQL queries for retrieving SQLite database metadata."
'@PredeclaredId
'@Exposed
'@IgnoreModule DefaultMemberRequired, DefaultMemberRequired, ProcedureNotUsed
Option Explicit

Private Type TSQLiteSQLDbInfo
    Schema As String
    Engine As SQLiteSQLEngineInfo
End Type
Private this As TSQLiteSQLDbInfo


'''' @ClassMethodStrict
'''' This method should only be used on the default instance
''''
'@DefaultMember
'@Description "Default factory"
Public Function Create(Optional ByVal Schema As String = "main") As SQLiteSQLDbInfo
Attribute Create.VB_Description = "Default factory"
Attribute Create.VB_UserMemId = 0
    Dim Instance As SQLiteSQLDbInfo
    Set Instance = New SQLiteSQLDbInfo
    Instance.Init Schema
    Set Create = Instance
End Function


Public Sub Init(Optional ByVal Schema As String = "main")
    this.Schema = Schema
    Set this.Engine = SQLiteSQLEngineInfo
End Sub


'@Description "Exposes SQLiteSQLEngineInfo introspection queries"
Public Property Get Engine() As SQLiteSQLEngineInfo
Attribute Engine.VB_Description = "Exposes SQLiteSQLEngineInfo introspection queries"
    Set Engine = this.Engine
End Property


'@Description "Generates a query returning the list of attached databases"
Public Property Get Databases() As String
Attribute Databases.VB_Description = "Generates a query returning the list of attached databases"
    Databases = "SELECT name, file FROM pragma_database_list"
End Property


'''' @Proxy
'@Description "Generates a query returning all non-system database objects."
Public Function GetDbSchema(Optional ByVal Schema As String = vbNullString) As String
Attribute GetDbSchema.VB_Description = "Generates a query returning all non-system database objects."
    GetDbSchema = SQLiteSQLDbIdxFK.DbSchema(IIf(Len(Schema) > 0, Schema, this.Schema))
End Function


'''' @Proxy
'@Description "Generates a query returning all non-system database objects, but triggers"
Public Function DbSchemaNoTriggers(Optional ByVal Schema As String = vbNullString) As String
Attribute DbSchemaNoTriggers.VB_Description = "Generates a query returning all non-system database objects, but triggers"
    DbSchemaNoTriggers = SQLiteSQLDbIdxFK.DbSchemaNoTriggers(IIf(Len(Schema) > 0, Schema, this.Schema))
End Function


'''' @Proxy
'@Description "Generates a query returning triggers"
Public Function Triggers(Optional ByVal Schema As String = vbNullString) As String
Attribute Triggers.VB_Description = "Generates a query returning triggers"
    Triggers = SQLiteSQLDbIdxFK.Triggers(IIf(Len(Schema) > 0, Schema, this.Schema))
End Function


'''' For some reason, running SELECT * FROM <schema>.pragma_integrity_check
'''' with several attached databases gives the result as if <schema> is
'''' ignored and all attached databases are checked. Prefer to run this
'''' check when the only attached database is the one being checked.
'@Description "Generates a query running integrity check."
Public Property Get CheckIntegrity() As String
Attribute CheckIntegrity.VB_Description = "Generates a query running integrity check."
    CheckIntegrity = "SELECT * FROM pragma_integrity_check"
End Property


'''' For some reason, running SELECT * FROM <schema>.pragma_foreign_key_check
'''' with several attached databases gives the result as if <schema> is
'''' ignored and all attached databases are checked. Prefer to run this
'''' check when the only attached database is the one being checked.
'@Description "Generates a query running integrity check."
Public Property Get CheckFKs() As String
Attribute CheckFKs.VB_Description = "Generates a query running integrity check."
    CheckFKs = "SELECT * FROM pragma_foreign_key_check"
End Property


'''' @Proxy
'@Description "Generates a query returning database tables."
Public Function Tables(Optional ByVal Schema As String = vbNullString) As String
Attribute Tables.VB_Description = "Generates a query returning database tables."
    Tables = SQLiteSQLDbIdxFK.Tables(IIf(Len(Schema) > 0, Schema, this.Schema))
End Function


'''' @Proxy
'@Description "Generates a query returning all foreing keys in the SQLite database"
Public Property Get ForeingKeys() As String
Attribute ForeingKeys.VB_Description = "Generates a query returning all foreing keys in the SQLite database"
    ForeingKeys = SQLiteSQLDbIdxFK.ForeingKeys(this.Schema)
End Property


'''' @Proxy
'@Description "Generates a query returning all indices in the SQLite database"
Public Function Indices(Optional ByVal NonSys As Boolean = True) As String
Attribute Indices.VB_Description = "Generates a query returning all indices in the SQLite database"
    Indices = SQLiteSQLDbIdxFK.Indices(this.Schema, NonSys)
End Function


'''' @Proxy
'''' See the called class for details
'@Description "Generates a query returning child columns for all foreing keys and corresponding indices."
Public Property Get FKChildIndices() As String
Attribute FKChildIndices.VB_Description = "Generates a query returning child columns for all foreing keys and corresponding indices."
    FKChildIndices = SQLiteSQLDbIdxFK.FKChildIndices(this.Schema)
End Property


'''' @Proxy
'''' See the called class for details
'@Description "Generates a query returning similar indices."
Public Property Get SimilarIndices() As String
Attribute SimilarIndices.VB_Description = "Generates a query returning similar indices."
    SimilarIndices = SQLiteSQLDbIdxFK.SimilarIndices(this.Schema)
End Property


'@Description "Generates a query returning table's columns."
Public Function TableColumns(ByVal TableName As String) As String
Attribute TableColumns.VB_Description = "Generates a query returning table's columns."
    Guard.EmptyString TableName
    TableColumns = "SELECT * " & _
                   "FROM " & this.Schema & ".pragma_table_xinfo('" & TableName & "')"
End Function


'@Description "Generates a query returning table's columns with placeholder columns."
Public Function TableColumnsEx(ByVal TableName As String) As String
Attribute TableColumnsEx.VB_Description = "Generates a query returning table's columns with placeholder columns."
    Guard.EmptyString TableName
    TableColumnsEx = "SELECT * , 0 AS [unique], '' as [check], '' as [collate] " & _
                     "FROM " & this.Schema & ".pragma_table_info('" & TableName & "')"
End Function


'@Description "Generates a query returning table's SQL."
Public Function TableSQL(ByVal TableName As String) As String
Attribute TableSQL.VB_Description = "Generates a query returning table's SQL."
    Guard.EmptyString TableName
    TableSQL = "SELECT sql " & _
               "FROM sqlite_master " & _
               "WHERE type = 'table' AND name = '" & TableName & "'"
End Function


'@Description "Generates a query returning table's foreign keys."
Public Function TableForeingKeys(ByVal TableName As String) As String
Attribute TableForeingKeys.VB_Description = "Generates a query returning table's foreign keys."
    TableForeingKeys = "SELECT * " & _
                       "FROM " & this.Schema & ".pragma_foreign_key_list('" & TableName & "')"
End Function
