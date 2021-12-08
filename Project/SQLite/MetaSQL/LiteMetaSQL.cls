VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LiteMetaSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "SQL queries for retrieving SQLite database metadata."
'@Folder "SQLite.MetaSQL"
'@ModuleDescription "SQL queries for retrieving SQLite database metadata."
'@PredeclaredId
'@Exposed
'@IgnoreModule DefaultMemberRequired, DefaultMemberRequired, DefaultMemberRequired, ProcedureNotUsed
Option Explicit

Private Type TLiteMetaSQLSchema
    Schema As String
End Type
Private this As TLiteMetaSQLSchema


'''' @ClassMethodStrict
'''' This method should only be used on the default instance
''''
'@Description "Default factory"
Public Function Create(Optional ByVal Schema As String = "main") As LiteMetaSQL
Attribute Create.VB_Description = "Default factory"
    Dim Instance As LiteMetaSQL
    Set Instance = New LiteMetaSQL
    Instance.Init Schema
    Set Create = Instance
End Function

Public Sub Init(Optional ByVal Schema As String = "main")
    this.Schema = Schema
End Sub

Public Property Get Schema() As String
    Schema = this.Schema
End Property

Public Property Let Schema(ByVal Value As String)
    this.Schema = Value
End Property

'''' ==================== DATABASE ==================== ''''

'@Description "Generates a query returning the list of attached databases"
Public Property Get Databases() As String
Attribute Databases.VB_Description = "Generates a query returning the list of attached databases"
    Databases = "SELECT name, file FROM pragma_database_list"
End Property

'''' @Proxy
'@Description "Generates a query returning all non-system database objects."
Public Function GetDbSchema(Optional ByVal Schema As String = vbNullString) As String
Attribute GetDbSchema.VB_Description = "Generates a query returning all non-system database objects."
    GetDbSchema = LiteMetaSQLIdxFK.DbSchema(IIf(Len(Schema) > 0, Schema, this.Schema))
End Function

'''' @Proxy
'@Description "Generates a query returning all non-system database objects, but triggers"
Public Function DbSchemaNoTriggers(Optional ByVal Schema As String = vbNullString) As String
Attribute DbSchemaNoTriggers.VB_Description = "Generates a query returning all non-system database objects, but triggers"
    DbSchemaNoTriggers = LiteMetaSQLIdxFK.DbSchemaNoTriggers(IIf(Len(Schema) > 0, Schema, this.Schema))
End Function

'''' @Proxy
'@Description "Generates a query returning triggers"
Public Function Triggers(Optional ByVal Schema As String = vbNullString) As String
Attribute Triggers.VB_Description = "Generates a query returning triggers"
    Triggers = LiteMetaSQLIdxFK.Triggers(IIf(Len(Schema) > 0, Schema, this.Schema))
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
    Tables = LiteMetaSQLIdxFK.Tables(IIf(Len(Schema) > 0, Schema, this.Schema))
End Function

'''' @Proxy
'@Description "Generates a query returning all foreing keys in the SQLite database"
Public Property Get ForeingKeys() As String
Attribute ForeingKeys.VB_Description = "Generates a query returning all foreing keys in the SQLite database"
    ForeingKeys = LiteMetaSQLIdxFK.ForeingKeys(this.Schema)
End Property

'''' @Proxy
'@Description "Generates a query returning all indices in the SQLite database"
Public Function Indices(Optional ByVal NonSys As Boolean = True) As String
Attribute Indices.VB_Description = "Generates a query returning all indices in the SQLite database"
    Indices = LiteMetaSQLIdxFK.Indices(this.Schema, NonSys)
End Function

'''' @Proxy
'''' See the called class for details
'@Description "Generates a query returning child columns for all foreing keys and corresponding indices."
Public Property Get FKChildIndices() As String
Attribute FKChildIndices.VB_Description = "Generates a query returning child columns for all foreing keys and corresponding indices."
    FKChildIndices = LiteMetaSQLIdxFK.FKChildIndices(this.Schema)
End Property

'''' @Proxy
'''' See the called class for details
'@Description "Generates a query returning similar indices."
Public Property Get SimilarIndices() As String
Attribute SimilarIndices.VB_Description = "Generates a query returning similar indices."
    SimilarIndices = LiteMetaSQLIdxFK.SimilarIndices(this.Schema)
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

'''' ==================== ENGINE ==================== ''''
'''' Engine related queries are classmethods ''''

'@Description "Generates query returning available SQLite collations"
Public Property Get Collations() As String
Attribute Collations.VB_Description = "Generates query returning available SQLite collations"
    Collations = "SELECT * FROM pragma_collation_list() AS collations ORDER BY name"
End Property

'@Description "Generates query returning compile options"
Public Property Get CompileOptions() As String
Attribute CompileOptions.VB_Description = "Generates query returning compile options"
    CompileOptions = "SELECT * FROM pragma_compile_options() AS compile_options"
End Property

'@Description "Generates query returning available SQLite functions"
Public Property Get Functions() As String
Attribute Functions.VB_Description = "Generates query returning available SQLite functions"
    Functions = "SELECT * FROM pragma_function_list() AS functions ORDER BY name"
End Property

'@Description "Generates query returning available SQLite functions count"
Public Property Get FunctionsCount() As String
Attribute FunctionsCount.VB_Description = "Generates query returning available SQLite functions count"
    FunctionsCount = "SELECT count(*) As counter FROM pragma_function_list()"
End Property

'@Description "Generates query returning available SQLite modules"
Public Property Get Modules() As String
Attribute Modules.VB_Description = "Generates query returning available SQLite modules"
    Modules = "SELECT * FROM pragma_module_list() AS modules ORDER BY name"
End Property

'@Description "Generates query returning available SQLite pragmas"
Public Property Get Pragmas() As String
Attribute Pragmas.VB_Description = "Generates query returning available SQLite pragmas"
    Pragmas = "SELECT * FROM pragma_pragma_list() AS pargmas ORDER BY name"
End Property

'@Description "Generates query returning SQLite version"
Public Property Get Version() As String
Attribute Version.VB_Description = "Generates query returning SQLite version"
    Version = "SELECT sqlite_version() AS version"
End Property
