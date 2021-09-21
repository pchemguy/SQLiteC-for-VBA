VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LiteMetaSQLEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "SQL queries for retrieving information about the engine configuration and available features."
'@Folder "SQLiteDB.Introspection"
'@ModuleDescription "SQL queries for retrieving information about the engine configuration and available features."
'@PredeclaredId
'@Exposed
'@IgnoreModule ProcedureNotUsed
'''' All methods in this module are class methods and can be safely called on the default instance
'''' @ClassModule
Option Explicit


'@Description "Generates query returning available SQLite collations"
Public Property Get Collations() As String
Attribute Collations.VB_Description = "Generates query returning available SQLite collations"
    Collations = "SELECT * FROM pragma_collation_list AS collations ORDER BY name"
End Property


'@Description "Generates query returning compile options"
Public Property Get CompileOptions() As String
Attribute CompileOptions.VB_Description = "Generates query returning compile options"
    CompileOptions = "SELECT * FROM pragma_compile_options AS compile_options"
End Property


'@Description "Generates query returning available SQLite functions"
Public Property Get Functions() As String
Attribute Functions.VB_Description = "Generates query returning available SQLite functions"
    Functions = "SELECT * FROM pragma_function_list AS functions ORDER BY name"
End Property


'@Description "Generates query returning available SQLite modules"
Public Property Get Modules() As String
Attribute Modules.VB_Description = "Generates query returning available SQLite modules"
    Modules = "SELECT * FROM pragma_module_list AS modules ORDER BY name"
End Property


'@Description "Generates query returning available SQLite pragmas"
Public Property Get Pragmas() As String
Attribute Pragmas.VB_Description = "Generates query returning available SQLite pragmas"
    Pragmas = "SELECT * FROM pragma_pragma_list AS pargmas ORDER BY name"
End Property


'@Description "Generates query returning SQLite version"
Public Property Get Version() As String
Attribute Version.VB_Description = "Generates query returning SQLite version"
    Version = "SELECT sqlite_version() AS version"
End Property
