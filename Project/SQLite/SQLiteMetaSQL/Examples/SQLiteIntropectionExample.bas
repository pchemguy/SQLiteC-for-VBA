Attribute VB_Name = "SQLiteIntropectionExample"
'@Folder "SQLite.SQLiteMetaSQL.Examples"
'@IgnoreModule VariableNotUsed, IndexedDefaultMemberAccess
Option Explicit
Option Private Module
Option Compare Text


'@EntryPoint
'@Description "Collects SQLite engine information and ouputs via QueryTables and to 'immediate'"
Private Sub Engine()
Attribute Engine.VB_Description = "Collects SQLite engine information and ouputs via QueryTables and to 'immediate'"
    Dim SourceDb As String
    SourceDb = ":memory:"
    Dim DbManager As LiteDB
    Set DbManager = LiteDB.Create(SourceDb)
    DbManager.DebugPrintRecordset LiteMetaSQLEngine.Version, EngineInfo.Range("A1")
    DbManager.DebugPrintRecordset LiteMetaSQLEngine.CompileOptions, EngineInfo.Range("B1")
    DbManager.DebugPrintRecordset LiteMetaSQLEngine.Modules, EngineInfo.Range("C1")
    DbManager.DebugPrintRecordset LiteMetaSQLEngine.Pragmas, EngineInfo.Range("D1")
    DbManager.DebugPrintRecordset LiteMetaSQLEngine.Functions, EngineInfo.Range("E1")
End Sub


'@EntryPoint
'@Description "Collects SQLite database metadata"
Private Sub Database()
Attribute Database.VB_Description = "Collects SQLite database metadata"
    Dim SourceDb As String
    SourceDb = "SQLiteDB.db"
    Dim DbManager As LiteDB
    Set DbManager = LiteDB.Create(SourceDb)
    DbManager.DebugPrintRecordset DbManager.SQLInfo.Tables, Tables.Range("A1")
    DbManager.DebugPrintRecordset DbManager.SQLInfo.ForeingKeys, ForeignKeys.Range("A1")
    DbManager.DebugPrintRecordset DbManager.SQLInfo.Indices(True), Indices.Range("A1")
    DbManager.DebugPrintRecordset DbManager.SQLInfo.FKChildIndices, FKChildIndices.Range("A1")
    DbManager.DebugPrintRecordset DbManager.SQLInfo.SimilarIndices, SimilarIndices.Range("A1")
    DbManager.DebugPrintRecordset DbManager.SQLInfo.TableColumns("companies"), Columns.Range("A1")
    ADOlib.RecordsetToQT DbManager.GetTableColumnsEx("test_table"), ColumnsEx.Range("A1")
End Sub


'@EntryPoint
Private Sub CloneDb()
    Dim SourceDb As String
    SourceDb = "SQLiteDB.db"
    Dim TargetDb As String
    TargetDb = "Dest.db"

    '@Ignore FunctionReturnValueDiscarded
    LiteDB.CloneDb TargetDb, SourceDb
End Sub
