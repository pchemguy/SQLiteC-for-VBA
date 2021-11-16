Attribute VB_Name = "SQLiteIntropectionExample"
'@Folder "SQLite.MetaSQL.Examples"
'@IgnoreModule VariableNotUsed, IndexedDefaultMemberAccess, ProcedureNotUsed
Option Explicit
Option Private Module
Option Compare Text


'@Description "Collects SQLite engine information and ouputs via QueryTables and to 'immediate'"
Private Sub Engine()
Attribute Engine.VB_Description = "Collects SQLite engine information and ouputs via QueryTables and to 'immediate'"
    Dim dbm As LiteMan
    Set dbm = LiteMan(":memory:")
    With dbm
        .DebugPrintRecordset LiteMetaSQL.Version, EngineInfo.Range("A1")
        .DebugPrintRecordset LiteMetaSQL.Modules, EngineInfo.Range("C1")
        .DebugPrintRecordset LiteMetaSQL.Pragmas, EngineInfo.Range("D1")
        .DebugPrintRecordset LiteMetaSQL.Functions, EngineInfo.Range("E1")
        .DebugPrintRecordset LiteMetaSQL.CompileOptions, EngineInfo.Range("B1")
    End With
End Sub


'@Description "Collects SQLite database metadata"
Private Sub Database()
Attribute Database.VB_Description = "Collects SQLite database metadata"
    Dim dbm As LiteMan
    Set dbm = LiteMan(FixObjAdo.DefaultDbName)
    With dbm
        .DebugPrintRecordset .MetaSQL.Tables, Tables.Range("A1")
        .DebugPrintRecordset .MetaSQL.ForeingKeys, ForeignKeys.Range("A1")
        .DebugPrintRecordset .MetaSQL.Indices(True), Indices.Range("A1")
        .DebugPrintRecordset .MetaSQL.FKChildIndices, FKChildIndices.Range("A1")
        .DebugPrintRecordset .MetaSQL.SimilarIndices, SimilarIndices.Range("A1")
        .DebugPrintRecordset .MetaSQL.TableColumns("companies"), Columns.Range("A1")
        
        ADOlib.RecordsetToQT .MetaADO.GetTableColumnsEx("test_table"), ColumnsEx.Range("A1")
    End With
End Sub
