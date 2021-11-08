Attribute VB_Name = "SQLiteIntropectionExample"
'@Folder "SQLite.MetaSQL.Examples"
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
    Dim dbm As LiteMan
    Set dbm = LiteMan(SourceDb)
    With dbm
        .DebugPrintRecordset .MetaSQL.Version, EngineInfo.Range("A1")
        .DebugPrintRecordset .MetaSQL.Modules, EngineInfo.Range("C1")
        .DebugPrintRecordset .MetaSQL.Pragmas, EngineInfo.Range("D1")
        .DebugPrintRecordset .MetaSQL.Functions, EngineInfo.Range("E1")
        .DebugPrintRecordset .MetaSQL.CompileOptions, EngineInfo.Range("B1")
    End With
End Sub


'@EntryPoint
'@Description "Collects SQLite database metadata"
Private Sub Database()
Attribute Database.VB_Description = "Collects SQLite database metadata"
    Dim SourceDb As String
    SourceDb = FixObjAdo.DefaultDbName
    
    Dim dbm As LiteMan
    Set dbm = LiteMan(SourceDb)
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


'@EntryPoint
Private Sub CloneDb()
    Dim SourceDb As String
    SourceDb = FixObjAdo.DefaultDbName
    Dim DestinationDb As String
    DestinationDb = "Dest.db"
    Dim dbm As LiteMan
    Set dbm = LiteMan.CloneDb(DestinationDb, SourceDb)
    Debug.Print dbm.ExecADO.MainDB
End Sub
