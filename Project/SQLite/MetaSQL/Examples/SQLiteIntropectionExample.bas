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
    Dim dbmADO As LiteDB
    Set dbmADO = LiteDB(SourceDb)
    Dim dbu As LiteUtils
    Set dbu = dbmADO.Util
    With LiteMetaSQLEngine
        dbu.DebugPrintRecordset .Version, EngineInfo.Range("A1")
        dbu.DebugPrintRecordset .Modules, EngineInfo.Range("C1")
        dbu.DebugPrintRecordset .Pragmas, EngineInfo.Range("D1")
        dbu.DebugPrintRecordset .functions, EngineInfo.Range("E1")
        dbu.DebugPrintRecordset .CompileOptions, EngineInfo.Range("B1")
    End With
End Sub


'@EntryPoint
'@Description "Collects SQLite database metadata"
Private Sub Database()
Attribute Database.VB_Description = "Collects SQLite database metadata"
    Dim SourceDb As String
    SourceDb = "SQLiteCDBVBA.db"
    
    Dim dbmADO As LiteDB
    Set dbmADO = LiteDB(SourceDb)
    Dim dbu As LiteUtils
    Set dbu = dbmADO.Util
    With LiteMetaSQL.Create()
        dbu.DebugPrintRecordset .Tables, Tables.Range("A1")
        dbu.DebugPrintRecordset .ForeingKeys, ForeignKeys.Range("A1")
        dbu.DebugPrintRecordset .Indices(True), Indices.Range("A1")
        dbu.DebugPrintRecordset .FKChildIndices, FKChildIndices.Range("A1")
        dbu.DebugPrintRecordset .SimilarIndices, SimilarIndices.Range("A1")
        dbu.DebugPrintRecordset .TableColumns("companies"), Columns.Range("A1")
    End With
    ADOlib.RecordsetToQT dbmADO.Meta.GetTableColumnsEx("test_table"), ColumnsEx.Range("A1")
End Sub


'@EntryPoint
Private Sub CloneDb()
    Dim SourceDb As String
    SourceDb = "SQLiteCDBVBA.db"
    Dim TargetDb As String
    TargetDb = "Dest.db"

    Dim dbmADO As LiteDB
    Set dbmADO = LiteDB(SourceDb)
    Dim dbeTarget As ILiteADO
    Set dbeTarget = dbmADO.Util.CloneDb(TargetDb, SourceDb)
    Debug.Print dbeTarget.MainDB
End Sub
