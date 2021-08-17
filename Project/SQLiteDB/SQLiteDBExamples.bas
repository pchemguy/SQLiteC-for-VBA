Attribute VB_Name = "SQLiteDBExamples"
'@Folder "SQLiteDB"
'@IgnoreModule ProcedureNotUsed, VariableNotUsed, IndexedDefaultMemberAccess
Option Explicit
Option Private Module
Option Compare Text


Private Sub CloneDb()
    Dim SourceDb As String
    SourceDb = "SQLiteDB.db"
    Dim TargetDb As String
    TargetDb = "Dest.db"

    '@Ignore FunctionReturnValueDiscarded
    SQLiteDB.CloneDb TargetDb, SourceDb
End Sub


Private Sub TestJournalMode()
    Dim DbManager As SQLiteDB
    Set DbManager = SQLiteDB("TestA.db")
    DbManager.AttachDatabase "TestB.db"
    DbManager.AttachDatabase "TestC.db"
    DbManager.JournalModeSet "WAL", "ALL"
End Sub


Private Sub PrintTable()
    Dim OutputWS As Excel.Worksheet
    Set OutputWS = Buffer
        
    Dim DbManager As SQLiteDB
    Set DbManager = SQLiteDB("SQLiteDb.db")
    
    Dim SQLTool As SQLlib
    Set SQLTool = SQLlib("contacts")
    SQLTool.Limit = 1000
    DbManager.DebugPrintRecordset SQLTool.SelectAll, OutputWS.Range("A1")
End Sub

