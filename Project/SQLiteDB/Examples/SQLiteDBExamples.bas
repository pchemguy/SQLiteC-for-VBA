Attribute VB_Name = "SQLiteDBExamples"
'@Folder "SQLiteDB.Examples"
'@IgnoreModule ProcedureNotUsed, VariableNotUsed, IndexedDefaultMemberAccess
Option Explicit
Option Private Module
Option Compare Text

Private Const LIB_NAME As String = "SQLiteDBVBA"
Private Const PATH_SEP As String = "\"
Private Const REL_PREFIX As String = "Library" & PATH_SEP & LIB_NAME & PATH_SEP


Private Sub CloneDb()
    Dim SourceDb As String
    SourceDb = REL_PREFIX & "SQLiteDBVBA.db"
    Dim TargetDb As String
    TargetDb = REL_PREFIX & "Dest.db"

    '@Ignore FunctionReturnValueDiscarded
    SQLiteDB.CloneDb TargetDb, SourceDb
End Sub


Private Sub SetJournalMode()
    Dim FileName As String
    FileName = REL_PREFIX & "TestA.db"
    
    Dim DbManager As SQLiteDB
    Set DbManager = SQLiteDB(FileName)
    DbManager.AttachDatabase REL_PREFIX & "TestB.db"
    DbManager.AttachDatabase REL_PREFIX & "TestC.db"
    
    DbManager.JournalModeSet "WAL", "ALL"
End Sub


Private Sub PrintTable()
    Dim OutputWS As Excel.Worksheet
    Set OutputWS = Buffer
        
    Dim FileName As String
    FileName = REL_PREFIX & LIB_NAME & ".db"
    
    Dim DbManager As SQLiteDB
    Set DbManager = SQLiteDB(FileName)
    
    Dim SQLTool As SQLlib
    Set SQLTool = SQLlib("contacts")
    SQLTool.Limit = 1000
    DbManager.DebugPrintRecordset SQLTool.SelectAll, OutputWS.Range("A1")
End Sub
