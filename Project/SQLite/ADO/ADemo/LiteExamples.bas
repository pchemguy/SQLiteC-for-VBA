Attribute VB_Name = "LiteExamples"
'@Folder "SQLite.ADO.ADemo"
'@IgnoreModule ProcedureNotUsed: This is a demo/examples module
'@IgnoreModule IndexedDefaultMemberAccess
Option Explicit


'''' Should print "True" if the SQLiteODBC driver is available
'@Description "Checks and prints SQLiteODBC driver status."
Private Sub SQLite3ODBCDriverCheck()
Attribute SQLite3ODBCDriverCheck.VB_Description = "Checks and prints SQLiteODBC driver status."
    Debug.Print LiteMan.SQLite3ODBCDriverCheck()
End Sub

'''' Should print SQLite version number, e.g., "3.32.3"
'@Description "Queries and prints SQLite version number."
Private Sub CheckConnectionAndVersion()
Attribute CheckConnectionAndVersion.VB_Description = "Queries and prints SQLite version number."
    Debug.Print LiteMan(":mem:").ExecADO.GetScalar(vbNullString)
End Sub

'''' Should print pathname of the new database file, e.g.,
''''   %temp%\2021_11_18-15_24_13-6D49DA51.db
'@Description "Prints pathname of the created temp database."
Private Sub CreateTmpDb()
Attribute CreateTmpDb.VB_Description = "Prints pathname of the created temp database."
    Debug.Print Replace(LiteMan(":tmp:").ExecADO.MainDB, Environ("Temp"), "%temp%")
End Sub

'''' Should create "Dest.db" clone of the test database in ThisWorkbook.Path
'@Description "Clones default database and returns DB manager bound the new clone."
Private Sub CloneDb()
Attribute CloneDb.VB_Description = "Clones default database and returns DB manager bound the new clone."
    Dim SourceDb As String
    SourceDb = FixObjAdo.DefaultDbName
    Dim DestinationDb As String
    DestinationDb = "Dest.db"
    Dim dbm As LiteMan
    Set dbm = LiteMan.CloneDb(DestinationDb, SourceDb)
    Debug.Print dbm.ExecADO.MainDB
End Sub

'@Description "Sets and gets/prints database journal mode."
Private Sub SetGetJournalMode()
Attribute SetGetJournalMode.VB_Description = "Sets and gets/prints database journal mode."
    Dim dbm As LiteMan
    Set dbm = LiteMan(":tmp:")
    Debug.Print dbm.ExecADO.MainDB
    dbm.JournalModeSet "DELETE"
    Debug.Print dbm.JournalModeGet '''' Prints "delete"
    dbm.JournalModeSet "WAL"
    Debug.Print dbm.JournalModeGet '''' Prints "wal"
End Sub

Private Sub DemoHostFreezeWithBusyDb()
    Const PROC_NAME As String = "DemoHostFreezeWithBusyDb"
    Dim dbm As LiteMan
    Set dbm = LiteMan(":tmp:")
    Debug.Print dbm.ExecADO.MainDB
    
    Dim dbAdo As LiteADO
    Set dbAdo = dbm.ExecADO
    dbAdo.AdoConnection.CommandTimeout = 1
    dbAdo.AdoCommand.CommandTimeout = 1
    
    dbm.JournalModeSet "DELETE"
    '@Ignore StopKeyword
    Stop '''' Lock Db. For example, open in GUI admin tool and start a transaction
    Dim Start As Single
    Start = Timer
    On Error Resume Next
    dbm.JournalModeSet "WAL"
    Dim Delta As Long
    Delta = Round((Timer - Start), 2)
    On Error GoTo 0
    Debug.Print PROC_NAME & ":" & " App has been locked due to busy db for " & _
        Delta & " s."
End Sub
