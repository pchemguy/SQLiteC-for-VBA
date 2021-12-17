Attribute VB_Name = "LiteExamples"
'@Folder "SQLite.ADO.ADemo"
'@IgnoreModule ProcedureNotUsed: This is a demo/examples module
'@IgnoreModule IndexedDefaultMemberAccess
Option Explicit


'''' Should print "SQLiteODBC Found: True" if the SQLiteODBC driver is available
'@Description "Checks and prints SQLiteODBC driver status."
Private Sub SQLite3ODBCDriverCheck()
Attribute SQLite3ODBCDriverCheck.VB_Description = "Checks and prints SQLiteODBC driver status."
    Debug.Print "SQLiteODBC Found: " & LiteMan.SQLite3ODBCDriverCheck()
End Sub

'''' Should print SQLite version number, e.g., "SQLite version: 3.32.3"
'@Description "Queries and prints SQLite version number."
Private Sub CheckConnectionAndVersion()
Attribute CheckConnectionAndVersion.VB_Description = "Queries and prints SQLite version number."
    Debug.Print "SQLite version: " & LiteMan(":mem:").ExecADO.GetScalar(vbNullString)
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
    Set dbm = LiteMan(":tmp:", , "StepAPI=True;Timeout=10000;SyncPragma=NORMAL;FKSupport=True;")
    Debug.Print dbm.ExecADO.MainDB
    Dim AffectedRows As Long
    AffectedRows = dbm.ExecADO.ExecuteNonQuery(FixSQLFunc.CreateWithData)
    Debug.Assert AffectedRows > 0
    
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

Private Sub ConnectSQLiteAdoCommandSourceFreezeWithBusyDb()
    Dim Driver As String
    Driver = "SQLite3 ODBC Driver"
    Dim Database As String
    Database = Environ("Temp") & "\" & CStr(Format(Now, "yyyy-mm-dd_hh-mm-ss.")) _
        & CStr((Timer * 10000) Mod 10000) & CStr(Round(Rnd * 10000, 0)) & ".db"
    Debug.Print Database
    Dim Options As String
    Options = "JournalMode=DELETE;SyncPragma=NORMAL;FKSupport=True;"

    Dim AdoConnStr As String
    AdoConnStr = "Driver=" & Driver & ";" & "Database=" & Database & ";" & Options
    
    Dim SQLQuery As String
    Dim RecordsAffected As Long: RecordsAffected = 0 '''' RD workaround
    Dim AdoCommand As ADODB.Command
    Set AdoCommand = New ADODB.Command
    With AdoCommand
        .CommandType = adCmdText
        .ActiveConnection = AdoConnStr
        .ActiveConnection.CursorLocation = adUseClient
    End With
    
    '''' ===== Create Functions table ===== ''''
    SQLQuery = Join(Array( _
        "CREATE TABLE functions(", _
        "    name    TEXT COLLATE NOCASE NOT NULL,", _
        "    builtin INTEGER             NOT NULL,", _
        "    type    TEXT COLLATE NOCASE NOT NULL,", _
        "    enc     TEXT COLLATE NOCASE NOT NULL,", _
        "    narg    INTEGER             NOT NULL,", _
        "    flags   INTEGER             NOT NULL", _
        ")" _
    ), vbLf)
    With AdoCommand
        .CommandText = SQLQuery
        .Execute RecordsAffected, Options:=adExecuteNoRecords
    End With
    
    '''' ===== Insert rows into Functions table ===== ''''
    SQLQuery = Join(Array( _
        "INSERT INTO functions", _
        "SELECT * FROM pragma_function_list" _
    ), vbLf)
    With AdoCommand
        .CommandText = SQLQuery
        .Execute RecordsAffected, Options:=adExecuteNoRecords
    End With
    
    '@Ignore StopKeyword
    Stop '''' Lock Db. For example, open in GUI admin tool and start a transaction
    '''' ===== Try changing journal mode ===== ''''
    On Error Resume Next
    With AdoCommand
        .CommandText = "PRAGMA journal_mode = 'WAL'"
        .Execute RecordsAffected, Options:=adExecuteNoRecords
    End With
    If Err.Number <> 0 Then
        Debug.Print "Error: #" & CStr(Err.Number) & ". " & vbNewLine & _
                    "Error description: " & Err.Description
    End If
    '@Ignore StopKeyword
    Stop
    On Error GoTo 0
    
    AdoCommand.ActiveConnection.Close
End Sub
