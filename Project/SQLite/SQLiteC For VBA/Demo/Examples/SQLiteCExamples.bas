Attribute VB_Name = "SQLiteCExamples"
'@Folder "SQLite.SQLiteC For VBA.Demo.Examples"
'@IgnoreModule
Option Explicit

Private Const LITE_LIB As String = "SQLiteCDBVBA"
Private Const PATH_SEP As String = "\"
Private Const LITE_RPREFIX As String = "Library" & PATH_SEP & LITE_LIB & PATH_SEP

Private Type TSQLiteCExamples
    dbm As SQLiteC
    dbc As SQLiteCConnection
    dbs As SQLiteCStatement
    dbr As SQLiteCRecordsetADO
End Type
Private this As TSQLiteCExamples
Private fso As New Scripting.FileSystemObject


'@EntryPoint "Runs this demo"
'@Description "Main entry point"
Public Sub Main()
Attribute Main.VB_Description = "Main entry point"
    InitDBM
    InitDBC
    InitDBS
    OpenDb
    CheckFunctionality
    CreateFunctionsTableWithData
    GetTableMetaFunctions
    GetFabRecordset
        
    Dim Result As Variant
    CreateTestTable
    InsertTestRows
    GetTableMeta
    
    Result = GetPagedTestRowsSet

    Result = GetScalarDbVersion
    Debug.Print Result
    Result = GetScalarDbPath
    Debug.Print Result
    
    Result = GetPagedRowSetFunctions
    Dim ResultB As Variant
    ResultB = GetRowSet2DFunctions
    
    PrepareStatementGetRowSetFilteredPlain
    
    PrepareStatementGetRowSetFilteredParams
    BindParamArray
    BindParamDict
    FinalizeStatement
    
    GetFirstFunctionName

    Result = Empty
    Result = RunFunctionsQuery
    
    Result = Empty
    Result = RunFunctionsQueryWithParamArray
    
    Result = Empty
    Result = RunFunctionsQueryWithParamDict
    
    Debug.Assert Not (IsNull(Result) Or IsError(Result))
    PrepareStatementGetScalar
    FinalizeStatement
    CloseDb
    Cleanup
End Sub


Private Sub Cleanup()
    Set this.dbs = Nothing
    Set this.dbc = Nothing
    Set this.dbm = Nothing
End Sub


'@Description "Creates database manager (SQLiteC) instance and loads DLLs via the DllManager class."
Private Sub InitDBM()
Attribute InitDBM.VB_Description = "Creates database manager (SQLiteC) instance and loads DLLs via the DllManager class."
    Dim DllPath As String
    DllPath = LITE_RPREFIX & "dll\" & ARCH
    Dim DllNames As Variant
    #If WIN64 Then
        DllNames = "sqlite3.dll"
    #Else
        DllNames = Array("icudt68.dll", "icuuc68.dll", "icuin68.dll", _
                         "icuio68.dll", "icutu68.dll", "sqlite3.dll")
    #End If
    Dim dbm As SQLiteC
    '@Ignore IndexedDefaultMemberAccess
    Set dbm = SQLiteC(DllPath, DllNames)
    If dbm Is Nothing Then
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                  "Failed to create an SQLiteC instance."
    Else
        Debug.Print "Database manager instance (SQLiteC class) is ready"
    End If
    
    '''' Test SQLite3.dll
    If Replace(dbm.Version(False), ".", "0") & "0" = CStr(dbm.Version) Then
        Debug.Print "Database engine version functionality test passed."
    Else
        Debug.Print "Database engine version functionality test failed."
    End If
    Set this.dbm = dbm
End Sub


'@Description "Creates connection object (SQLiteCConnection)."
Private Sub InitDBC()
Attribute InitDBC.VB_Description = "Creates connection object (SQLiteCConnection)."
    Dim dbm As SQLiteC
    Set dbm = this.dbm
    
    Dim DbFilePath As String
    DbFilePath = Environ("TEMP")
    Dim DbFileName As String
    DbFileName = Replace(Replace(Replace(Now(), "/", "-"), " ", "_"), ":", "-")
    Dim DbFileExt As String
    DbFileExt = ".db"
    Dim DbPathName As String
    DbPathName = fso.BuildPath(DbFilePath, DbFileName) & DbFileExt

    Dim dbc As SQLiteCConnection
    Set dbc = dbm.CreateConnection(DbPathName)
    If dbc Is Nothing Then
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                  "Failed to create an SQLiteCConnection instance."
    Else
        Debug.Print "Database SQLiteCConnection instance is ready."
    End If
    Set this.dbc = dbc
End Sub


'@Description "Creates statement object (SQLiteCStatement)."
Private Sub InitDBS()
Attribute InitDBS.VB_Description = "Creates statement object (SQLiteCStatement)."
    Dim dbc As SQLiteCConnection
    Set dbc = this.dbc

    Dim DbStmtName As String
    DbStmtName = vbNullString
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(DbStmtName)
    If dbs Is Nothing Then
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                  "Failed to create an SQLiteCStatement instance."
    Else
        Debug.Print "Database SQLiteCStatement instance is ready."
    End If
    Set this.dbs = dbs
End Sub


'@Description "Opens database connection."
Private Sub OpenDb()
Attribute OpenDb.VB_Description = "Opens database connection."
    Dim dbc As SQLiteCConnection
    Set dbc = this.dbc
    Dim ResultCode As SQLiteResultCodes
    
    ResultCode = dbc.OpenDb
    If ResultCode <> SQLITE_OK Or dbc.DbHandle = 0 Then
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                  "Failed to open db connection."
    Else
        Debug.Print "Database connection is ready."
    End If
End Sub


'@Description "Opens database connection."
Private Sub CloseDb()
Attribute CloseDb.VB_Description = "Opens database connection."
    Dim dbc As SQLiteCConnection
    Set dbc = this.dbc
    Dim ResultCode As SQLiteResultCodes
    
    ResultCode = dbc.CloseDb
    If ResultCode <> SQLITE_OK Or dbc.DbHandle <> 0 Then
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                  "Failed to close db connection."
    Else
        Debug.Print "Database connection is closed."
    End If
End Sub


'@Description "Runs a few basic checks (not involving the statement APIs)."
Private Sub CheckFunctionality()
Attribute CheckFunctionality.VB_Description = "Runs a few basic checks (not involving the statement APIs)."
    Dim dbc As SQLiteCConnection
    Set dbc = this.dbc
    Dim DbPathName As String
    DbPathName = dbc.DbPathName
    Dim ResultCode As SQLiteResultCodes
    
    '''' Check that the write mode is available
    Dim DbAccessMode As SQLiteDbAccess
    DbAccessMode = dbc.AccessMode
    If DbAccessMode <> SQLITE_DB_FULL Then
        Err.Raise ErrNo.PermissionDeniedErr, "SQLiteCExamples", _
                  "Database is not writable."
    Else
        Debug.Print "Database access mode is READ/WRITE."
    End If
        
    '''' Set journal mode to DELETE, verify that WAL files do not exist
    ResultCode = dbc.JournalModeSet(SQLITE_PAGER_JOURNALMODE_DELETE)
    If ResultCode <> SQLITE_OK Then
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                  "Failed to set journal mode DELETE."
    Else
        Debug.Print "Journal mode 'DELETE' set query completed."
    End If
    If fso.FileExists(DbPathName & "-shm") Or _
       fso.FileExists(DbPathName & "-wal") Then
        Debug.Print "Journal mode set error: WAL journal files found!"
    Else
        Debug.Print "Journal mode set appears OK: WAL journal files not found."
    End If
        
    '''' Set journal mode to WAL
    ResultCode = dbc.JournalModeSet(SQLITE_PAGER_JOURNALMODE_WAL)
    If ResultCode <> SQLITE_OK Then
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                  "Failed to set journal mode 'WAL'."
    Else
        Debug.Print "Journal mode 'WAL' set query completed."
    End If
    
    '''' Start immediate transaction
    ResultCode = dbc.Begin(SQLITE_TXN_IMMEDIATE)
    If ResultCode <> SQLITE_OK Then
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                  "Failed to start immediate transaction."
    Else
        Debug.Print "Started immediate transaction."
    End If
    
    '''' Verify transaction state
    Dim TxnStateCode As SQLiteTxnState
    TxnStateCode = dbc.TxnState
    If Not TxnStateCode > SQLITE_TXN_NONE Then
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                  "Transaction state is not confirmed."
    Else
        Debug.Print "Transaction state is: " & Array("NONE", "READ", "WRITE" _
                                                    )(TxnStateCode)
    End If

    '''' Verify 'WAL' journal files exist
    If fso.FileExists(DbPathName & "-shm") And _
       fso.FileExists(DbPathName & "-wal") Then
        Debug.Print "WAL journal files found, as expected."
    Else
        Debug.Print "WAL journal file(s) not found unexpectedly."
    End If

    '''' Commit transaction
    ResultCode = dbc.Commit
    If ResultCode <> SQLITE_OK Then
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                  "Failed to commit transaction."
    Else
        Debug.Print "Transaction commited."
    End If

    '''' Verify transaction state
    TxnStateCode = dbc.TxnState
    If Not TxnStateCode = SQLITE_TXN_NONE Then
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                  "Transaction state is not confirmed."
    Else
        Debug.Print "Transaction state is OK."
    End If

    '''' Set journal mode to DELETE, verify that WAL files do not exist
    ResultCode = dbc.JournalModeSet(SQLITE_PAGER_JOURNALMODE_DELETE)
    If ResultCode <> SQLITE_OK Then
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                  "Failed to set journal mode DELETE."
    Else
        Debug.Print "Journal mode 'DELETE' set query completed."
    End If
    If fso.FileExists(DbPathName & "-shm") Or _
       fso.FileExists(DbPathName & "-wal") Then
        Debug.Print "Journal mode set error: WAL journal files found!"
    Else
        Debug.Print "Journal mode set appears OK: WAL journal files are gone."
    End If
End Sub


Private Sub CreateTestTable()
    Dim dbs As SQLiteCStatement
    Set dbs = this.dbs
    Dim ResultCode As SQLiteResultCodes
    
    Dim SQLQuery As String
    SQLQuery = SQLiteCExamplesSQL.CreateTestTable
    Dim AffectedRows As Long
    AffectedRows = -2
    ResultCode = dbs.ExecuteNonQuery(SQLQuery, , AffectedRows)
    If ResultCode <> SQLITE_OK And ResultCode <> SQLITE_DONE Then
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                  "Failed to create table."
    Else
        Debug.Print "Table create query is complete, AffectedRows = " & CStr(AffectedRows) & "."
    End If
End Sub


Private Sub InsertTestRows()
    Dim dbs As SQLiteCStatement
    Set dbs = this.dbs
    Dim ResultCode As SQLiteResultCodes
    
    Dim SQLQuery As String
    SQLQuery = SQLiteCExamplesSQL.InsertTestRows
    Dim AffectedRows As Long
    AffectedRows = -2
    ResultCode = dbs.ExecuteNonQuery(SQLQuery, , AffectedRows)
    If ResultCode <> SQLITE_OK And ResultCode <> SQLITE_DONE Then
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                  "Failed to insert rows."
    Else
        Debug.Print "Insert query is complete, AffectedRows = " & CStr(AffectedRows) & "."
    End If
    Debug.Assert AffectedRows = dbs.AffectedRowsCount
End Sub


Private Function GetPagedTestRowsSet() As Variant
    Dim dbs As SQLiteCStatement
    Set dbs = this.dbs
    
    Dim SQLQuery As String
    SQLQuery = SQLiteCExamplesSQL.SelectFromTestTable
    
    GetPagedTestRowsSet = dbs.GetPagedRowSet(SQLQuery)
End Function


Private Sub GetTableMetaFunctions()
    Dim dbs As SQLiteCStatement
    Set dbs = this.dbs
    Dim ResultCode As SQLiteResultCodes
    
    Dim SQLQuery As String
    SQLQuery = SQLiteCExamplesSQL.SelectFromFunctionsTable
    
    ResultCode = dbs.Prepare16V2(SQLQuery)
    If ResultCode <> SQLITE_OK Or dbs.StmtHandle = 0 Then
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                  "Failed to prepare statement."
    Else
        Debug.Print "Statement is prepared."
    End If
    
    ResultCode = dbs.DbExecutor.ExecuteStepAPI
    Select Case ResultCode
        Case SQLITE_ROW
            Debug.Print "Step API returned row."
        Case SQLITE_OK, SQLITE_DONE
            Debug.Print "Step API returned NoData."
            Exit Sub
        Case Else
            Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                      "Failed to execute the Step API."
    End Select
    
    ResultCode = dbs.DbExecutor.TableMetaCollect
    If ResultCode <> SQLITE_OK Then
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                  "Failed to get columns meta."
    Else
        Debug.Print "Retrieved columns meta,"
    End If

    ResultCode = dbs.Reset
    If ResultCode <> SQLITE_OK Then
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                  "Failed to reset statement."
    Else
        Debug.Print "The statement is reset."
    End If
End Sub


Private Sub GetTableMeta()
    Dim dbs As SQLiteCStatement
    Set dbs = this.dbs
    Dim ResultCode As SQLiteResultCodes
    
    Dim SQLQuery As String
    SQLQuery = SQLiteCExamplesSQL.SelectFromTestTable
    
    ResultCode = dbs.Prepare16V2(SQLQuery)
    If ResultCode <> SQLITE_OK Or dbs.StmtHandle = 0 Then
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                  "Failed to prepare statement."
    Else
        Debug.Print "Statement is prepared."
    End If
    
    ResultCode = dbs.DbExecutor.ExecuteStepAPI
    Select Case ResultCode
        Case SQLITE_ROW
            Debug.Print "Step API returned row."
        Case SQLITE_OK, SQLITE_DONE
            Debug.Print "Step API returned NoData."
            Exit Sub
        Case Else
            Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                      "Failed to execute the Step API."
    End Select
    
    ResultCode = dbs.DbExecutor.TableMetaCollect
    If ResultCode <> SQLITE_OK Then
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                  "Failed to get columns meta."
    Else
        Debug.Print "Retrieved columns meta,"
    End If

    ResultCode = dbs.Reset
    If ResultCode <> SQLITE_OK Then
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                  "Failed to reset statement."
    Else
        Debug.Print "The statement is reset."
    End If
End Sub


Private Function GetScalarDbVersion() As Variant
    Dim dbs As SQLiteCStatement
    Set dbs = this.dbs
    
    Dim SQLQuery As String
    SQLQuery = SQLiteCExamplesSQL.GetSQLiteVersion
    
    GetScalarDbVersion = dbs.GetScalar(SQLQuery)
End Function


Private Function GetScalarDbPath() As Variant
    Dim dbs As SQLiteCStatement
    Set dbs = this.dbs
    
    Dim SQLQuery As String
    SQLQuery = SQLiteCExamplesSQL.GetDbPath
    
    GetScalarDbPath = dbs.GetScalar(SQLQuery)
End Function


Private Sub FinalizeStatement()
    Dim dbs As SQLiteCStatement
    Set dbs = this.dbs
    Dim ResultCode As SQLiteResultCodes
        
    ResultCode = dbs.Finalize
    If ResultCode <> SQLITE_OK Or dbs.StmtHandle <> 0 Then
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                  "Failed to finalize statement."
    Else
        Debug.Print "Statement is finalized."
    End If
End Sub


Private Sub PrepareStatementGetScalar()
    Dim dbs As SQLiteCStatement
    Set dbs = this.dbs
    Dim ResultCode As SQLiteResultCodes
    
    Dim SQLQuery As String
    SQLQuery = SQLiteCExamplesSQL.GetSQLiteVersion
    
    ResultCode = dbs.Prepare16V2(SQLQuery)
    If ResultCode <> SQLITE_OK Or dbs.StmtHandle = 0 Then
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                  "Failed to prepare statement."
    Else
        Debug.Print "Statement is prepared."
    End If
End Sub


Private Sub PrepareStatementGetRowSetFilteredPlain()
    Dim dbs As SQLiteCStatement
    Set dbs = this.dbs
    Dim ResultCode As SQLiteResultCodes
    
    Dim SQLQuery As String
    SQLQuery = SQLiteCExamplesSQL.FunctionsPragmaTableFiltered
    
    ResultCode = dbs.Prepare16V2(SQLQuery)
    If ResultCode <> SQLITE_OK Or dbs.StmtHandle = 0 Then
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                  "Failed to prepare statement."
    Else
        Debug.Print "Statement is prepared."
    End If
End Sub


Private Sub PrepareStatementGetRowSetFilteredParams()
    Dim dbs As SQLiteCStatement
    Set dbs = this.dbs
    Dim ResultCode As SQLiteResultCodes
    
    Dim SQLQuery As String
    SQLQuery = SQLiteCExamplesSQL.FunctionsPragmaTableNamedParams
    
    ResultCode = dbs.Prepare16V2(SQLQuery)
    If ResultCode <> SQLITE_OK Or dbs.StmtHandle = 0 Then
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                  "Failed to prepare statement."
    Else
        Debug.Print "Statement is prepared."
    End If
End Sub


Private Sub BindParamArray()
    Dim dbs As SQLiteCStatement
    Set dbs = this.dbs
    Dim ResultCode As SQLiteResultCodes
    
    Dim ParamsArray As Variant
    ParamsArray = SQLiteCExamplesSQL.FunctionsFilteredNamedParamsArray
    
    dbs.DbParameters.BindClear
    ResultCode = dbs.DbParameters.BindDictOrArray(ParamsArray)
    If ResultCode <> SQLITE_OK Then
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                  "Failed to bind array parameters."
    Else
        Debug.Print "Array parameters are bound."
    End If
End Sub


Private Sub BindParamDict()
    Dim dbs As SQLiteCStatement
    Set dbs = this.dbs
    Dim ResultCode As SQLiteResultCodes
    
    Dim ParamsDict As Scripting.Dictionary
    Set ParamsDict = SQLiteCExamplesSQL.FunctionsFilteredNamedParamsDict
    
    dbs.DbParameters.BindClear
    ResultCode = dbs.DbParameters.BindDictOrArray(ParamsDict)
    If ResultCode <> SQLITE_OK Then
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                  "Failed to bind dict parameters."
    Else
        Debug.Print "Dict parameters are bound."
    End If
End Sub


Private Function RunFunctionsQuery() As Variant
    Dim dbs As SQLiteCStatement
    Set dbs = this.dbs
    Dim ResultCode As SQLiteResultCodes
    
    Dim SQLQuery As String
    SQLQuery = SQLiteCExamplesSQL.FunctionsTableFiltered
    Dim ParamsArray As Variant
    ParamsArray = Null
    
    RunFunctionsQuery = dbs.GetPagedRowSet(SQLQuery, ParamsArray)
End Function


Private Function RunFunctionsQueryWithParamArray() As Variant
    Dim dbs As SQLiteCStatement
    Set dbs = this.dbs
    Dim ResultCode As SQLiteResultCodes
    
    Dim SQLQuery As String
    SQLQuery = SQLiteCExamplesSQL.FunctionsTableNamedParams
    Dim ParamsArray As Variant
    ParamsArray = SQLiteCExamplesSQL.FunctionsFilteredNamedParamsArray
    
    RunFunctionsQueryWithParamArray = dbs.GetPagedRowSet(SQLQuery, ParamsArray)
End Function


Private Function RunFunctionsQueryWithParamDict() As Variant
    Dim dbs As SQLiteCStatement
    Set dbs = this.dbs
    Dim ResultCode As SQLiteResultCodes
    
    Dim SQLQuery As String
    SQLQuery = SQLiteCExamplesSQL.FunctionsTableNamedParams
    Dim ParamsDict As Variant
    Set ParamsDict = SQLiteCExamplesSQL.FunctionsFilteredNamedParamsDict
    
    RunFunctionsQueryWithParamDict = dbs.GetPagedRowSet(SQLQuery, ParamsDict)
End Function


Private Function GetFirstFunctionName() As Variant
    Dim dbs As SQLiteCStatement
    Set dbs = this.dbs
    
    Dim SQLQuery As String
    SQLQuery = SQLiteCExamplesSQL.SelectFromFunctionsTable
    
    GetFirstFunctionName = dbs.GetScalar(SQLQuery)
End Function


Private Function GetPagedRowSetFunctions() As Variant
    Dim dbs As SQLiteCStatement
    Set dbs = this.dbs
        
    Dim SQLQuery As String
    SQLQuery = SQLiteCExamplesSQL.SelectFromFunctionsTable
    
    GetPagedRowSetFunctions = dbs.GetPagedRowSet(SQLQuery)
End Function


Private Function GetRowSet2DFunctions() As Variant
    Dim dbs As SQLiteCStatement
    Set dbs = this.dbs
    
    Dim SQLQuery As String
    SQLQuery = SQLiteCExamplesSQL.SelectFromFunctionsTable
    
    GetRowSet2DFunctions = dbs.GetRowSet2D(SQLQuery)
End Function


Private Sub CreateFunctionsTableWithData()
    Dim dbc As SQLiteCConnection
    Set dbc = this.dbc
    Dim ResultCode As SQLiteResultCodes
    
    Dim SQLQuery As String
    SQLQuery = SQLiteCExamplesSQL.CreateFunctionsTableWithData
    Dim AffectedRows As Long
    AffectedRows = -2
    ResultCode = dbc.ExecuteNonQueryPlain(SQLQuery, AffectedRows)
    If ResultCode <> SQLITE_OK Then
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                  "Failed to create table."
    Else
        Debug.Print "Table create query is complete, AffectedRows = " & CStr(AffectedRows) & "."
    End If
End Sub


Private Sub GetFabRecordset()
    Dim SQLQuery As String
    SQLQuery = SQLiteCExamplesSQL.SelectFromFunctionsTable
    Set this.dbr = this.dbs.GetRecordset(SQLQuery)
End Sub
