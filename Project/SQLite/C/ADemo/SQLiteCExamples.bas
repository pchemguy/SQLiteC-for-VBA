Attribute VB_Name = "SQLiteCExamples"
'@Folder "SQLite.C.ADemo"
'@IgnoreModule
Option Explicit

Private Const LITE_LIB As String = "SQLiteCAdo"
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
    
    FunctionsTableCREATEWithData
    FunctionsTableSELECT
    GetFabRecordset
        
    Dim Result As Variant
    ITRBTableCREATE
    ITRBTableINSERT
    GetTableMeta
    
    Result = GetPagedRowsSetITRB

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
    Result = GetPagedRowSetFunctionsFiltered
    
    Result = Empty
    Result = GetPagedRowSetFunctionsFilteredNamedParamsArray
    
    Result = Empty
    Result = GetPagedRowSetFunctionsFilteredNamedParamsDict
    
    PrepareStatementGetScalar
    FinalizeStatement
    CloseDb
    CleanUp
End Sub


Private Sub CleanUp()
    Set this.dbs = Nothing
    Set this.dbc = Nothing
    Set this.dbm = Nothing
    Set this.dbr = Nothing
    FixObjC.CleanUp
End Sub


'@Description "Creates database manager (SQLiteC) instance and loads DLLs via the DllManager class."
Private Sub InitDBM()
Attribute InitDBM.VB_Description = "Creates database manager (SQLiteC) instance and loads DLLs via the DllManager class."
    Dim DllPath As String
    DllPath = LITE_RPREFIX & "dll\" & ARCH
    Dim DllNames As Variant
    #If Win64 Then
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
    
    Dim DbPathName As String
    DbPathName = Environ("TEMP") & PATH_SEP & _
        Format(Now, "yyyy_mm_dd-hh_mm_ss-") & Left(GenerateGUID, 8) & ".db"
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
        Debug.Print "Transaction state is: " & _
                    Array("NONE", "READ", "WRITE")(TxnStateCode)
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


Private Sub ITRBTableCREATE()
    Dim dbs As SQLiteCStatement
    Set dbs = this.dbs
    Dim ResultCode As SQLiteResultCodes
    
    Dim SQLQuery As String
    SQLQuery = SQLiteCExamplesSQL.ITRBTableCREATE
    Dim AffectedRows As Long
    AffectedRows = -2
    ResultCode = dbs.ExecuteNonQuery(SQLQuery, , AffectedRows)
    If ResultCode <> SQLITE_OK And ResultCode <> SQLITE_DONE Then
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                  "Failed to create table."
    Else
        Debug.Print "Table create query is complete, AffectedRows = " & _
                    CStr(AffectedRows) & "."
    End If
End Sub


Private Sub ITRBTableINSERT()
    Dim dbs As SQLiteCStatement
    Set dbs = this.dbs
    Dim ResultCode As SQLiteResultCodes
    
    Dim SQLQuery As String
    SQLQuery = SQLiteCExamplesSQL.ITRBTableINSERT
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


Private Function GetPagedRowsSetITRB() As Variant
    Dim dbs As SQLiteCStatement
    Set dbs = this.dbs
    
    Dim SQLQuery As String
    SQLQuery = SQLiteCExamplesSQL.ITRBTableSELECT
    
    GetPagedRowsSetITRB = dbs.GetPagedRowSet(SQLQuery)
End Function


Private Sub FunctionsTableSELECT()
    Dim dbs As SQLiteCStatement
    Set dbs = this.dbs
    Dim dbsm As SQLiteCMeta
    Set dbsm = SQLiteCMeta(dbs)
    Dim ResultCode As SQLiteResultCodes
    
    Dim SQLQuery As String
    SQLQuery = SQLiteCExamplesSQL.FunctionsTableSELECT
    
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
    
    ResultCode = dbsm.TableMetaCollect
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
    Dim dbsm As SQLiteCMeta
    Set dbsm = SQLiteCMeta(dbs)
    Dim ResultCode As SQLiteResultCodes
    
    Dim SQLQuery As String
    SQLQuery = SQLiteCExamplesSQL.ITRBTableSELECT
    
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
    
    ResultCode = dbsm.TableMetaCollect
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


Private Function GetPagedRowSetFunctionsFiltered() As Variant
    Dim dbs As SQLiteCStatement
    Set dbs = this.dbs
    Dim ResultCode As SQLiteResultCodes
    
    Dim SQLQuery As String
    SQLQuery = SQLiteCExamplesSQL.FunctionsTableFiltered
    Dim ParamsArray As Variant
    ParamsArray = Null
    
    GetPagedRowSetFunctionsFiltered = dbs.GetPagedRowSet(SQLQuery, ParamsArray)
End Function


Private Function GetPagedRowSetFunctionsFilteredNamedParamsArray() As Variant
    Dim dbs As SQLiteCStatement
    Set dbs = this.dbs
    Dim ResultCode As SQLiteResultCodes
    
    Dim SQLQuery As String
    SQLQuery = SQLiteCExamplesSQL.FunctionsTableNamedParams
    Dim ParamsArray As Variant
    ParamsArray = SQLiteCExamplesSQL.FunctionsFilteredNamedParamsArray
    
    GetPagedRowSetFunctionsFilteredNamedParamsArray = dbs.GetPagedRowSet(SQLQuery, ParamsArray)
End Function


Private Function GetPagedRowSetFunctionsFilteredNamedParamsDict() As Variant
    Dim dbs As SQLiteCStatement
    Set dbs = this.dbs
    Dim ResultCode As SQLiteResultCodes
    
    Dim SQLQuery As String
    SQLQuery = SQLiteCExamplesSQL.FunctionsTableNamedParams
    Dim ParamsDict As Variant
    Set ParamsDict = SQLiteCExamplesSQL.FunctionsFilteredNamedParamsDict
    
    GetPagedRowSetFunctionsFilteredNamedParamsDict = dbs.GetPagedRowSet(SQLQuery, ParamsDict)
End Function


Private Function GetFirstFunctionName() As Variant
    Dim dbs As SQLiteCStatement
    Set dbs = this.dbs
    
    Dim SQLQuery As String
    SQLQuery = SQLiteCExamplesSQL.FunctionsTableSELECT
    
    GetFirstFunctionName = dbs.GetScalar(SQLQuery)
End Function


Private Function GetPagedRowSetFunctions() As Variant
    Dim dbs As SQLiteCStatement
    Set dbs = this.dbs
        
    Dim SQLQuery As String
    SQLQuery = SQLiteCExamplesSQL.FunctionsTableSELECT
    
    GetPagedRowSetFunctions = dbs.GetPagedRowSet(SQLQuery)
End Function


Private Function GetRowSet2DFunctions() As Variant
    Dim dbs As SQLiteCStatement
    Set dbs = this.dbs
    
    Dim SQLQuery As String
    SQLQuery = SQLiteCExamplesSQL.FunctionsTableSELECT
    
    GetRowSet2DFunctions = dbs.GetRowSet2D(SQLQuery)
End Function


Private Sub FunctionsTableCREATEWithData()
    Dim dbc As SQLiteCConnection
    Set dbc = this.dbc
    Dim ResultCode As SQLiteResultCodes
    
    Dim SQLQuery As String
    SQLQuery = SQLiteCExamplesSQL.FunctionsTableCREATEWithData
    Dim AffectedRows As Long
    AffectedRows = -2
    ResultCode = dbc.ExecuteNonQueryPlain(SQLQuery, AffectedRows)
    If ResultCode <> SQLITE_OK Then
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                  "Failed to create table."
    Else
        Debug.Print "Table create query is complete, AffectedRows = " & _
                    CStr(AffectedRows) & "."
    End If
End Sub


Private Sub GetFabRecordset()
    Dim SQLQuery As String
    SQLQuery = SQLiteCExamplesSQL.FunctionsTableSELECT
    Set this.dbr = this.dbs.GetRecordset(SQLQuery)
End Sub


Private Sub Txn()
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCTmpFuncWithData
    
    Dim ResultCode As SQLiteResultCodes
    ResultCode = dbc.OpenDb
    If ResultCode <> SQLITE_OK Then
        Debug.Print "Unexpected OpenDb error"
        Exit Sub
    End If
    
    ResultCode = dbc.Begin(SQLITE_TXN_IMMEDIATE)
    If ResultCode <> SQLITE_OK Then
        Debug.Print "Unexpected Txn Begin error"
        Exit Sub
    End If
    Dim TxnStateCode As SQLiteTxnState
    TxnStateCode = SQLITE_TXN_NULL
    TxnStateCode = dbc.TxnState("main")
    If TxnStateCode = SQLITE_TXN_NONE Then
        Debug.Print "Failed to begin transaction"
        Exit Sub
    Else
        Debug.Print "Transaction has succeessfully begun"
    End If
    ResultCode = dbc.Rollback
    If ResultCode <> SQLITE_OK Then Debug.Print "Unexpected Txn Commit error"
    TxnStateCode = dbc.TxnState("main")
    If TxnStateCode = SQLITE_TXN_NONE Then
        Debug.Print "Transaction rolled back."
    Else
        Debug.Print "Failed to roll back transaction"
    End If
    ResultCode = dbc.CloseDb
    If ResultCode <> SQLITE_OK Then Debug.Print "Unexpected CloseDb error"
End Sub


Private Sub TxnSave()
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCTmp
    
    Dim ResultCode As SQLiteResultCodes
    Dim SavePointName As String
    Dim TxnStateCode As SQLiteTxnState
    Dim AffectedRows As Long
    
    ResultCode = dbc.OpenDb
    If ResultCode <> SQLITE_OK Then
        Debug.Print "Unexpected OpenDb error"
        Exit Sub
    End If
    
    SavePointName = Left$(GenerateGUID, 8)
    ResultCode = dbc.SavePoint(SavePointName)
    If ResultCode = SQLITE_OK Then
        TxnStateCode = SQLITE_TXN_NULL
        TxnStateCode = dbc.TxnState("main")
        If TxnStateCode <> SQLITE_TXN_NONE Then
            Debug.Print "Unexpected Txn state"
            Exit Sub
        End If
    Else
        Debug.Print "Unexpected SavePoint error"
        Exit Sub
    End If
    
    ResultCode = dbc.ExecuteNonQueryPlain("SELECT * FROM pragma_function_list()", AffectedRows)
    If ResultCode = SQLITE_OK Then
        TxnStateCode = SQLITE_TXN_NULL
        TxnStateCode = dbc.TxnState("main")
        If TxnStateCode <> SQLITE_TXN_READ Then
            Debug.Print "Unexpected Txn state"
            Exit Sub
        End If
    Else
        Debug.Print "Unexpected ExecuteNonQueryPlain error"
        Exit Sub
    End If
    
    AffectedRows = FixObjC.CreateFunctionsTableWithData(dbc)
    If AffectedRows > 10 Then
        TxnStateCode = SQLITE_TXN_NULL
        TxnStateCode = dbc.TxnState("main")
        If TxnStateCode <> SQLITE_TXN_WRITE Then
            Debug.Print "Unexpected Txn state"
            Exit Sub
        End If
    Else
        Debug.Print "Unexpected ExecuteNonQueryPlain result"
        Exit Sub
    End If
    
    ResultCode = dbc.ReleasePoint(SavePointName)
    If ResultCode = SQLITE_OK Then
        TxnStateCode = SQLITE_TXN_NULL
        TxnStateCode = dbc.TxnState("main")
        If TxnStateCode <> SQLITE_TXN_NONE Then
            Debug.Print "Unexpected Txn state"
            Exit Sub
        End If
    Else
        Debug.Print "Unexpected ReleasePoint error"
        Exit Sub
    End If
    
    ResultCode = dbc.CloseDb
    If ResultCode <> SQLITE_OK Then Debug.Print "Unexpected CloseDb error"
End Sub


Private Sub TxnBusy()
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCTmpFuncWithData
    Dim dbcA As SQLiteCConnection
    Set dbcA = SQLiteCConnection(dbc.DbPathName)

    
    Dim ResultCode As SQLiteResultCodes
    If dbc.OpenDb <> SQLITE_OK Then
        Debug.Print "Unexpected OpenDb error"
        Exit Sub
    End If
    If dbcA.OpenDb <> SQLITE_OK Then
        Debug.Print "Unexpected OpenDb error"
        Exit Sub
    End If
    
    ResultCode = dbc.Begin(SQLITE_TXN_IMMEDIATE)
    If ResultCode <> SQLITE_OK Then
        Debug.Print "Unexpected Txn Begin error"
        Exit Sub
    End If
        
    Dim TxnStateCode As SQLiteTxnState
    TxnStateCode = SQLITE_TXN_NULL
    TxnStateCode = dbc.TxnState("main")
    If TxnStateCode = SQLITE_TXN_NONE Then
        Debug.Print "Failed to begin transaction"
        Exit Sub
    Else
        Debug.Print "Transaction has succeessfully begun"
    End If
    
    ResultCode = dbcA.Begin(SQLITE_TXN_IMMEDIATE)
    If ResultCode <> SQLITE_OK Then
        Debug.Print "Unexpected Txn Begin error"
    End If
    
    TxnStateCode = SQLITE_TXN_NULL
    TxnStateCode = dbcA.TxnState("main")
    If TxnStateCode <> SQLITE_TXN_NONE Then
        Debug.Print "Transaction has begun unexpectedly"
    End If
    
    If dbc.Rollback <> SQLITE_OK Then Debug.Print "Unexpected Txn Commit error"
    If dbc.TxnState("main") = SQLITE_TXN_NONE Then
        Debug.Print "Transaction rolled back."
    Else
        Debug.Print "Failed to roll back transaction"
    End If

    If dbc.CloseDb <> SQLITE_OK Then Debug.Print "Unexpected CloseDb error"
    If dbcA.CloseDb <> SQLITE_OK Then Debug.Print "Unexpected CloseDb error"
End Sub


Private Sub DbIsLocked()
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCTmpFuncWithData
    Dim dbcA As SQLiteCConnection
    Set dbcA = SQLiteCConnection(dbc.DbPathName)
    Dim dbcB As SQLiteCConnection
    Set dbcB = SQLiteCConnection(dbc.DbPathName)

    Dim TxnStateCode As SQLiteTxnState
    Dim ResultCode As SQLiteResultCodes
    Dim DbStatus As Variant
    
    If dbc.OpenDb <> SQLITE_OK Then
        Debug.Print "Unexpected OpenDb error"
        Exit Sub
    End If
    If dbcA.OpenDb <> SQLITE_OK Then
        Debug.Print "Unexpected OpenDb error"
        Exit Sub
    End If
    If dbcB.OpenDb(SQLITE_OPEN_READONLY) <> SQLITE_OK Then
        Debug.Print "Unexpected OpenDb error"
        Exit Sub
    End If
    Dim DbAccess As SQLiteDbAccess
    DbAccess = dbcB.AccessMode
    
    DbStatus = dbc.DbIsLocked
    
    ResultCode = dbcB.Begin(SQLITE_TXN_IMMEDIATE)
    TxnStateCode = SQLITE_TXN_NULL
    TxnStateCode = dbcB.TxnState("main")
    If TxnStateCode = SQLITE_TXN_NONE Then
        Debug.Print "Failed to begin transaction"
        Exit Sub
    Else
        Debug.Print "Transaction has succeessfully begun"
    End If
    
    
    ResultCode = dbc.Begin(SQLITE_TXN_IMMEDIATE)
    If ResultCode <> SQLITE_OK Then
        Debug.Print "Unexpected Txn Begin error"
        Exit Sub
    End If
        
    DbStatus = dbc.DbIsLocked
    DbStatus = dbcA.DbIsLocked
    
    TxnStateCode = SQLITE_TXN_NULL
    TxnStateCode = dbc.TxnState("main")
    If TxnStateCode = SQLITE_TXN_NONE Then
        Debug.Print "Failed to begin transaction"
        Exit Sub
    Else
        Debug.Print "Transaction has succeessfully begun"
    End If
    
    ResultCode = dbcA.Begin(SQLITE_TXN_IMMEDIATE)
    If ResultCode <> SQLITE_OK Then
        Debug.Print "Unexpected Txn Begin error"
    End If
    
    TxnStateCode = SQLITE_TXN_NULL
    TxnStateCode = dbcA.TxnState("main")
    If TxnStateCode <> SQLITE_TXN_NONE Then
        Debug.Print "Transaction has begun unexpectedly"
    End If
    
    If dbc.Rollback <> SQLITE_OK Then Debug.Print "Unexpected Txn Commit error"
    If dbc.TxnState("main") = SQLITE_TXN_NONE Then
        Debug.Print "Transaction rolled back."
    Else
        Debug.Print "Failed to roll back transaction"
    End If

    If dbc.CloseDb <> SQLITE_OK Then Debug.Print "Unexpected CloseDb error"
    If dbcA.CloseDb <> SQLITE_OK Then Debug.Print "Unexpected CloseDb error"
End Sub


Private Sub DupInMemoryToTempOnline()
    Dim dbcSrc As SQLiteCConnection
    Set dbcSrc = FixObjC.GetDBCMemITRBWithData
    Dim dbcDst As SQLiteCConnection
    Set dbcDst = FixObjC.GetDBCTmp
    
    Dim DbStmtName As String
    DbStmtName = vbNullString
    Dim dbsSrc As SQLiteCStatement
    Set dbsSrc = dbcSrc.CreateStatement(DbStmtName)
    If dbsSrc Is Nothing Then
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                  "Failed to create an SQLiteCStatement instance."
    Else
        Debug.Print "Database SQLiteCStatement instance is ready."
    End If

    If dbcSrc.OpenDb <> SQLITE_OK Then
        Debug.Print "Unexpected OpenDb error"
        Exit Sub
    End If
    If dbcDst.OpenDb <> SQLITE_OK Then
        Debug.Print "Unexpected OpenDb error"
        Exit Sub
    End If
    
    Dim Result As Variant
    Result = dbsSrc.GetScalar("SELECT count(*) As counter FROM itrb")
    If Result <> 5 Then
        Debug.Print "Unexpected RowCount."
        Exit Sub
    End If
        
    Dim ResultCode As SQLiteResultCodes
    ResultCode = SQLiteC.DupDbOnlineFull(dbcDst, "main", dbcSrc, "main")
End Sub


Private Sub DupInMemoryToTempVacuum()
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMemITRBWithData
    
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

    If dbc.OpenDb <> SQLITE_OK Then
        Debug.Print "Unexpected OpenDb error"
        Exit Sub
    End If
    
    Dim Result As Variant
    Result = dbs.GetScalar("SELECT count(*) As counter FROM itrb")
    If Result <> 5 Then
        Debug.Print "Unexpected RowCount."
        Exit Sub
    End If
        
    Dim ResultCode As SQLiteResultCodes
End Sub


Private Sub AttachDetach()
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCTmp
    Dim dbcTemp As SQLiteCConnection
    Set dbcTemp = FixObjC.GetDBCTmp
    
    If dbcTemp.OpenDb <> SQLITE_OK Then
        Debug.Print "Unexpected OpenDb error"
        Exit Sub
    End If
    If dbcTemp.ExecuteNonQueryPlain(FixSQLBase.CreateBasicTable) <> SQLITE_OK Then
        Debug.Print "Unexpected ExecuteNonQueryPlain error"
        Exit Sub
    End If
    If dbcTemp.CloseDb <> SQLITE_OK Then
        Debug.Print "Unexpected CloseDb error"
        Exit Sub
    End If
    
    If dbc.OpenDb <> SQLITE_OK Then
        Debug.Print "Unexpected OpenDb error"
        Exit Sub
    End If
    
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
    
    Dim fso As New Scripting.FileSystemObject
    Dim SQLDbCount As String
    SQLDbCount = "SELECT count(*) As counter FROM pragma_database_list"
    
    If dbc.Attach(dbcTemp.DbPathName) <> SQLITE_OK Then
        Debug.Print "Unexpected AttachDatabase error"
        Exit Sub
    End If
    If dbs.GetScalar(SQLDbCount) <> 2 Then
        Debug.Print "Unexpected DbCount."
        Exit Sub
    End If
    
    If dbc.Attach(":memory:") <> SQLITE_OK Then
        Debug.Print "Unexpected AttachDatabase error"
        Exit Sub
    End If
    If dbs.GetScalar(SQLDbCount) <> 3 Then
        Debug.Print "Unexpected DbCount."
        Exit Sub
    End If
    
    If dbc.Detach("memory") <> SQLITE_OK Then
        Debug.Print "Unexpected AttachDatabase error"
        Exit Sub
    End If
    If dbs.GetScalar(SQLDbCount) <> 2 Then
        Debug.Print "Unexpected DbCount."
        Exit Sub
    End If
    
    If dbc.Detach(fso.GetBaseName(dbcTemp.DbPathName)) <> SQLITE_OK Then
        Debug.Print "Unexpected AttachDatabase error"
        Exit Sub
    End If
    If dbs.GetScalar(SQLDbCount) <> 1 Then
        Debug.Print "Unexpected DbCount."
        Exit Sub
    End If
End Sub


Private Function GetBlankDb()
    Dim DbPathName As String
    DbPathName = FixObjC.RandomTempFileName("-----" & ".db")
    Dim dbc As SQLiteCConnection
    Set dbc = SQLiteCConnection(DbPathName)
    With dbc
        .OpenDb
        .Vacuum
        .CloseDb
    End With
End Function
