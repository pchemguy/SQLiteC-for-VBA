Attribute VB_Name = "SQLiteCConnectionQueryTests"
'@Folder "SQLite.SQLiteC.Connection"
'@TestModule
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed, UnhandledOnErrorResumeNext
Option Explicit
Option Private Module

#If LateBind Then
    Private Assert As Object
#Else
    Private Assert As Rubberduck.PermissiveAssertClass
#End If


'This method runs once per module.
'@ModuleInitialize
Private Sub ModuleInitialize()
    #If LateBind Then
        Set Assert = CreateObject("Rubberduck.PermissiveAssertClass")
    #Else
        Set Assert = New Rubberduck.PermissiveAssertClass
    #End If
End Sub


'This method runs once per module.
'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
End Sub


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("Query")
Private Sub ztcExecuteNonQueryPlain_VerifiesTxnStateAndAffectedRecords()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim AffectedRecords As Long
    Dim ResultCode As SQLiteResultCodes
    Dim TxnStateCode As SQLiteTxnState
Act:
    Dim SQLQuery As String
    SQLQuery = FixSQLITRB.CreateWithData
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error"
Assert:
    ResultCode = dbc.SavePoint("ABCDEFG")
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Txn SavePoint error"
    ResultCode = dbc.ExecuteNonQueryPlain(SQLQuery, AffectedRecords)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected query error"
    Assert.AreEqual 5, AffectedRecords, "AffectedRecords mismatch"
    TxnStateCode = SQLITE_TXN_NULL
    TxnStateCode = dbc.TxnState("main")
    Assert.IsTrue TxnStateCode = SQLITE_TXN_WRITE, "Unexpected Txn state"
    ResultCode = dbc.ReleasePoint("ABCDEFG")
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Txn ReleasePoint error"
Cleanup:
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Query")
Private Sub ztcExecuteNonQueryPlain_VerifiesCreateTable()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem

    Dim AffectedRecords As Long
    Dim ResultCode As SQLiteResultCodes
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error"
Act:
    Dim SQLQuery As String
    SQLQuery = FixSQLITRB.CreateWithData
    ResultCode = dbc.ExecuteNonQueryPlain(SQLQuery, AffectedRecords)
Assert:
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected ExecuteNonQueryPlain error"
    Assert.AreEqual 5, AffectedRecords, "AffectedRecords mismatch"
Cleanup:
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'''' N.B.: The only difference from the test below is the SQLQuery prefix
''''       FixSQL.DROPTableITRB & vbNewLine
'@TestMethod("Query")
Private Sub ztcExecuteNonQueryPlain_VerifiesModifyQueryOnlyError()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem

    Dim AffectedRecords As Long
    Dim ResultCode As SQLiteResultCodes
    Dim TxnStateCode As SQLiteTxnState
Act:
    Dim SQLQuery As String
    SQLQuery = FixSQLITRB.CreateWithData
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error"
Assert:
    ResultCode = dbc.SavePoint("ABCDEFG")
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Txn SavePoint error"
    ResultCode = dbc.ExecuteNonQueryPlain("PRAGMA query_only=1")
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected query error"
    ResultCode = dbc.ExecuteNonQueryPlain(SQLQuery, AffectedRecords)
    Assert.AreEqual SQLITE_READONLY, ResultCode, "Expected SQLITE_READONLY error"
    Assert.AreEqual -1, AffectedRecords, "AffectedRecords mismatch"
    TxnStateCode = SQLITE_TXN_NULL
    TxnStateCode = dbc.TxnState("main")
    Assert.IsTrue TxnStateCode = SQLITE_TXN_NONE, "Unexpected Txn state"
    ResultCode = dbc.ReleasePoint("ABCDEFG")
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Txn ReleasePoint error"
Cleanup:
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'''' N.B.: The only difference from the test below is the SQLQuery prefix
''''       FixSQL.DROPTableITRB & vbNewLine
'@TestMethod("Query")
Private Sub ztcExecuteNonQueryPlain_TransactionTriggeredByAttemptedTableDrop()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem

    Dim AffectedRecords As Long
    Dim ResultCode As SQLiteResultCodes
    Dim TxnStateCode As SQLiteTxnState
Act:
    Dim SQLQuery As String
    SQLQuery = FixSQLITRB.Drop & vbNewLine & FixSQLITRB.CreateWithData
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error"
Assert:
    ResultCode = dbc.SavePoint("ABCDEFG")
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Txn SavePoint error"
    ResultCode = dbc.ExecuteNonQueryPlain("PRAGMA query_only=1")
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected query error"
    ResultCode = dbc.ExecuteNonQueryPlain(SQLQuery, AffectedRecords)
    Assert.AreEqual SQLITE_READONLY, ResultCode, "Expected SQLITE_READONLY error"
    Assert.AreEqual -1, AffectedRecords, "AffectedRecords mismatch"
    TxnStateCode = SQLITE_TXN_NULL
    TxnStateCode = dbc.TxnState("main")
    Assert.IsTrue TxnStateCode = SQLITE_TXN_READ, "Unexpected Txn state"
    ResultCode = dbc.ReleasePoint("ABCDEFG")
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Txn ReleasePoint error"
Cleanup:
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Query")
Private Sub ztcChangesCount_ThrowsOnClosedConnection()
    On Error Resume Next
    
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim SQLQuery As String
    SQLQuery = FixSQLITRB.CreateWithData
    Dim AffectedRecords As Long
    AffectedRecords = -2
    
    Dim ResultCode As SQLiteResultCodes
    ResultCode = SQLITE_ERROR
    ResultCode = dbc.ExecuteNonQueryPlain(SQLQuery, AffectedRecords)
    Assert.AreEqual SQLITE_ERROR, ResultCode, "ResultCode changed unexpectedly."
    
    Guard.AssertExpectedError Assert, ConnectionNotOpenedErr
End Sub


'@TestMethod("DbStatement")
Private Sub ztcCreateStatement_VerifiesNewStatement()
    On Error GoTo TestFail

Arrange:
Act:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim DbStmt As SQLiteCStatement
    Set DbStmt = dbc.CreateStatement(vbNullString)
Assert:
    Assert.IsNotNothing DbStmt, "DbStmt is not set."
    Assert.AreSame DbStmt, dbc.StmtDb(vbNullString), "Statement object mismatch"
    Assert.AreSame DbStmt, dbc.StmtDb, "Statement object mismatch (default name)"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Query")
Private Sub ztcAtDetach_VerifiesAttachExistingNewMem()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCTemp
    Dim dbcTmp As SQLiteCConnection
    Set dbcTmp = FixObjC.GetDBCTemp
    
    Dim ResultCode As SQLiteResultCodes
    Assert.AreEqual SQLITE_OK, dbcTmp.OpenDb, "Unexpected OpenDb error"
    ResultCode = dbcTmp.ExecuteNonQueryPlain(FixSQLMisc.CreateBasicTable)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected ExecuteNonQueryPlain error"
    Assert.AreEqual SQLITE_OK, dbcTmp.CloseDb, "Unexpected CloseDb error"
    Dim NewDbPathName As String
    NewDbPathName = FixObjC.RandomTempFileName
    
    Assert.AreEqual SQLITE_OK, dbc.OpenDb, "Unexpected OpenDb error"

    Dim DbStmtName As String
    DbStmtName = vbNullString
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(DbStmtName)
    Assert.IsNotNothing dbs, "Failed to create an SQLiteCStatement instance."
    
    '@Ignore SelfAssignedDeclaration
    Dim fso As New Scripting.FileSystemObject
    Dim SQLDbCount As String
    SQLDbCount = "SELECT count(*) As counter FROM pragma_database_list"
Act:
Assert:
    ResultCode = dbc.Attach(dbcTmp.DbPathName)
    Assert.AreEqual SQLITE_OK, ResultCode, "Failed to attach existing db"
    Assert.AreEqual 2, dbs.GetScalar(SQLDbCount), "Unexpected DbCount (exist)."
    ResultCode = dbc.Attach(":memory:")
    Assert.AreEqual SQLITE_OK, ResultCode, "Failed to attach memory db"
    Assert.AreEqual 3, dbs.GetScalar(SQLDbCount), "Unexpected DbCount (memory)."
    ResultCode = dbc.Attach(NewDbPathName)
    Assert.AreEqual SQLITE_OK, ResultCode, "Failed to attach new db"
    Assert.AreEqual 4, dbs.GetScalar(SQLDbCount), "Unexpected DbCount (new)."
    
    ResultCode = dbc.Detach(fso.GetBaseName(NewDbPathName))
    Assert.AreEqual SQLITE_OK, ResultCode, "Failed to detach new db"
    Assert.AreEqual 3, dbs.GetScalar(SQLDbCount), "Unexpected DbCount (new)."
    ResultCode = dbc.Detach("memory")
    Assert.AreEqual SQLITE_OK, ResultCode, "Failed to detach memory db"
    Assert.AreEqual 2, dbs.GetScalar(SQLDbCount), "Unexpected DbCount (memory)."
    ResultCode = dbc.Detach(fso.GetBaseName(dbcTmp.DbPathName))
    Assert.AreEqual SQLITE_OK, ResultCode, "Failed to detach existing db"
    Assert.AreEqual 1, dbs.GetScalar(SQLDbCount), "Unexpected DbCount (exist)."
Cleanup:
    Assert.AreEqual SQLITE_OK, dbc.CloseDb, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Query")
Private Sub ztcVacuum_VerifiesVacuumMainInPlace()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCTempFuncWithData
        
    Assert.AreEqual SQLITE_OK, dbc.OpenDb, "Unexpected OpenDb error"
Act:
Assert:
    Assert.AreEqual SQLITE_OK, dbc.Vacuum(), "Vacuum in-place error"
Cleanup:
    Assert.AreEqual SQLITE_OK, dbc.CloseDb, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Query")
Private Sub ztcVacuum_VerifiesVacuumMainToNew()
    On Error GoTo TestFail

Arrange:
    Dim dbcSrc As SQLiteCConnection
    Set dbcSrc = FixObjC.GetDBCTempFuncWithData
    Dim dbcDst As SQLiteCConnection
    Set dbcDst = FixObjC.GetDBCTemp

    Dim DbStmtNameSrc As String
    DbStmtNameSrc = Left(GenerateGUID, 8)
    Dim dbsSrc As SQLiteCStatement
    Set dbsSrc = dbcSrc.CreateStatement(DbStmtNameSrc)
    Dim DbStmtNameDst As String
    DbStmtNameDst = Left(GenerateGUID, 8)
    Dim dbsDst As SQLiteCStatement
    Set dbsDst = dbcDst.CreateStatement(DbStmtNameDst)

    Assert.AreEqual SQLITE_OK, dbcSrc.OpenDb, "Unexpected OpenDb error"
    
    Dim SQLQuery As String
    SQLQuery = "SELECT count(*) As counter FROM functions"
Act:
    Assert.AreEqual SQLITE_OK, dbcSrc.Vacuum("main", dbcDst.DbPathName), _
                    "Vacuum copy database error"
    Assert.AreEqual SQLITE_OK, dbcDst.OpenDb, "Unexpected OpenDb error"
Assert:
    Dim Expected As Long
    Expected = dbsSrc.GetScalar(SQLQuery)
    Dim Actual As Long
    Actual = dbsDst.GetScalar(SQLQuery)
    Assert.AreEqual Expected, Actual, "Row count mismatch."
Cleanup:
    Assert.AreEqual SQLITE_OK, dbcSrc.CloseDb, "Unexpected CloseDb error"
    Assert.AreEqual SQLITE_OK, dbcDst.CloseDb, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
