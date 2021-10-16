Attribute VB_Name = "SQLiteCConnectionQueryTests"
'@Folder "SQLiteC For VBA.Connection"
'@TestModule
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed
Option Explicit
Option Private Module

#Const LateBind = LateBindTests
#If LateBind Then
    Private Assert As Object
#Else
    Private Assert As Rubberduck.PermissiveAssertClass
#End If
Private Fixtures As SQLiteCTestFixtures


'This method runs once per module.
'@ModuleInitialize
Private Sub ModuleInitialize()
    #If LateBind Then
        Set Assert = CreateObject("Rubberduck.PermissiveAssertClass")
    #Else
        Set Assert = New Rubberduck.PermissiveAssertClass
    #End If
    Set Fixtures = New SQLiteCTestFixtures
End Sub


'This method runs once per module.
'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
    Set Fixtures = Nothing
End Sub


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("Query")
Private Sub ztcExecuteNonQueryPlain_VerifiesTxnStateAndAffectedRecords()
    On Error GoTo TestFail

Arrange:
Act:
    Dim dbc As SQLiteCConnection
    Set dbc = Fixtures.zfxGetConnDbMemory
    Dim SQLQuery As String
    SQLQuery = Fixtures.zfxGetStmtCreateTableIRBNT
    Dim AffectedRecords As Long
    Dim ResultCode As SQLiteResultCodes
    Dim TxnStateCode As SQLiteTxnState
Assert:
        ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error"
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
        ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Query")
Private Sub ztcExecuteNonQueryPlain_VerifiesModifyQueryOnlyError()
    On Error GoTo TestFail

Arrange:
Act:
    Dim dbc As SQLiteCConnection
    Set dbc = Fixtures.zfxGetConnDbMemory
    Dim SQLQuery As String
    SQLQuery = Fixtures.zfxGetStmtCreateTableIRBNT
    Dim AffectedRecords As Long
    Dim ResultCode As SQLiteResultCodes
    Dim TxnStateCode As SQLiteTxnState
Assert:
        ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error"
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
        ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DbStatement")
Private Sub ztcCreateStatement_VerifiesNewStatement()
    On Error GoTo TestFail

Arrange:
Act:
    Dim dbc As SQLiteCConnection
    Set dbc = Fixtures.zfxGetConnDbMemory
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
