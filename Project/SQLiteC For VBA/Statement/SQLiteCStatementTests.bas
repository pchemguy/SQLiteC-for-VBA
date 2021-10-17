Attribute VB_Name = "SQLiteCStatementTests"
'@Folder "SQLiteC For VBA.Statement"
'@TestModule
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed, UnhandledOnErrorResumeNext
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
    Assert.IsNotNothing DbStmt.DbConnection, "Connection object not set."
    Assert.IsNotNothing DbStmt.DbExecutor, "Executor object not set."
    Assert.IsNothing DbStmt.DbParameters, "Parameters object should not be set."
    Assert.AreEqual 0, DbStmt.StmtHandle, "StmtHandle should be zero."
    Assert.AreSame dbc, DbStmt.DbConnection, "Connection object mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub



'@TestMethod("Query")
Private Sub ztcPrepare16V2_ThrowsOnClosedConnection()
    On Error Resume Next
    
    Dim dbc As SQLiteCConnection
    Set dbc = Fixtures.zfxGetConnDbMemory
    Dim DbStmt As SQLiteCStatement
    Set DbStmt = dbc.CreateStatement(vbNullString)
    
    Dim SQLQuery As String
    SQLQuery = Fixtures.zfxGetStmtSQLiteVersion
    Dim ResultCode As SQLiteResultCodes
    ResultCode = DbStmt.Prepare16V2(SQLQuery)
    Assert.AreEqual 0, DbStmt.StmtHandle, "StmtHandle should be zero."
    
    Guard.AssertExpectedError Assert, ConnectionNotOpenedErr
End Sub


'@TestMethod("Query")
Private Sub ztcPrepare16V2_VerifiesStatementVersionPeraration()
    On Error GoTo TestFail
    Set Fixtures = New SQLiteCTestFixtures

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = Fixtures.zfxGetConnDbMemory
    Dim DbStmt As SQLiteCStatement
    Set DbStmt = dbc.CreateStatement(vbNullString)
Act:
    Dim ResultCode As SQLiteResultCodes
Assert:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    
    Dim SQLQuery As String
    SQLQuery = Fixtures.zfxGetStmtSQLiteVersion
    
    ResultCode = DbStmt.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    Assert.AreNotEqual 0, DbStmt.StmtHandle, "StmtHandle should not be zero."
    Assert.IsNotNothing DbStmt.DbParameters, "Parameters object should be set."
    Assert.AreEqual SQLQuery, DbStmt.SQLQueryOriginal, "Original query mismatch"
    Assert.AreEqual SQLQuery, DbStmt.SQLQueryExpanded, "Expanded query mismatch"
    
    ResultCode = DbStmt.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    Assert.AreEqual 0, DbStmt.StmtHandle, "StmtHandle should be zero."
    
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Query")
Private Sub ztcPrepare16V2_VerifiesStatementCreateTablePeraration()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = Fixtures.zfxGetConnDbMemory
    Dim DbStmt As SQLiteCStatement
    Set DbStmt = dbc.CreateStatement(vbNullString)
    Dim ResultCode As SQLiteResultCodes
Act:
Assert:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    
    Dim SQLQuery As String
    SQLQuery = Fixtures.zfxGetStmtCreateTableIRBNT2V
    
    ResultCode = DbStmt.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    Assert.AreNotEqual 0, DbStmt.StmtHandle, "StmtHandle should not be zero."
    Assert.AreEqual SQLQuery, DbStmt.SQLQueryOriginal, "Original query mismatch"
    Assert.AreEqual SQLQuery, DbStmt.SQLQueryExpanded, "Expanded query mismatch"
    
    ResultCode = DbStmt.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    Assert.AreEqual 0, DbStmt.StmtHandle, "StmtHandle should be zero."
    
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Query")
Private Sub ztcPrepare16V2_VerifiesInvalidSQLHandling()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = Fixtures.zfxGetConnDbMemory
    Dim DbStmt As SQLiteCStatement
    Set DbStmt = dbc.CreateStatement(vbNullString)
    Dim ResultCode As SQLiteResultCodes
Act:
Assert:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    
    Dim SQLQuery As String
    
    SQLQuery = "SELECT --"
    ResultCode = DbStmt.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_ERROR, ResultCode, "Expected SQLITE_ERROR error: '" & SQLQuery & "'."
    Assert.AreEqual 0, DbStmt.StmtHandle, "StmtHandle should be zero: '" & SQLQuery & "'."
    
    SQLQuery = "-- SELECT --"
    ResultCode = DbStmt.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_OK, ResultCode, "Expected SQLITE_OK result: '" & SQLQuery & "'."
    Assert.AreEqual 0, DbStmt.StmtHandle, "StmtHandle should be zero: '" & SQLQuery & "'."
    
    SQLQuery = "ABC SELECT --"
    ResultCode = DbStmt.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_ERROR, ResultCode, "Expected SQLITE_ERROR error: '" & SQLQuery & "'."
    Assert.AreEqual 0, DbStmt.StmtHandle, "StmtHandle should be zero: '" & SQLQuery & "'."
    
    SQLQuery = "SELECT * FROM ABC"
    ResultCode = DbStmt.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_ERROR, ResultCode, "Expected SQLITE_ERROR error: '" & SQLQuery & "'."
    Assert.AreEqual 0, DbStmt.StmtHandle, "StmtHandle should be zero: '" & SQLQuery & "'."
    
    ResultCode = DbStmt.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    Assert.AreEqual 0, DbStmt.StmtHandle, "StmtHandle should be zero."
    
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Query")
Private Sub ztcPrepare16V2_VerifiesValidSelect()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = Fixtures.zfxGetConnDbMemoryWithTable
    Dim DbStmt As SQLiteCStatement
    Set DbStmt = dbc.CreateStatement(vbNullString)
    Dim ResultCode As SQLiteResultCodes
Act:
Assert:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    
    Dim SQLQuery As String
    
    SQLQuery = Fixtures.zfxGetStmtSELECTt1
    ResultCode = DbStmt.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_ERROR, ResultCode, "Expected SQLITE_ERROR error."
    Assert.AreEqual 0, DbStmt.StmtHandle, "StmtHandle should be zero."
    
    ResultCode = DbStmt.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    Assert.AreEqual 0, DbStmt.StmtHandle, "StmtHandle should be zero."
    
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
