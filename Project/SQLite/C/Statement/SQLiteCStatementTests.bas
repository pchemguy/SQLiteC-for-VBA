Attribute VB_Name = "SQLiteCStatementTests"
'@Folder "SQLite.C.Statement"
'@TestModule
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed, UnhandledOnErrorResumeNext
Option Explicit
Option Private Module

Private Const MODULE_NAME As String = "SQLiteCStatementTests"
Private TestCounter As Long

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
    With Logger
        .ClearLog
        .DebugLevelDatabase = DEBUGLEVEL_MAX
        .DebugLevelImmediate = DEBUGLEVEL_NONE
        .UseIdPadding = True
        .UseTimeStamp = False
        .RecordIdDigits 3
        .TimerSet MODULE_NAME
    End With
    TestCounter = 0
End Sub


'This method runs once per module.
'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
    Logger.TimerLogClear MODULE_NAME, TestCounter
    Logger.PrintLog
    FixObjC.CleanUp
End Sub


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("DbStatement")
Private Sub ztcCreate_VerifiesNewStatement()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
Act:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
Assert:
    Assert.IsFalse dbs Is Nothing, "DbStmt is not set."
    Assert.IsFalse dbs.DbConnection Is Nothing, "Connection object not set."
    Assert.IsFalse dbs.DbExecutor Is Nothing, "Executor object not set."
    Assert.IsTrue dbs.DbParameters Is Nothing, "Parameters object should not be set."
    Assert.AreEqual 0, dbs.StmtHandle, "StmtHandle should be zero."
    Assert.AreSame dbc, dbs.DbConnection, "Connection object mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DbStatement")
Private Sub ztcExecADO_VerifiesILiteADOStatement()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
Act:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement("ILiteADO")
Assert:
    Assert.IsFalse dbs Is Nothing, "DbStmt is not set."
    Assert.IsTrue dbs.ExecADO Is dbc.ExecADO, "ExecADO reference mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Query")
Private Sub ztcPrepare16V2_ThrowsOnClosedConnection()
    On Error Resume Next
    TestCounter = TestCounter + 1
    
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
    
    Dim SQLQuery As String
    SQLQuery = LiteMetaSQL.Version
    Dim ResultCode As SQLiteResultCodes
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual 0, dbs.StmtHandle, "StmtHandle should be zero."
    
    Guard.AssertExpectedError Assert, ConnectionNotOpenedErr
End Sub


'@TestMethod("Query")
Private Sub ztcPrepare16V2_VerifiesPrepareSQLiteVersion()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
Act:
    Dim ResultCode As SQLiteResultCodes
Assert:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    
    Dim SQLQuery As String
    SQLQuery = LiteMetaSQL.Version
    
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    Assert.AreNotEqual 0, dbs.StmtHandle, "StmtHandle should not be zero."
    Assert.IsFalse dbs.DbParameters Is Nothing, "Parameters object should be set."
    Assert.AreEqual SQLQuery, dbs.SQLQueryOriginal, "Original query mismatch"
    Assert.AreEqual SQLQuery, dbs.SQLQueryExpanded, "Expanded query mismatch"
    
    ResultCode = dbs.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    Assert.AreEqual 0, dbs.StmtHandle, "StmtHandle should be zero."
    
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Query")
Private Sub ztcPrepare16V2_VerifiesPrepareOfCreateTable()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
    Dim ResultCode As SQLiteResultCodes
Act:
Assert:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    
    Dim SQLQuery As String
    SQLQuery = FixSQLITRB.Create
    
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    Assert.AreNotEqual 0, dbs.StmtHandle, "StmtHandle should not be zero."
    Assert.AreEqual SQLQuery, dbs.SQLQueryOriginal, "Original query mismatch"
    Assert.AreEqual SQLQuery, dbs.SQLQueryExpanded, "Expanded query mismatch"
    
    ResultCode = dbs.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    Assert.AreEqual 0, dbs.StmtHandle, "StmtHandle should be zero."
    
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Query")
Private Sub ztcPrepare16V2_VerifiesErrorOnInvalidSQL()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
    Dim ResultCode As SQLiteResultCodes
Act:
Assert:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    
    Dim SQLQuery As String
    
    SQLQuery = "SELECT --"
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_ERROR, ResultCode, "Expected SQLITE_ERROR error: '" & SQLQuery & "'."
    Assert.AreEqual 0, dbs.StmtHandle, "StmtHandle should be zero: '" & SQLQuery & "'."
    
    SQLQuery = "-- SELECT --"
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_OK, ResultCode, "Expected SQLITE_OK result: '" & SQLQuery & "'."
    Assert.AreEqual 0, dbs.StmtHandle, "StmtHandle should be zero: '" & SQLQuery & "'."
    
    SQLQuery = "ABC SELECT --"
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_ERROR, ResultCode, "Expected SQLITE_ERROR error: '" & SQLQuery & "'."
    Assert.AreEqual 0, dbs.StmtHandle, "StmtHandle should be zero: '" & SQLQuery & "'."
    
    SQLQuery = "SELECT * FROM ABC"
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_ERROR, ResultCode, "Expected SQLITE_ERROR error: '" & SQLQuery & "'."
    Assert.AreEqual 0, dbs.StmtHandle, "StmtHandle should be zero: '" & SQLQuery & "'."
    
    ResultCode = dbs.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    Assert.AreEqual 0, dbs.StmtHandle, "StmtHandle should be zero."
    
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Query")
Private Sub ztcGetBusy_VerifiesBusyStatus()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
Act:
    Dim ResultCode As SQLiteResultCodes
    Dim Result As Variant
Assert:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    
    Dim SQLQuery As String
    SQLQuery = LiteMetaSQL.Collations
    
        ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    Assert.AreEqual False, dbs.Busy, "Busy status should be False"
        Result = dbs.GetScalar(SQLQuery)
    Assert.AreEqual True, dbs.Busy, "Busy status should be True"
        ResultCode = dbs.Reset
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Reset error."
    Assert.AreEqual False, dbs.Busy, "Busy status should be False"
    
    ResultCode = dbs.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Query")
Private Sub ztcPrepare16V2_VerifiesGetScalar()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbm As SQLiteC
    Set dbm = FixObjC.GetDBM
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
Act:
    Dim ResultCode As SQLiteResultCodes
    Dim Result As Variant
Assert:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    
    Dim SQLQuery As String
    SQLQuery = LiteMetaSQL.Version
    
    Result = dbs.GetScalar(SQLQuery)
    Assert.AreEqual dbm.Version(False), Result, "GetScalar mismatch."
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Query")
Private Sub ztcExecuteNonQuery_ThrowsOnBlankQueryAndNullParams()
    On Error Resume Next
    TestCounter = TestCounter + 1
    
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
    
    Dim SQLQueryDummy As String
    SQLQueryDummy = FixSQLBase.SelectLiteralAtParam
    Dim SQLQuery As String
    SQLQuery = vbNullString
    Dim AffectedRows As Long
    AffectedRows = 0
    Dim QueryParams As Variant
    QueryParams = Empty
    Dim ResultCode As SQLiteResultCodes
    
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    ResultCode = dbs.Prepare16V2(SQLQueryDummy)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    
    Dim Result As Variant
    Result = dbs.ExecuteNonQuery(SQLQuery, QueryParams, AffectedRows)

    Guard.AssertExpectedError Assert, ErrNo.InvalidParameterErr
End Sub


'@TestMethod("Query")
Private Sub ztcExecuteNonQuery_ThrowsOnInvalidParamsType()
    On Error Resume Next
    TestCounter = TestCounter + 1
    
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
    
    Dim SQLQueryDummy As String
    SQLQueryDummy = FixSQLBase.SelectLiteralAtParam
    Dim SQLQuery As String
    SQLQuery = vbNullString
    Dim AffectedRows As Long
    AffectedRows = 0
    Dim QueryParams As Variant
    QueryParams = "ABC"
    Dim ResultCode As SQLiteResultCodes
    
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    ResultCode = dbs.Prepare16V2(SQLQueryDummy)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    
    Dim Result As Variant
    Result = dbs.ExecuteNonQuery(SQLQuery, QueryParams, AffectedRows)

    Guard.AssertExpectedError Assert, ErrNo.InvalidParameterErr
End Sub


'@TestMethod("Query")
Private Sub ztcExecuteNonQuery_ThrowsOnBlankQueryToUnpreparedStatement()
    On Error Resume Next
    TestCounter = TestCounter + 1
    
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
    
    Dim SQLQuery As String
    SQLQuery = vbNullString
    Dim AffectedRows As Long
    AffectedRows = 0
    Dim QueryParams As Variant
    QueryParams = Array("ABC")
    Dim ResultCode As SQLiteResultCodes
    
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    
    Dim Result As Variant
    Result = dbs.ExecuteNonQuery(SQLQuery, QueryParams, AffectedRows)

    Guard.AssertExpectedError Assert, ErrNo.InvalidParameterErr
End Sub


'@TestMethod("Query")
Private Sub ztcExecuteNonQuery_VerifiesCreateTable()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCTmp
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)

    Dim ResultCode As SQLiteResultCodes

    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
Act:
    Dim SQLQuery As String
    SQLQuery = FixSQLITRB.CreateRowid
    Dim AffectedRows As Long: AffectedRows = 0 '''' RD ByRef workaround.
    ResultCode = dbs.ExecuteNonQuery(SQLQuery, , AffectedRows)
    Assert.IsTrue ResultCode = SQLITE_OK Or ResultCode = SQLITE_DONE, _
        "Unexpected ExecuteNonQuery error."
Assert:
    Assert.AreEqual 0, AffectedRows, "AffectedRows mismatch"
    Assert.AreEqual SQLQuery, dbs.SQLQueryOriginal, "Original query mismatch"
CleanUp:
    ResultCode = dbs.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Finalize error."
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Query Paged RowSet")
Private Sub ztcGetPagedRowSet_VerifyPagedRowSetGeometry()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)

    Dim ResultCode As SQLiteResultCodes

    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    Dim PageSize As Long
    PageSize = 8
    Dim PageCount As Long
    PageCount = 28
    dbs.DbExecutor.PageSize = PageSize
    dbs.DbExecutor.PageCount = PageCount
    Dim AffectedRows As Long
    AffectedRows = FixObjC.CreateFunctionsTableWithData(dbc)
Act:
    Dim SQLQuery As String
    SQLQuery = FixSQLFunc.SelectMinMaxSubstrTrimParamName
    Dim QueryParams As Scripting.Dictionary
    Set QueryParams = FixSQLFunc.SelectMinMaxSubstrTrimParamNameValues
    Dim PagedRowSet As Variant
    PagedRowSet = dbs.GetPagedRowSet(SQLQuery, QueryParams, True)
Assert:
    Assert.IsFalse IsError(PagedRowSet), "Unexpected error from GetPagedRowSet."
    Assert.IsFalse IsEmpty(PagedRowSet), "GetPagedRowSet should not be empty."
    Assert.IsFalse IsNull(PagedRowSet), "GetPagedRowSet should not be null."
    Assert.AreEqual 0, LBound(PagedRowSet), "PagesArray base mismatch"
    Assert.AreEqual PageCount - 1, UBound(PagedRowSet), "PagesArray size mismatch"
    Assert.AreEqual 0, LBound(PagedRowSet(0)), "RowSet base mismatch"
    Assert.AreEqual PageSize - 1, UBound(PagedRowSet(0)), "RowSet size mismatch"
    Assert.AreEqual 0, LBound(PagedRowSet(0)(0)), "FieldSet base mismatch"
    Assert.AreEqual dbs.DbExecutor.ColumnCountAPI - 1, UBound(PagedRowSet(0)(0)), "FieldSet size mismatch"
CleanUp:
    ResultCode = dbs.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Finalize error."
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Query Paged RowSet")
Private Sub ztcGetPagedRowSet_SelectSubsetOfFunctions()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)

    Dim ResultCode As SQLiteResultCodes

    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    With dbs.DbExecutor
        .PageSize = 99
        .PageCount = 9
    End With
    Dim AffectedRows As Long
    AffectedRows = FixObjC.CreateFunctionsTableWithData(dbc)

    Dim SQLQuery As String
    SQLQuery = FixSQLFunc.SelectMinMaxSubstrTrimPlain
    Dim SQLQueryCount As String
    SQLQueryCount = SQLlib.CountSelect(SQLQuery)
    Dim RecordCount As Variant
    RecordCount = dbs.GetScalar(SQLQueryCount)
Act:
    SQLQuery = FixSQLFunc.SelectMinMaxSubstrTrimParamName
    Dim QueryParams As Scripting.Dictionary
    Set QueryParams = FixSQLFunc.SelectMinMaxSubstrTrimParamNameValues
    Dim PagedRowSet As Variant
    PagedRowSet = dbs.GetPagedRowSet(SQLQuery, QueryParams, True)
Assert:
    Assert.IsFalse IsEmpty(PagedRowSet(0)(RecordCount - 1)), "RowSet is too small"
    Assert.IsTrue IsEmpty(PagedRowSet(0)(RecordCount)), "RowSet is too big"
CleanUp:
    ResultCode = dbs.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Finalize error."
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Query 2D RowSet")
Private Sub ztcGetRowSet2D_VerifyRowSet2DGeometry()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)

    Dim ResultCode As SQLiteResultCodes

    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    Dim AffectedRows As Long
    AffectedRows = FixObjC.CreateFunctionsTableWithData(dbc)
Act:
    Dim SQLQuery As String
    SQLQuery = FixSQLFunc.SelectNoRowid
    Dim RowSet2D As Variant
    RowSet2D = dbs.GetRowSet2D(SQLQuery)
Assert:
    Assert.IsFalse IsError(RowSet2D), "Unexpected error from RowSet2D."
    Assert.IsFalse IsEmpty(RowSet2D), "RowSet2D should not be empty."
    Assert.IsFalse IsNull(RowSet2D), "RowSet2D should not be null."
    Assert.AreEqual SQLQuery, dbs.SQLQueryOriginal, "Original query mismatch"
    Assert.AreEqual 0, LBound(RowSet2D, 1), "RowSet2D R-base mismatch"
    Assert.AreEqual 0, LBound(RowSet2D, 2), "RowSet2D C-base mismatch"
    Assert.AreEqual dbs.DbExecutor.RowCount - 1, UBound(RowSet2D, 1), "RowSet2D R-size mismatch"
    Assert.AreEqual dbs.DbExecutor.ColumnCountAPI - 1, UBound(RowSet2D, 2), "RowSet2D C-size mismatch"
CleanUp:
    ResultCode = dbs.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Finalize error."
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Query 2D RowSet")
Private Sub ztcGetRowSet2D_NamedParamsSelectWithDictVsArrayValues()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)

    Dim ResultCode As SQLiteResultCodes

    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    Dim AffectedRows As Long
    AffectedRows = FixObjC.CreateFunctionsTableWithData(dbc)
    Assert.IsTrue AffectedRows > 0, "Failed to INSERT test data."
Act:
    Dim SQLQuery As String
    SQLQuery = FixSQLFunc.SelectMinMaxSubstrTrimParamName
    Dim QueryParamsDict As Scripting.Dictionary
    Set QueryParamsDict = FixSQLFunc.SelectMinMaxSubstrTrimParamNameValues
    Dim RowSet2DNamedParams As Variant
    RowSet2DNamedParams = dbs.GetRowSet2D(SQLQuery, QueryParamsDict, True)
    Assert.IsFalse IsError(RowSet2DNamedParams), "Unexpected GetRowSet2D error."
    Dim QueryParamsArray As Variant
    QueryParamsArray = FixSQLFunc.SelectMinMaxSubstrTrimParamAnonValues
    Dim RowSet2DAnonValues As Variant
    RowSet2DAnonValues = dbs.GetRowSet2D(SQLQuery, QueryParamsArray, True)
Assert:
    Assert.AreEqual UBound(RowSet2DNamedParams, 1), UBound(RowSet2DAnonValues, 1), "Record size mismatch."
    Assert.AreEqual UBound(RowSet2DNamedParams, 2), UBound(RowSet2DAnonValues, 2), "Column size mismatch."
    Assert.AreEqual RowSet2DNamedParams(0, 0), RowSet2DAnonValues(0, 0), "Bottom-Left mismatch."
CleanUp:
    ResultCode = dbs.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Finalize error."
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Query 2D RowSet")
Private Sub ztcGetRowSet2D_SelectPragmaTable()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)

    Dim ResultCode As SQLiteResultCodes

    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    Dim AffectedRows As Long
    AffectedRows = FixObjC.CreateFunctionsTableWithData(dbc)
Act:
    Dim SQLQuery As String
    SQLQuery = FixSQLFunc.SelectPragmaNoRowid
    Dim RowSet2D As Variant
Assert:
    RowSet2D = dbs.GetRowSet2D(SQLQuery)
    Assert.IsFalse IsError(RowSet2D), "Unexpected error from RowSet2D."
    Assert.IsTrue IsArray(RowSet2D), "Expected a rowset result."
CleanUp:
    ResultCode = dbs.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Finalize error."
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Query 2D RowSet")
Private Sub ztcGetRowSet2D_InsertPlainSelectFromITRBTable()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)

    Dim ResultCode As SQLiteResultCodes
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    Dim AffectedRows As Long
Act:
    Dim SQLQuery As String
    SQLQuery = FixSQLITRB.CreateWithData
    ResultCode = dbc.ExecuteNonQueryPlain(SQLQuery, AffectedRows)
    Assert.AreEqual 5, AffectedRows, "AffectedRows mismatch"
    
    SQLQuery = FixSQLITRB.SelectRowid
    Dim RowSet2D As Variant
    RowSet2D = dbs.GetRowSet2D(SQLQuery)
Assert:
    Assert.IsFalse IsError(RowSet2D), "Unexpected error from RowSet2D."
    Assert.IsFalse IsEmpty(RowSet2D), "RowSet2D should not be empty."
    Assert.IsFalse IsNull(RowSet2D), "RowSet2D should not be null."
    Assert.AreEqual SQLQuery, dbs.SQLQueryOriginal, "Original query mismatch"
    Assert.AreEqual 0, LBound(RowSet2D, 1), "RowSet2D R-base mismatch"
    Assert.AreEqual 0, LBound(RowSet2D, 2), "RowSet2D C-base mismatch"
    Assert.AreEqual 4, UBound(RowSet2D, 1), "RowSet2D R-size mismatch"
    Assert.AreEqual 5, UBound(RowSet2D, 2), "RowSet2D C-size mismatch"
    Assert.AreEqual dbs.DbExecutor.RowCount - 1, UBound(RowSet2D, 1), "RowSet2D R-size mismatch"
    Assert.AreEqual dbs.DbExecutor.ColumnCountAPI - 1, UBound(RowSet2D, 2), "RowSet2D C-size mismatch"
CleanUp:
    ResultCode = dbs.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Finalize error."
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Query ADO Recordset")
Private Sub ztcGetRecordset_VerifyGetRecordsetGeometry()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)

    Dim ResultCode As SQLiteResultCodes

    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    Dim AffectedRows As Long
    AffectedRows = FixObjC.CreateFunctionsTableWithData(dbc)
Act:
    Dim SQLQuery As String
    SQLQuery = FixSQLFunc.SelectNoRowid
    Dim dbr As SQLiteCRecordsetADO
    Set dbr = dbs.GetRecordset(SQLQuery)
Assert:
    Assert.IsFalse dbr Is Nothing, "Unexpected error from FabRecordset."
    Assert.AreEqual dbs.DbExecutor.RowCount, dbr.AdoRecordset.RecordCount, "Recordset.RecordCount mismatch"
    Assert.AreEqual dbs.DbExecutor.PageSize, dbr.AdoRecordset.PageSize, "Recordset.PageSize mismatch"
    Assert.AreEqual dbs.DbExecutor.PageSize, dbr.AdoRecordset.CacheSize, "Recordset.CacheSize mismatch"
    Assert.AreEqual adUseClient, dbr.AdoRecordset.CursorLocation, "CursorLocation mismatch"
    Assert.AreEqual adOpenStatic, dbr.AdoRecordset.CursorType, "CursorType mismatch"
    Assert.AreEqual adLockBatchOptimistic, dbr.AdoRecordset.LockType, "LockType mismatch"
    Assert.AreEqual dbs.DbExecutor.ColumnCountAPI, dbr.AdoRecordset.Fields.Count, "Fields.Count mismatch"
CleanUp:
    ResultCode = dbs.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Finalize error."
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Query ADO Recordset")
Private Sub ztcGetRecordset_InsertPlainSelectFromITRBTable()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)

    Dim ResultCode As SQLiteResultCodes
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    dbs.DbExecutor.PageSize = 3
    Dim AffectedRows As Long
Act:
    Dim SQLQuery As String
    SQLQuery = FixSQLITRB.CreateWithData
    ResultCode = dbc.ExecuteNonQueryPlain(SQLQuery, AffectedRows)
    Assert.AreEqual 5, AffectedRows, "AffectedRows mismatch"
    
    SQLQuery = FixSQLITRB.SelectNoRowid
    Dim dbr As SQLiteCRecordsetADO
    Set dbr = dbs.GetRecordset(SQLQuery)
Assert:
    Assert.IsFalse dbr Is Nothing, "Unexpected error from FabRecordset."
    Assert.AreEqual SQLQuery, dbs.SQLQueryOriginal, "Original query mismatch"
    Assert.AreEqual 5, dbr.AdoRecordset.RecordCount, "Recordset.RecordCount mismatch"
    Assert.AreEqual 5, dbr.AdoRecordset.Fields.Count, "Fields.Count mismatch"
    Assert.AreEqual 3, dbr.AdoRecordset.PageSize, "Recordset.PageSize mismatch"
    Assert.AreEqual 3, dbr.AdoRecordset.CacheSize, "Recordset.CacheSize mismatch"
CleanUp:
    ResultCode = dbs.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Finalize error."
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Query ADO Recordset")
Private Sub ztcGetRecordset_InsertPlainSelectFromITRBTableRowid()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)

    Dim ResultCode As SQLiteResultCodes
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    dbs.DbExecutor.PageSize = 3
    Dim AffectedRows As Long
Act:
    Dim SQLQuery As String
    SQLQuery = FixSQLITRB.CreateRowidWithData
    ResultCode = dbc.ExecuteNonQueryPlain(SQLQuery, AffectedRows)
    Assert.AreEqual 5, AffectedRows, "AffectedRows mismatch"
    
    SQLQuery = FixSQLITRB.SelectRowid
    Dim dbr As SQLiteCRecordsetADO
    Set dbr = dbs.GetRecordset(SQLQuery)
Assert:
    Assert.IsFalse dbr Is Nothing, "Unexpected error from FabRecordset."
    Assert.AreEqual SQLQuery, dbs.SQLQueryOriginal, "Original query mismatch"
    Assert.AreEqual 5, dbr.AdoRecordset.RecordCount, "Recordset.RecordCount mismatch"
    Assert.AreEqual 6, dbr.AdoRecordset.Fields.Count, "Fields.Count mismatch"
    Assert.AreEqual 3, dbr.AdoRecordset.PageSize, "Recordset.PageSize mismatch"
    Assert.AreEqual 3, dbr.AdoRecordset.CacheSize, "Recordset.CacheSize mismatch"
CleanUp:
    ResultCode = dbs.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Finalize error."
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Query ADO Recordset")
Private Sub ztcGetRecordset_InsertPlainSelectFromITRBTableRowidVerifyData()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)

    Dim ResultCode As SQLiteResultCodes
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    dbs.DbExecutor.PageSize = 3
    Dim AffectedRows As Long
Act:
    Dim SQLQuery As String
    SQLQuery = FixSQLITRB.CreateRowidWithData
    ResultCode = dbc.ExecuteNonQueryPlain(SQLQuery, AffectedRows)
    Assert.AreEqual 5, AffectedRows, "AffectedRows mismatch"
    
    SQLQuery = FixSQLITRB.SelectRowid
    Dim RowSet2D As Variant
    RowSet2D = dbs.GetRowSet2D(SQLQuery)
    Dim RowSet2DRst As Variant
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = dbs.GetRecordset(SQLQuery).AdoRecordset
    AdoRecordset.MoveFirst
    RowSet2DRst = ArrayLib.TransposeArray(AdoRecordset.GetRows)
Assert:
    Assert.AreEqual LBound(RowSet2D, 1), LBound(RowSet2DRst, 1), "R-base mismatch"
    Assert.AreEqual LBound(RowSet2D, 2), LBound(RowSet2DRst, 2), "C-base mismatch"
    Assert.AreEqual UBound(RowSet2D, 1), UBound(RowSet2DRst, 1), "R-size mismatch"
    Assert.AreEqual UBound(RowSet2D, 2), UBound(RowSet2DRst, 2), "C-size mismatch"
    Assert.AreEqual RowSet2DRst(0, 3), RowSet2D(0, 3), "Value mismatch"
    Assert.AreEqual RowSet2DRst(0, 5)(2), RowSet2D(0, 5)(2), "Value mismatch"
CleanUp:
    ResultCode = dbs.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Finalize error."
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("INSERT")
Private Sub ztcGetRowSet2D_InsertWithParamsSelectFromITRBTable()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCTmp
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)

    Dim ResultCode As SQLiteResultCodes
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    Dim AffectedRows As Long
Act:
    Dim SQLQuery As String
    SQLQuery = FixSQLITRB.Create
    ResultCode = dbc.ExecuteNonQueryPlain(SQLQuery, AffectedRows)
    Assert.AreEqual 0, AffectedRows, "AffectedRows mismatch"
    
    SQLQuery = FixSQLITRB.InsertParamName
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    
    Dim ParamNames As Variant
    ParamNames = FixSQLITRB.InsertParamNameNames
    Dim ParamValuesSets As Variant
    ParamValuesSets = FixSQLITRB.InsertParamNameValueSets
    Dim ParamValueMap As Scripting.Dictionary
    Dim RowIndex As Long
    For RowIndex = LBound(ParamValuesSets) To UBound(ParamValuesSets)
        Set ParamValueMap = FixUtils.KeysValuesToDict(ParamNames, ParamValuesSets(RowIndex))
        ResultCode = dbs.ExecuteNonQuery(vbNullString, ParamValueMap, AffectedRows)
        Assert.AreEqual SQLITE_DONE, ResultCode, "Unexpected ExecuteNonQuery error."
        Assert.AreEqual 1, AffectedRows, "AffectedRows mismatch"
    Next RowIndex
    
    SQLQuery = FixSQLITRB.SelectRowid
    Dim RowSet2D As Variant
    RowSet2D = dbs.GetRowSet2D(SQLQuery)
Assert:
    Assert.IsFalse IsError(RowSet2D), "Unexpected error from RowSet2D."
    Assert.IsFalse IsEmpty(RowSet2D), "RowSet2D should not be empty."
    Assert.IsFalse IsNull(RowSet2D), "RowSet2D should not be null."
    Assert.AreEqual SQLQuery, dbs.SQLQueryOriginal, "Original query mismatch"
    Assert.AreEqual 0, LBound(RowSet2D, 1), "RowSet2D R-base mismatch"
    Assert.AreEqual 0, LBound(RowSet2D, 2), "RowSet2D C-base mismatch"
    Assert.AreEqual 4, UBound(RowSet2D, 1), "RowSet2D R-size mismatch"
    Assert.AreEqual 5, UBound(RowSet2D, 2), "RowSet2D C-size mismatch"
    Assert.AreEqual dbs.DbExecutor.RowCount - 1, UBound(RowSet2D, 1), "RowSet2D R-size mismatch"
    Assert.AreEqual dbs.DbExecutor.ColumnCountAPI - 1, UBound(RowSet2D, 2), "RowSet2D C-size mismatch"
    Assert.AreEqual 7, UBound(RowSet2D(0, 5)), "Blob size mismatch."
    Assert.AreEqual 79, FixUtils.XorElements(RowSet2D(0, 5)), "Blob XOR hash mismatch"
CleanUp:
    ResultCode = dbs.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Finalize error."
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("UPDATE")
Private Sub ztcGetRowSet2D_UpdateWithParamsSelectFromITRBTable()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)

    Dim ResultCode As SQLiteResultCodes
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    Dim AffectedRows As Long
Act:
    Dim SQLQuery As String
    SQLQuery = FixSQLITRB.CreateWithData
    ResultCode = dbc.ExecuteNonQueryPlain(SQLQuery, AffectedRows)
    Assert.AreEqual 5, AffectedRows, "AffectedRows mismatch CREATE INSERT <" & CStr(AffectedRows) & ">"
    
    SQLQuery = FixSQLITRB.UpdateParamName
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    
    Dim ParamValueMap As Scripting.Dictionary
    Set ParamValueMap = FixSQLITRB.UpdateParamValueDict
    ResultCode = dbs.ExecuteNonQuery(vbNullString, ParamValueMap, AffectedRows)
    Assert.AreEqual SQLITE_DONE, ResultCode, "Unexpected ExecuteNonQuery error."
    Assert.AreEqual 3, AffectedRows, "AffectedRows mismatch <" & CStr(AffectedRows) & ">"
    
    SQLQuery = FixSQLITRB.SelectRowid
    Dim RowSet2D As Variant
    RowSet2D = dbs.GetRowSet2D(SQLQuery)
Assert:
    Assert.IsFalse IsError(RowSet2D), "Unexpected error from RowSet2D."
    Assert.IsFalse IsEmpty(RowSet2D), "RowSet2D should not be empty."
    Assert.IsFalse IsNull(RowSet2D), "RowSet2D should not be null."
    Assert.AreEqual SQLQuery, dbs.SQLQueryOriginal, "Original query mismatch"
    Assert.AreEqual 0, LBound(RowSet2D, 1), "RowSet2D R-base mismatch"
    Assert.AreEqual 0, LBound(RowSet2D, 2), "RowSet2D C-base mismatch"
    Assert.AreEqual 4, UBound(RowSet2D, 1), "RowSet2D R-size mismatch"
    Assert.AreEqual 5, UBound(RowSet2D, 2), "RowSet2D C-size mismatch"
    Assert.AreEqual dbs.DbExecutor.RowCount - 1, UBound(RowSet2D, 1), "RowSet2D R-size mismatch"
    Assert.AreEqual dbs.DbExecutor.ColumnCountAPI - 1, UBound(RowSet2D, 2), "RowSet2D C-size mismatch"
    Assert.AreEqual 14.4, RowSet2D(2, 4), "Control value mismatch."
CleanUp:
    ResultCode = dbs.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Finalize error."
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Query")
Private Sub ztcGetScalar_VerifiesGetScalar()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMemITRB
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
    Dim ResultCode As SQLiteResultCodes

    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
Act:
    Dim SQLQuery As String
    SQLQuery = "SELECT name FROM pragma_database_list()"
    Dim Result As Variant
    Result = dbs.GetScalar(SQLQuery)
Assert:
    Assert.AreEqual "main", Result, "Expected 'name'"
CleanUp:
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
