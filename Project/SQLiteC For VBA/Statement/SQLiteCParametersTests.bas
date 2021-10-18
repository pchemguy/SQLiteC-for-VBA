Attribute VB_Name = "SQLiteCParametersTests"
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
Private FixObj As SQLiteCTestFixObj
Private FixSQL As SQLiteCTestFixSQL


'This method runs once per module.
'@ModuleInitialize
Private Sub ModuleInitialize()
    #If LateBind Then
        Set Assert = CreateObject("Rubberduck.PermissiveAssertClass")
    #Else
        Set Assert = New Rubberduck.PermissiveAssertClass
    #End If
    Set FixObj = New SQLiteCTestFixObj
    Set FixSQL = New SQLiteCTestFixSQL
End Sub


'This method runs once per module.
'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
    Set FixObj = Nothing
End Sub


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("Parameterized Query")
Private Sub ztcBindDictOrArray_VerifiesQueryWithAnonParams()
    On Error GoTo TestFail

Arrange:
    Dim dbm As SQLiteC
    Set dbm = FixObj.zfxGetDefaultDBM
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.zfxGetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
Act:
    Dim ResultCode As SQLiteResultCodes
    Dim Result As Variant
Assert:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    
    Dim SQLQuery As String
    SQLQuery = FixSQL.SELECTFunctionsNamedParamsAnon
    
    Result = dbs.Prepare16V2(SQLQuery)
    Assert.IsNotNothing dbs.DbParameters, "DbParameters should be set"
    Assert.AreEqual 6, dbs.DbParameters.ParameterCount, "Anon parameter count mismatch."
    ResultCode = dbs.DbParameters.BindDictOrArray(FixSQL.SELECTFunctionsNamedParamsAnonValues)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected BindDictOrArray error."
    Assert.AreEqual FixSQL.SELECTFunctionsTableWHERE, dbs.SQLQueryExpanded, "Expanded query mismatch."
    
    ResultCode = dbs.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Parameterized Query")
Private Sub ztcBindDictOrArray_VerifiesQueryWithNumberedParams()
    On Error GoTo TestFail

Arrange:
    Dim dbm As SQLiteC
    Set dbm = FixObj.zfxGetDefaultDBM
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.zfxGetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
Act:
    Dim ResultCode As SQLiteResultCodes
    Dim Result As Variant
Assert:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    
    Dim SQLQuery As String
    SQLQuery = FixSQL.SELECTFunctionsNamedParamsNo
    
    Result = dbs.Prepare16V2(SQLQuery)
    Assert.IsNotNothing dbs.DbParameters, "DbParameters should be set"
    Assert.AreEqual 6, dbs.DbParameters.ParameterCount, "Anon parameter count mismatch."
    ResultCode = dbs.DbParameters.BindDictOrArray(FixSQL.SELECTFunctionsNamedParamsNoValues)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected BindDictOrArray error."
    Assert.AreEqual FixSQL.SELECTFunctionsTableWHERE, dbs.SQLQueryExpanded, "Expanded query mismatch."
    
    ResultCode = dbs.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Parameterized Query")
Private Sub ztcBindDictOrArray_VerifiesQueryWithColonParams()
    On Error GoTo TestFail

Arrange:
    Dim dbm As SQLiteC
    Set dbm = FixObj.zfxGetDefaultDBM
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.zfxGetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
Act:
    Dim ResultCode As SQLiteResultCodes
    Dim Result As Variant
Assert:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    
    Dim SQLQuery As String
    SQLQuery = FixSQL.SELECTFunctionsNamedParamsCOL
    
    Result = dbs.Prepare16V2(SQLQuery)
    Assert.IsNotNothing dbs.DbParameters, "DbParameters should be set"
    Assert.AreEqual 6, dbs.DbParameters.ParameterCount, "Anon parameter count mismatch."
    ResultCode = dbs.DbParameters.BindDictOrArray(FixSQL.SELECTFunctionsNamedParamsCOLValues)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected BindDictOrArray error."
    Assert.AreEqual FixSQL.SELECTFunctionsTableWHERE, dbs.SQLQueryExpanded, "Expanded query mismatch."
    
    ResultCode = dbs.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Parameterized Query")
Private Sub ztcBindDictOrArray_VerifiesQueryWithSParams()
    On Error GoTo TestFail

Arrange:
    Dim dbm As SQLiteC
    Set dbm = FixObj.zfxGetDefaultDBM
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.zfxGetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
Act:
    Dim ResultCode As SQLiteResultCodes
    Dim Result As Variant
Assert:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    
    Dim SQLQuery As String
    SQLQuery = FixSQL.SELECTFunctionsNamedParamsS
    
    Result = dbs.Prepare16V2(SQLQuery)
    Assert.IsNotNothing dbs.DbParameters, "DbParameters should be set"
    Assert.AreEqual 6, dbs.DbParameters.ParameterCount, "Anon parameter count mismatch."
    ResultCode = dbs.DbParameters.BindDictOrArray(FixSQL.SELECTFunctionsNamedParamsSValues)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected BindDictOrArray error."
    Assert.AreEqual FixSQL.SELECTFunctionsTableWHERE, dbs.SQLQueryExpanded, "Expanded query mismatch."
    
    ResultCode = dbs.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Parameterized Query")
Private Sub ztcBindDictOrArray_VerifiesQueryWithAtParams()
    On Error GoTo TestFail

Arrange:
    Dim dbm As SQLiteC
    Set dbm = FixObj.zfxGetDefaultDBM
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.zfxGetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
Act:
    Dim ResultCode As SQLiteResultCodes
    Dim Result As Variant
Assert:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    
    Dim SQLQuery As String
    SQLQuery = FixSQL.SELECTFunctionsNamedParamsAt
    
    Result = dbs.Prepare16V2(SQLQuery)
    Assert.IsNotNothing dbs.DbParameters, "DbParameters should be set"
    Assert.AreEqual 6, dbs.DbParameters.ParameterCount, "Anon parameter count mismatch."
    ResultCode = dbs.DbParameters.BindDictOrArray(FixSQL.SELECTFunctionsNamedParamsAtValues)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected BindDictOrArray error."
    Assert.AreEqual FixSQL.SELECTFunctionsTableWHERE, dbs.SQLQueryExpanded, "Expanded query mismatch."
    
    ResultCode = dbs.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
