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
Private Sub ztcBindDictOrArray_VerifiesQueryWithBlobLiteralAtParam()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.GetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)

    Dim TestChar As Byte
    TestChar = &H41
    Dim TestCharCount As Long
    TestCharCount = 5
    Dim TestStr As String
    TestStr = String(TestCharCount, Chr$(TestChar))
    Dim Expected As String
    Expected = Replace(String(TestCharCount, "*"), "*", FixObj.ByteToHex(TestChar))
    Expected = "SELECT x'" & Expected & "';"
Act:
    Dim ResultCode As SQLiteResultCodes
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    
    Dim SQLQuery As String
    SQLQuery = FixSQL.SELECTLiteralAtParam
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    ResultCode = dbs.DbParameters.BindDictOrArray(Array(FixObj.ByteArray(TestStr)))
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected BindDictOrArray error."
Assert:
    Assert.AreEqual Expected, dbs.SQLQueryExpanded, "Expanded query mismatch."
    
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
Private Sub ztcBindDictOrArray_VerifiesQueryWithLiteralAtParam()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.GetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
Act:
    Dim ResultCode As SQLiteResultCodes
Assert:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    
    Dim SQLQuery As String
    SQLQuery = FixSQL.SELECTLiteralAtParam
    
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    Assert.AreEqual 1, dbs.DbParameters.ParameterCount, "Named @ parameter count mismatch."
    Assert.AreEqual "SELECT NULL;", dbs.SQLQueryExpanded, "Template query mismatch."
    
    ResultCode = dbs.DbParameters.BindDictOrArray(Array(10241024))
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected BindDictOrArray error."
    Assert.AreEqual "SELECT 10241024;", dbs.SQLQueryExpanded, "Integer bound query mismatch."
    
    ResultCode = dbs.DbParameters.BindDictOrArray(Array(Null))
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected BindDictOrArray error."
    Assert.AreEqual "SELECT NULL;", dbs.SQLQueryExpanded, "Null bound query mismatch."
    
    ResultCode = dbs.DbParameters.BindDictOrArray(Array("ABC"))
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected BindDictOrArray error."
    Assert.AreEqual "SELECT 'ABC';", dbs.SQLQueryExpanded, "String bound query mismatch."
    
    ResultCode = dbs.DbParameters.BindDictOrArray(Array(1024.1024))
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected BindDictOrArray error."
    Assert.AreEqual "SELECT 1024.1024;", dbs.SQLQueryExpanded, "Real bound query mismatch."
    
    ResultCode = dbs.DbParameters.BindDictOrArray(Array(1024.1024@))
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected BindDictOrArray error."
    Assert.AreEqual "SELECT 10241024;", dbs.SQLQueryExpanded, "Currency bound query mismatch."
    
    ResultCode = dbs.DbParameters.BindDictOrArray(Array(True))
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected BindDictOrArray error."
    Assert.AreEqual "SELECT 1;", dbs.SQLQueryExpanded, "Boolean bound query mismatch."
    
    ResultCode = dbs.DbParameters.BindDictOrArray(Array(False))
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected BindDictOrArray error."
    Assert.AreEqual "SELECT 0;", dbs.SQLQueryExpanded, "Boolean bound query mismatch."
    
    ResultCode = dbs.DbParameters.BindDictOrArray(Array(CDec("123456789")))
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected BindDictOrArray error."
    Assert.AreEqual "SELECT '123456789';", dbs.SQLQueryExpanded, "Decimal bound query mismatch."
    
    ResultCode = dbs.DbParameters.BindDictOrArray(Array(FixObj.ByteArray("ABC")))
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected BindDictOrArray error."
    Assert.AreEqual "SELECT x'414243';", dbs.SQLQueryExpanded, "Blob bound query mismatch."
    
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
Private Sub ztcBindDictOrArray_VerifiesQueryWithAnonParams()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.GetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
Act:
    Dim ResultCode As SQLiteResultCodes
Assert:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    
    Dim SQLQuery As String
    SQLQuery = FixSQL.SELECTFunctionsNamedParamsAnon
    
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
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
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.GetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
Act:
    Dim ResultCode As SQLiteResultCodes
Assert:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    
    Dim SQLQuery As String
    SQLQuery = FixSQL.SELECTFunctionsNamedParamsNo
    
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    Assert.IsNotNothing dbs.DbParameters, "DbParameters should be set"
    Assert.AreEqual 6, dbs.DbParameters.ParameterCount, "Numbered parameter count mismatch."
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
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.GetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
Act:
    Dim ResultCode As SQLiteResultCodes
Assert:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    
    Dim SQLQuery As String
    SQLQuery = FixSQL.SELECTFunctionsNamedParamsCOL
    
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    Assert.IsNotNothing dbs.DbParameters, "DbParameters should be set"
    Assert.AreEqual 6, dbs.DbParameters.ParameterCount, "Named : parameter count mismatch."
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
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.GetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
Act:
    Dim ResultCode As SQLiteResultCodes
Assert:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    
    Dim SQLQuery As String
    SQLQuery = FixSQL.SELECTFunctionsNamedParamsS
    
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    Assert.IsNotNothing dbs.DbParameters, "DbParameters should be set"
    Assert.AreEqual 6, dbs.DbParameters.ParameterCount, "Named $ parameter count mismatch."
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
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.GetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
Act:
    Dim ResultCode As SQLiteResultCodes
Assert:
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    
    Dim SQLQuery As String
    SQLQuery = FixSQL.SELECTFunctionsNamedParamsAt
    
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    Assert.IsNotNothing dbs.DbParameters, "DbParameters should be set"
    Assert.AreEqual 6, dbs.DbParameters.ParameterCount, "Named @ parameter count mismatch."
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


'@TestMethod("Parameterized Query")
Private Sub ztcBindDictOrArray_VerifiesQueryWithAtParamsSeqValues()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.GetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)

    Dim ResultCode As SQLiteResultCodes
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
Act:
    Dim SQLQuery As String
    SQLQuery = FixSQL.SELECTFunctionsNamedParamsAt
    Dim SQLQueryExpandedUnbound As String
    SQLQueryExpandedUnbound = Replace(FixSQL.SELECTFunctionsNamedParamsAnon, _
        "?", "NULL")
Assert:
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    Assert.IsNotNothing dbs.DbParameters, "DbParameters should be set"
    Assert.AreEqual 6, dbs.DbParameters.ParameterCount, "Named @ parameter count mismatch."
    Assert.AreEqual SQLQueryExpandedUnbound, dbs.SQLQueryExpanded, "Expanded unbound query mismatch."
    
    ResultCode = dbs.DbParameters.BindDictOrArray(FixSQL.SELECTFunctionsNamedParamsAnonValues)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected BindDictOrArray error."
    Assert.AreEqual FixSQL.SELECTFunctionsTableWHERE, dbs.SQLQueryExpanded, "Expanded query mismatch."
Cleanup:
    ResultCode = dbs.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Metadata")
Private Sub ztcBindDictOrArray_ThrowsOnSequntialParamCountMismatch()
    On Error Resume Next
    
Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.GetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)

    Dim ResultCode As SQLiteResultCodes
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
Act:
    Dim SQLQuery As String
    SQLQuery = FixSQL.SELECTFunctionsNamedParamsAt
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    Dim ParamValues As Variant
    ParamValues = FixSQL.SELECTFunctionsNamedParamsAnonValues
    ResultCode = dbs.DbParameters.BindDictOrArray(Array(ParamValues(0), ParamValues(1)))

Assert:
    Guard.AssertExpectedError Assert, ErrNo.InvalidParameterErr
End Sub
