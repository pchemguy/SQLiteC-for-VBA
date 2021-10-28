Attribute VB_Name = "FixObjCTests"
'@Folder "SQLite.Fixtures"
'@TestModule
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed
'@IgnoreModule IndexedDefaultMemberAccess, UnhandledOnErrorResumeNext
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


'@TestMethod("Fixture")
Private Sub ztcGetDBCTempFuncWithData_VerifiesTempDatabase()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixMain.ObjC.GetDBCTempFuncWithData
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
    Dim ResultCode As SQLiteResultCodes

    Assert.AreEqual SQLITE_OK, dbc.OpenDb, "Unexpected OpenDb error."
    
    Dim SQLQuery As String
    SQLQuery = "SELECT count(*) FROM pragma_function_list()"
    Dim Expected As Variant
    Expected = dbs.GetScalar(SQLQuery)
    Assert.IsTrue IsNumeric(Expected), "Unexpected query result."
    SQLQuery = LiteMetaSQL("main").Tables
    Dim RowSet2D As Variant
    RowSet2D = dbs.GetRowSet2D(SQLQuery)
Act:
    SQLQuery = "SELECT count(*) FROM functions"
    Dim Actual As Variant
    Actual = dbs.GetScalar(SQLQuery)
Assert:
    Assert.IsTrue IsNumeric(Actual), "Unexpected query result."
    Assert.AreEqual Expected, Actual, ""
Cleanup:
    Assert.AreEqual SQLITE_OK, dbc.CloseDb, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Fixture")
Private Sub ztcGetDBCMemFuncWithData_VerifiesMemDatabase()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixMain.ObjC.GetDBCMemFuncWithData
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)
    Dim ResultCode As SQLiteResultCodes

    Assert.IsTrue dbc.DbHandle > 0, "Expected opened database."
    
    Dim SQLQuery As String
    SQLQuery = "SELECT count(*) FROM pragma_function_list()"
    Dim Expected As Variant
    Expected = dbs.GetScalar(SQLQuery)
    Assert.IsTrue IsNumeric(Expected), "Unexpected query result."
    
    SQLQuery = LiteMetaSQL("main").Tables
    Dim RowSet2D As Variant
    RowSet2D = dbs.GetRowSet2D(SQLQuery)
    Assert.IsTrue IsArray(RowSet2D), "Unexpected query result."
    Assert.AreEqual "functions", RowSet2D(0, 0), "Unexpected query result."
Act:
    SQLQuery = "SELECT count(*) FROM functions"
    Dim Actual As Variant
    Actual = dbs.GetScalar(SQLQuery)
Assert:
    Assert.IsTrue IsNumeric(Actual), "Unexpected query result."
    Assert.AreEqual Expected, Actual, ""
Cleanup:
    Assert.AreEqual SQLITE_OK, dbc.CloseDb, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
