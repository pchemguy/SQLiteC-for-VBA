Attribute VB_Name = "SQLiteCConnectionAccessModTests"
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
Private FixObj As SQLiteCTestFixObj


'This method runs once per module.
'@ModuleInitialize
Private Sub ModuleInitialize()
    #If LateBind Then
        Set Assert = CreateObject("Rubberduck.PermissiveAssertClass")
    #Else
        Set Assert = New Rubberduck.PermissiveAssertClass
    #End If
    Set FixObj = New SQLiteCTestFixObj
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


'@TestMethod("AccessMode")
Private Sub ztcAccessMode_VerifiesDefaultAccess()
    On Error GoTo TestFail

Arrange:
Act:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.GetConnDbMemory
    Dim ResultCode As SQLiteResultCodes
    Dim DbAccessMode As SQLiteDbAccess
    DbAccessMode = SQLITE_DB_NULL
Assert:
        ResultCode = dbc.OpenDb(SQLITE_OPEN_DEFAULT)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error"
        DbAccessMode = dbc.AccessMode("main")
    Assert.AreEqual SQLITE_DB_FULL, DbAccessMode, "Expected full db access mode"
        ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("AccessMode")
Private Sub ztcAccessMode_VerifiesReadAccess()
    On Error GoTo TestFail

Arrange:
Act:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.GetConnDbMemory
    Dim ResultCode As SQLiteResultCodes
    Dim DbAccessMode As SQLiteDbAccess
    DbAccessMode = SQLITE_DB_NULL
Assert:
        ResultCode = dbc.OpenDb(SQLITE_OPEN_READONLY)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error"
        DbAccessMode = dbc.AccessMode("main")
    Assert.AreEqual SQLITE_DB_READ, DbAccessMode, "Expected read db access mode"
        ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("AccessMode")
Private Sub ztcAccessMode_VerifiesDefaultAccessReadOnlyFile()
    On Error GoTo TestFail

Arrange:
Act:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.GetConnDbReadOnlyAttr
    Dim ResultCode As SQLiteResultCodes
    Dim DbAccessMode As SQLiteDbAccess
    DbAccessMode = SQLITE_DB_NULL
Assert:
        ResultCode = dbc.OpenDb(SQLITE_OPEN_DEFAULT)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error"
        DbAccessMode = dbc.AccessMode("main")
    Assert.AreEqual SQLITE_DB_READ, DbAccessMode, "Expected read db access mode"
        ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
