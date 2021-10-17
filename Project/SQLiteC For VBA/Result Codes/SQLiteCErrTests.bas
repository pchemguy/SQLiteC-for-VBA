Attribute VB_Name = "SQLiteCErrTests"
'@Folder "SQLiteC For VBA.Result Codes"
'@TestModule
'@IgnoreModule IndexedDefaultMemberAccess, FunctionReturnValueDiscarded, UseMeaningfulName
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed
'@IgnoreModule UnhandledOnErrorResumeNext
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
End Sub


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("Factory")
Private Sub ztcCreate_ThrowsOnNullConnection()
    On Error Resume Next
    Dim ErrorInfo As SQLiteCErr
    Set ErrorInfo = SQLiteCErr(Nothing)
    Guard.AssertExpectedError Assert, ErrNo.ObjectNotSetErr
End Sub


'@TestMethod("Factory")
Private Sub ztcCreate_VerifiesProperties()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.zfxGetConnDbRegular
    Dim dberr As SQLiteCErr
    Set dberr = dbc.ErrorInfo
Act:
Assert:
    Assert.IsNotNothing dbc.ErrorInfo, "ErrorInfo must be set."
    Assert.AreEqual SQLITE_OK, dberr.ErrorCode, "ErrorCode mismatch"
    Assert.AreEqual SQLITE_OK, dberr.ErrorCodeEx, "ErrorCodeEx mismatch"
    Assert.AreEqual "OK", dberr.ErrorName, "ErrorName mismatch"
    Assert.AreEqual "OK", dberr.ErrorCodeName, "ErrorCodeName mismatch"
    Assert.AreEqual "OK", dberr.ErrorCodeExName, "ErrorCodeExName mismatch"
    Assert.AreEqual vbNullString, dberr.ErrorMessage, "ErrorMessage mismatch"
    Assert.AreEqual vbNullString, dberr.ErrorString, "ErrorString mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Factory")
Private Sub ztcGetErr_VerifiesErrorInfo()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.zfxGetConnDbRegular
    dbc.ErrInfoRetrieve
    Dim dberr As SQLiteCErr
    Set dberr = dbc.ErrorInfo
Act:
Assert:
    Assert.AreEqual SQLITE_NOMEM, dberr.ErrorCode, "ErrorCode mismatch"
    Assert.AreEqual SQLITE_NOMEM, dberr.ErrorCodeEx, "ErrorCodeEx mismatch"
    Assert.AreEqual "NOMEM", dberr.ErrorName, "ErrorName mismatch"
    Assert.AreEqual "NOMEM", dberr.ErrorCodeName, "ErrorCodeName mismatch"
    Assert.AreEqual "NOMEM", dberr.ErrorCodeExName, "ErrorCodeExName mismatch"
    Assert.AreEqual "out of memory", dberr.ErrorMessage, "ErrorMessage mismatch"
    Assert.AreEqual "out of memory", dberr.ErrorString, "ErrorString mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
