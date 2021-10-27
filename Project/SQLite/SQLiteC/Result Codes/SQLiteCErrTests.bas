Attribute VB_Name = "SQLiteCErrTests"
'@Folder "SQLite.SQLiteC.Result Codes"
'@TestModule
'@IgnoreModule IndexedDefaultMemberAccess, FunctionReturnValueDiscarded, UseMeaningfulName
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed
'@IgnoreModule UnhandledOnErrorResumeNext
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
    Set dbc = FixMain.ObjC.GetDBCReg
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
Private Sub ztcGetErr_ThrowsOnClosedConnection()
    On Error Resume Next
    
    Dim dbc As SQLiteCConnection
    Set dbc = FixMain.ObjC.GetDBCReg
    dbc.ErrInfoRetrieve
    Dim dberr As SQLiteCErr
    Set dberr = dbc.ErrorInfo
    
    Guard.AssertExpectedError Assert, ConnectionNotOpenedErr
End Sub
