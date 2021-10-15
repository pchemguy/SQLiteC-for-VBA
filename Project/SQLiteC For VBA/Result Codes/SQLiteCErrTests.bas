Attribute VB_Name = "SQLiteCErrTests"
'@Folder "SQLiteC For VBA.Result Codes"
'@TestModule
'@IgnoreModule IndexedDefaultMemberAccess, FunctionReturnValueDiscarded, UseMeaningfulName
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed
'@IgnoreModule UnhandledOnErrorResumeNext
Option Explicit
Option Private Module

Private Const LITE_LIB As String = "SQLiteCforVBA"
Private Const PATH_SEP As String = "\"
Private Const LITE_RPREFIX As String = "Library" & PATH_SEP & LITE_LIB & PATH_SEP

#Const LateBind = LateBindTests
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
'===================== FIXTURES ====================='
'===================================================='


Private Function zfxGetDefaultDBM() As SQLiteC
    Dim DllPath As String
    DllPath = LITE_RPREFIX & "dll\" & ARCH
    Dim dbm As SQLiteC
    '''' Using default library names hardcoded in the SQLiteC constructor.
    Set dbm = SQLiteC(DllPath)
    If dbm Is Nothing Then Err.Raise ErrNo.UnknownClassErr, _
        "SQLiteCTests", "Failed to create an SQLiteC instance."
    Set zfxGetDefaultDBM = dbm
End Function

Private Function zfxGetConnection(ByVal DbPathName As String) As SQLiteCConnection
    Dim dbm As SQLiteC
    Set dbm = zfxGetDefaultDBM()
    Dim DbConn As SQLiteCConnection
    Set DbConn = dbm.CreateConnection(DbPathName)
    If DbConn Is Nothing Then Err.Raise ErrNo.UnknownClassErr, _
        "SQLiteCTests", "Failed to create an SQLiteCConnection instance."
    Set zfxGetConnection = DbConn
End Function

Private Function zfxGetConnDbRegular() As SQLiteCConnection
    Dim DbPathName As String
    DbPathName = ThisWorkbook.Path & PATH_SEP & LITE_RPREFIX & LITE_LIB & ".db"
    Set zfxGetConnDbRegular = zfxGetConnection(DbPathName)
End Function


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
    Set dbc = zfxGetConnDbRegular
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
    Set dbc = zfxGetConnDbRegular
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
