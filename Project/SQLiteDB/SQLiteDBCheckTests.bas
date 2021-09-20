Attribute VB_Name = "SQLiteDBCheckTests"
'@Folder "SQLiteDB"
'@TestModule
'@IgnoreModule LineLabelNotUsed, UnhandledOnErrorResumeNext, FunctionReturnValueDiscarded
'@IgnoreModule IndexedDefaultMemberAccess
Option Explicit
Option Private Module

#Const LateBind = LateBindTests
#If LateBind Then
    Private Assert As Object
#Else
    Private Assert As Rubberduck.PermissiveAssertClass
#End If

Private Const LIB_NAME As String = "SQLiteDBVBA"
Private Const PATH_SEP As String = "\"
Private Const REL_PREFIX As String = "Library" & PATH_SEP & LIB_NAME & PATH_SEP


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


'This method runs after every test in the module.
'@TestCleanup
Private Sub TestCleanup()
    Err.Clear
End Sub


'===================================================='
'===================== FIXTURES ====================='
'===================================================='


Private Function zfxDefaultDbManager() As SQLiteDBCheck
    Dim FilePathName As String
    FilePathName = REL_PREFIX & LIB_NAME & ".db"
    
    Dim dbm As SQLiteDBCheck
    Set dbm = SQLiteDBCheck.Create(FilePathName)
    
    Set zfxDefaultDbManager = dbm
End Function


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("Integrity checking")
Private Sub ztcCheckIntegrity_PassesDefaultDatabaseIntegrityCheck()
    On Error GoTo TestFail

Arrange:
    Dim dbm As SQLiteDBCheck
    Set dbm = zfxDefaultDbManager()
Act:
    Dim CheckResult As Boolean
    CheckResult = dbm.CheckIntegrity
Assert:
    Assert.IsTrue CheckResult, "Integrity check on default database failed"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Integrity checking")
Private Sub ztcCheckIntegrity_ThrowsOnFileNotDatabase()
    On Error Resume Next
    Dim dbm As SQLiteDBCheck
    Set dbm = SQLiteDBCheck(ThisWorkbook.Name)
    dbm.CheckIntegrity
    Guard.AssertExpectedError Assert, ErrNo.AdoInvalidFileFormatErr
End Sub


'@TestMethod("Integrity checking")
Private Sub ztcCheckIntegrity_ThrowsOnCorruptedDatabase()
    On Error Resume Next
    Dim dbm As SQLiteDBCheck
    Set dbm = SQLiteDBCheck(REL_PREFIX & "ICfailFKCfail.db")
    dbm.CheckIntegrity
    Guard.AssertExpectedError Assert, ErrNo.IntegrityCheckErr
End Sub


'@TestMethod("Integrity checking")
Private Sub ztcCheckIntegrity_ThrowsOnFailedFKCheck()
    On Error Resume Next
    Dim dbm As SQLiteDBCheck
    Set dbm = SQLiteDBCheck(REL_PREFIX & "ICokFKCfail.db")
    dbm.CheckIntegrity
    Guard.AssertExpectedError Assert, ErrNo.ConsistencyCheckErr
End Sub

