Attribute VB_Name = "LiteACIDTests"
'@Folder "SQLite.Checks"
'@TestModule
'@IgnoreModule LineLabelNotUsed, UnhandledOnErrorResumeNext, FunctionReturnValueDiscarded
'@IgnoreModule IndexedDefaultMemberAccess, AssignmentNotUsed
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


'This method runs after every test in the module.
'@TestCleanup
Private Sub TestCleanup()
    Err.Clear
End Sub


'===================================================='
'===================== FIXTURES ====================='
'===================================================='


Private Function zfxDefDBM( _
            Optional ByVal FilePathName As String = vbNullString) As LiteACID
    If Len(FilePathName) > 0 Then
        Set zfxDefDBM = LiteACID(LiteADO(FilePathName))
    Else
        Set zfxDefDBM = LiteACID(LiteADO(FixObjAdo.DefaultDbPathName))
    End If
End Function


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("Integrity checking")
Private Sub ztcIntegrityADODB_PassesDefaultDatabaseIntegrityCheck()
    On Error GoTo TestFail

Arrange:
Act:
    Dim CheckResult As Boolean
    CheckResult = zfxDefDBM().IntegrityADODB
Assert:
    Assert.IsTrue CheckResult, "Integrity check on default database failed"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Integrity checking")
Private Sub ztcIntegrityADODB_ThrowsOnFileNotDatabase()
    On Error Resume Next
    zfxDefDBM(ThisWorkbook.Name).IntegrityADODB
    Guard.AssertExpectedError Assert, ErrNo.OLE_DB_ODBC_Err
End Sub


'@TestMethod("Integrity checking")
Private Sub ztcIntegrityADODB_ThrowsOnCorruptedDatabase()
    On Error Resume Next
    zfxDefDBM(FixObjAdo.RelPrefix & "ICfailFKCfail.db").IntegrityADODB
    Guard.AssertExpectedError Assert, ErrNo.IntegrityCheckErr
End Sub


'@TestMethod("Integrity checking")
Private Sub ztcIntegrityADODB_ThrowsOnFailedFKCheck()
    On Error Resume Next
    zfxDefDBM(FixObjAdo.RelPrefix & "ICokFKCfail.db").IntegrityADODB
    Guard.AssertExpectedError Assert, ErrNo.ConsistencyCheckErr
End Sub


''@EntryPoint
'Private Sub ztcTest()
'    Dim FilePathName As String
'    FilePathName = REL_PREFIX & LIB_NAME & ".db"
'
'    Dim dbm As ILiteADO
'    Set dbm = LiteADO(FilePathName)
'    Dim dbmCI As LiteADO
'    Set dbmCI = dbm
'
'    '@Ignore VariableNotUsed
'    Dim PathCheck As LiteFSCheck
'    Set PathCheck = LiteFSCheck(FilePathName, False)
'
'    Dim ACIDTool As LiteACID
'    Set ACIDTool = LiteACID(dbm)
'
'    Debug.Print dbm.GetScalar("PRAGMA busy_timeout")
'    dbm.ExecuteNonQuery "PRAGMA busy_timeout=1000"
'    dbmCI.AdoCommand.CommandTimeout = 1
'    Debug.Print dbm.GetScalar("PRAGMA busy_timeout")
'    Do While True
'        Debug.Print ACIDTool.LockedReadOnly
'        '@Ignore StopKeyword
'        Stop
'    Loop
'
'    Debug.Print ACIDTool.JournalModeToggle
'    Debug.Print ACIDTool.JournalModeToggle
'    Debug.Print ACIDTool.JournalModeToggle
'    dbmCI.AdoConnection.Close
'    Set dbm = Nothing
'End Sub
