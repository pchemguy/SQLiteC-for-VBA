Attribute VB_Name = "LiteACIDTests"
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


Private Function zfxDefDBM( _
            Optional ByVal FilePathName As String = vbNullString) As LiteACID
    If Len(FilePathName) > 0 Then
        Set zfxDefDBM = LiteACID(LiteADO(FilePathName))
    Else
        Set zfxDefDBM = LiteACID(LiteADO(REL_PREFIX & LIB_NAME & ".db"))
    End If
End Function


Private Function zfxFixturePrefix() As String
    zfxFixturePrefix = ThisWorkbook.Path & PATH_SEP & REL_PREFIX & "Fixtures" & PATH_SEP
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
    zfxDefDBM(REL_PREFIX & "ICfailFKCfail.db").IntegrityADODB
    Guard.AssertExpectedError Assert, ErrNo.IntegrityCheckErr
End Sub


'@TestMethod("Integrity checking")
Private Sub ztcIntegrityADODB_ThrowsOnFailedFKCheck()
    On Error Resume Next
    zfxDefDBM(REL_PREFIX & "ICokFKCfail.db").IntegrityADODB
    Guard.AssertExpectedError Assert, ErrNo.ConsistencyCheckErr
End Sub


Private Sub ztcTest()
    Dim FilePathName As String
    FilePathName = REL_PREFIX & LIB_NAME & ".db"
    
    Dim dbm As ILiteADO
    Set dbm = LiteADO(FilePathName)
    
    Dim ACIDTool As LiteACID
    Set ACIDTool = LiteACID(dbm)
    
    
    Do While True
        Debug.Print ACIDTool.LockedReadOnly
        Stop
    Loop
    
    Debug.Print ACIDTool.JournalModeToggle
    Debug.Print ACIDTool.JournalModeToggle
    Debug.Print ACIDTool.JournalModeToggle

End Sub



'Private Sub ztcExistsAccesibleValid_ThrowsOnBadMagicA()
'    Dim FilePathName As String
'    FilePathName = zfxFixturePrefix & "TestCWAL.db"
'    Dim dbm As ILiteADO
'    Set dbm = LiteADO(FilePathName)
'
'    Dim AdoConnection As ADODB.Connection
'    Set AdoConnection = dbm.AdoConnection
'
'    LiteCheck(FilePathName).ExistsAccesibleValid
'    Dim Response As Variant
'    Response = dbm.GetScalar("PRAGMA journal_mode")
'    dbm.ExecuteNonQuery "PRAGMA journal_mode='DELETE'"
'    Response = dbm.GetScalar("PRAGMA journal_mode")
'
'    On Error Resume Next
'    FilePathName = zfxFixturePrefix & "TestCWAL.db"
'    Set dbm = LiteADO(FilePathName)
'    dbm.ExecuteNonQuery "BEGIN IMMEDIATE"
'    LiteCheck(FilePathName & "-shm").ExistsAccesibleValid
'    dbm.ExecuteNonQuery "ROLLBACK"
'    Guard.AssertExpectedError Assert, ErrNo.TextStreamReadErr
'End Sub


