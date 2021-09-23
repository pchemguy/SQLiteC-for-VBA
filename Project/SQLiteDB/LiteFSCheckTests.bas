Attribute VB_Name = "LiteFSCheckTests"
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

Private FilePathName As String
Private PathCheck As LiteFSCheck
Private ErrNumber As Long
Private ErrSource As String
Private ErrDescription As String
Private ErrStack As String


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


'''' ##########################################################################
''''
'''' Must run "Library\SQLiteDBVBA\Fixtures\acl-restrict.bat" for access
''''   checking tests to work properly.
'''' Must run "Library\SQLiteDBVBA\Fixtures\acl-restore.bat" for git client to
''''   work properly.
''''
'''' ##########################################################################


Private Function zfxFixturePrefix() As String
    zfxFixturePrefix = ThisWorkbook.Path & PATH_SEP & REL_PREFIX & "Fixtures" & PATH_SEP
End Function


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("Integrity checking")
Private Sub ztcCreate_TraversesLockedFolder()
    On Error GoTo TestFail

Arrange:
    FilePathName = zfxFixturePrefix & "ACLLocked\LockedFolder\SubFolder\TestC.db"
Act:
    Set PathCheck = LiteFSCheck(FilePathName)
Assert:
    With PathCheck
        Assert.AreEqual FilePathName, .Database, "Database should be set"
        Assert.AreEqual 0, .ErrNumber, "ErrNumber should be 0"
        Assert.AreEqual 0, Len(.ErrSource), "ErrSource should be blank"
        Assert.AreEqual 0, Len(.ErrDescription), "ErrDescription should be blank"
        Assert.AreEqual 0, Len(.ErrStack), "ErrStack should be blank"
    End With

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'''' ##########################################################################
''''
'''' Must run "Library\SQLiteDBVBA\Fixtures\acl-restrict.bat" for access
''''   checking tests to work properly.
'''' Must run "Library\SQLiteDBVBA\Fixtures\acl-restore.bat" for git client to
''''   work properly.
''''
'''' ##########################################################################
'@TestMethod("Integrity checking")
Private Sub ztcCreate_FailsOnLastFolderACLLock()
    On Error GoTo TestFail

Arrange:
    FilePathName = ThisWorkbook.Path & PATH_SEP & REL_PREFIX & _
                   "Fixtures\ACLLocked\LockedFolder\LT100.db"
    FilePathName = zfxFixturePrefix & "ACLLocked\LockedDb.db" '''' FailsOnFileACLLock
    ErrNumber = ErrNo.PermissionDeniedErr
    ErrSource = "LiteFSCheck"
    ErrDescription = "Permission denied" & vbNewLine & _
                     "Access is denied to the database file. " & _
                     "Check ACL permissions and file locks." & _
                     vbNewLine & "Source: " & FilePathName
    ErrStack = "ExistsAccesibleValid" & vbNewLine & _
               "FileAccessibleValid" & vbNewLine
Act:
    Set PathCheck = LiteFSCheck(FilePathName)
Assert:
    With PathCheck
        Assert.AreEqual 0, Len(.Database), "Database should not be set"
        Assert.AreEqual ErrNumber, .ErrNumber, "ErrNumber mismatch"
        Assert.AreEqual ErrSource, .ErrSource, "ErrSource mismatch"
        Assert.AreEqual ErrDescription, .ErrDescription, "ErrDescription mismatch"
        Assert.AreEqual ErrStack, .ErrStack, "ErrStack mismatch"
    End With

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Integrity checking")
Private Sub ztcCreate_FailsOnIllegalPath()
    On Error GoTo TestFail

Arrange:
    FilePathName = ":Illegal Path<|>:"
    ErrNumber = ErrNo.PathNotFoundErr
    ErrSource = "LiteFSCheck"
    ErrDescription = "Database path (folder) is not found. Expected " & _
                     "absolute path. Check ACL settings. Enable path " & _
                     "resolution feature, if necessary." & _
                     vbNewLine & "Source: " & FilePathName
    ErrStack = "ExistsAccesibleValid" & vbNewLine & _
               "PathExistsAccessible" & vbNewLine
Act:
    Set PathCheck = LiteFSCheck(FilePathName)
Assert:
    With PathCheck
        Assert.AreEqual 0, Len(.Database), "Database should not be set"
        Assert.AreEqual ErrNumber, .ErrNumber, "ErrNumber mismatch"
        Assert.AreEqual ErrSource, .ErrSource, "ErrSource mismatch"
        Assert.AreEqual ErrDescription, .ErrDescription, "ErrDescription mismatch"
        Assert.AreEqual ErrStack, .ErrStack, "ErrStack mismatch"
    End With

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Integrity checking")
Private Sub ztcCreate_FailsOnNonExistentPath()
    On Error GoTo TestFail

Arrange:
    FilePathName = zfxFixturePrefix & "Dummy" & PATH_SEP & "Dummy.db"
    ErrNumber = ErrNo.PathNotFoundErr
    ErrSource = "LiteFSCheck"
    ErrDescription = "Database path (folder) is not found. Expected " & _
                     "absolute path. Check ACL settings. Enable path " & _
                     "resolution feature, if necessary." & _
                     vbNewLine & "Source: " & FilePathName
    ErrStack = "ExistsAccesibleValid" & vbNewLine & _
               "PathExistsAccessible" & vbNewLine
Act:
    Set PathCheck = LiteFSCheck(FilePathName)
Assert:
    With PathCheck
        Assert.AreEqual 0, Len(.Database), "Database should not be set"
        Assert.AreEqual ErrNumber, .ErrNumber, "ErrNumber mismatch"
        Assert.AreEqual ErrSource, .ErrSource, "ErrSource mismatch"
        Assert.AreEqual ErrDescription, .ErrDescription, "ErrDescription mismatch"
        Assert.AreEqual ErrStack, .ErrStack, "ErrStack mismatch"
    End With

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Integrity checking")
Private Sub ztcCreate_FailsOnNonExistentFile()
    On Error GoTo TestFail

Arrange:
    FilePathName = zfxFixturePrefix & "Dummy.db"
    ErrNumber = ErrNo.FileNotFoundErr
    ErrSource = "LiteFSCheck"
    ErrDescription = "Databse file is not found in the specified folder." & _
                     vbNewLine & "Source: " & FilePathName
    ErrStack = "ExistsAccesibleValid" & vbNewLine & _
               "PathExistsAccessible" & vbNewLine
Act:
    Set PathCheck = LiteFSCheck(FilePathName)
Assert:
    With PathCheck
        Assert.AreEqual 0, Len(.Database), "Database should not be set"
        Assert.AreEqual ErrNumber, .ErrNumber, "ErrNumber mismatch"
        Assert.AreEqual ErrSource, .ErrSource, "ErrSource mismatch"
        Assert.AreEqual ErrDescription, .ErrDescription, "ErrDescription mismatch"
        Assert.AreEqual ErrStack, .ErrStack, "ErrStack mismatch"
    End With

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Integrity checking")
Private Sub ztcCreate_FailsOnLT100File()
    On Error GoTo TestFail

Arrange:
    FilePathName = zfxFixturePrefix & "LT100.db"
    ErrNumber = ErrNo.OLE_DB_ODBC_Err
    ErrSource = "LiteFSCheck"
    ErrDescription = "File is not a database. SQLite header size is 100 bytes." & _
                     vbNewLine & "Source: " & FilePathName
    ErrStack = "ExistsAccesibleValid" & vbNewLine & _
               "FileAccessibleValid" & vbNewLine
Act:
    Set PathCheck = LiteFSCheck(FilePathName)
Assert:
    With PathCheck
        Assert.AreEqual 0, Len(.Database), "Database should not be set"
        Assert.AreEqual ErrNumber, .ErrNumber, "ErrNumber mismatch"
        Assert.AreEqual ErrSource, .ErrSource, "ErrSource mismatch"
        Assert.AreEqual ErrDescription, .ErrDescription, "ErrDescription mismatch"
        Assert.AreEqual ErrStack, .ErrStack, "ErrStack mismatch"
    End With

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Integrity checking")
Private Sub ztcCreate_FailsOnBadMagic()
    On Error GoTo TestFail

Arrange:
    FilePathName = zfxFixturePrefix & "BadMagic.db"
    ErrNumber = ErrNo.OLE_DB_ODBC_Err
    ErrSource = "LiteFSCheck"
    ErrDescription = "Database file is damaged. The magic string did not match." & _
                     vbNewLine & "Source: " & FilePathName
    ErrStack = "ExistsAccesibleValid" & vbNewLine & _
               "FileAccessibleValid" & vbNewLine
Act:
    Set PathCheck = LiteFSCheck(FilePathName)
Assert:
    With PathCheck
        Assert.AreEqual 0, Len(.Database), "Database should not be set"
        Assert.AreEqual ErrNumber, .ErrNumber, "ErrNumber mismatch"
        Assert.AreEqual ErrSource, .ErrSource, "ErrSource mismatch"
        Assert.AreEqual ErrDescription, .ErrDescription, "ErrDescription mismatch"
        Assert.AreEqual ErrStack, .ErrStack, "ErrStack mismatch"
    End With

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Integrity checking")
Private Sub ztcCreate_FailsOnReadLockedFile()
    On Error GoTo TestFail

Arrange:
    FilePathName = zfxFixturePrefix & "TestC.db"
    ErrNumber = ErrNo.TextStreamReadErr
    ErrSource = "LiteFSCheck"
    ErrDescription = "Method 'Read' of object 'ITextStream' failed" & vbNewLine & _
                     "Cannot read from the database file. " & _
                     "The file might be locked by another app." & _
                     vbNewLine & "Source: " & FilePathName & "-shm"
    ErrStack = "ExistsAccesibleValid" & vbNewLine & _
               "FileAccessibleValid" & vbNewLine
Act:
    Dim dbm As ILiteADO
    Set dbm = LiteADO(FilePathName)
    FilePathName = FilePathName & "-shm"
    dbm.ExecuteNonQuery "BEGIN IMMEDIATE"
    Set PathCheck = LiteFSCheck(FilePathName)
    dbm.ExecuteNonQuery "ROLLBACK"
Assert:
    With PathCheck
        Assert.AreEqual 0, Len(.Database), "Database should not be set"
        Assert.AreEqual ErrNumber, .ErrNumber, "ErrNumber mismatch"
        Assert.AreEqual ErrSource, .ErrSource, "ErrSource mismatch"
        Assert.AreEqual ErrDescription, .ErrDescription, "ErrDescription mismatch"
        Assert.AreEqual ErrStack, .ErrStack, "ErrStack mismatch"
    End With

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
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
