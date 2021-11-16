Attribute VB_Name = "LiteFSCheckTests"
'@Folder "SQLite.Checks"
'@TestModule
'@IgnoreModule LineLabelNotUsed, UnhandledOnErrorResumeNext, FunctionReturnValueDiscarded
'@IgnoreModule IndexedDefaultMemberAccess
Option Explicit
Option Private Module

Private Const MODULE_NAME As String = "LiteFSCheckTests"
Private TestCounter As Long
Private Const PATH_SEP As String = "\"

Private FilePathName As String
Private PathCheck As LiteFSCheck
Private ErrNumber As Long
Private ErrSource As String
Private ErrDescription As String
Private ErrStack As String

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
    With Logger
        .ClearLog
        .DebugLevelDatabase = DEBUGLEVEL_MAX
        .DebugLevelImmediate = DEBUGLEVEL_NONE
        .UseIdPadding = True
        .UseTimeStamp = False
        .RecordIdDigits 3
        .TimerSet MODULE_NAME
    End With
    TestCounter = 0
End Sub


'This method runs once per module.
'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
    Logger.TimerLogClear MODULE_NAME, TestCounter
    Logger.PrintLog
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


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("Path checking")
Private Sub ztcCreate_TraversesLockedFolder()
    On Error GoTo TestFail

Arrange:
    FilePathName = FixObjAdo.FixPath("ACLLocked\LockedFolder\SubFolder\TestC.db")
Act:
    Set PathCheck = LiteFSCheck(FilePathName)
Assert:
    With PathCheck
        Assert.AreEqual FilePathName, .DatabasePathName, "Database should be set"
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
'@TestMethod("Path checking")
Private Sub ztcCreate_FailsOnLastFolderACLLock()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    FilePathName = FixObjAdo.FixPath("ACLLocked\LockedFolder\LT100.db")
    FilePathName = FixObjAdo.FixPath("ACLLocked\LockedDb.db") '''' FailsOnFileACLLock
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
        If Len(.DatabasePathName) > 0 Then Assert.Inconclusive "Database should not be set"
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


'@TestMethod("Path checking")
Private Sub ztcCreate_FailsOnIllegalPath()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

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
        Assert.AreEqual 0, Len(.DatabasePathName), "Database should not be set"
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


'@TestMethod("Path checking")
Private Sub ztcCreate_FailsOnNonExistentPath()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    FilePathName = FixObjAdo.FixPath("Dummy" & PATH_SEP & "Dummy.db")
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
        Assert.AreEqual 0, Len(.DatabasePathName), "Database should not be set"
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


'@TestMethod("Path checking")
Private Sub ztcCreate_FailsOnNonExistentFile()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    FilePathName = FixObjAdo.FixPath("Dummy.db")
    ErrNumber = ErrNo.FileNotFoundErr
    ErrSource = "LiteFSCheck"
    ErrDescription = "Databse file is not found in the specified folder." & _
                     vbNewLine & "Source: " & FilePathName
    ErrStack = "ExistsAccesibleValid" & vbNewLine
Act:
    Set PathCheck = LiteFSCheck(FilePathName)
Assert:
    With PathCheck
        Assert.AreEqual 0, Len(.DatabasePathName), "Database should not be set"
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


'@TestMethod("Path checking")
Private Sub ztcCreate_FailsOnLT100File()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    FilePathName = FixObjAdo.FixPath("LT100.db")
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
        Assert.AreEqual 0, Len(.DatabasePathName), "Database should not be set"
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


'@TestMethod("Path checking")
Private Sub ztcCreate_FailsOnBadMagic()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    FilePathName = FixObjAdo.FixPath("BadMagic.db")
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
        Assert.AreEqual 0, Len(.DatabasePathName), "Database should not be set"
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


'@TestMethod("Path checking")
Private Sub ztcCreate_FailsOnReadLockedFile()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    FilePathName = FixObjAdo.FixPath("TestC.db")
    ErrNumber = ErrNo.TextStreamReadErr
    ErrSource = "LiteFSCheck"
    ErrDescription = "Method 'Read' of object 'ITextStream' failed" & vbNewLine & _
                     "Cannot read from the database file. " & _
                     "The file might be locked by another app." & _
                     vbNewLine & "Source: " & FilePathName & "-shm"
    ErrStack = "ExistsAccesibleValid" & vbNewLine & _
               "FileAccessibleValid" & vbNewLine
Act:
    Dim dbq As ILiteADO
    Set dbq = LiteADO(FilePathName)
    FilePathName = FilePathName & "-shm"
    dbq.ExecuteNonQuery "BEGIN IMMEDIATE"
    Set PathCheck = LiteFSCheck(FilePathName)
    dbq.ExecuteNonQuery "ROLLBACK"
Assert:
    With PathCheck
        Assert.AreEqual 0, Len(.DatabasePathName), "Database should not be set"
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


'@TestMethod("Path checking")
Private Sub ztcCreate_FailsOnEmptyPath()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    FilePathName = vbNullString
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
        Assert.AreEqual 0, Len(.DatabasePathName), "Database should not be set"
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


'@TestMethod("Path resolution")
Private Sub ztcCreate_ResolvesRelativePath()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    FilePathName = FixObjAdo.DefaultDbPathNameRel
    Dim Expected As String
    Expected = ThisWorkbook.Path & PATH_SEP & FilePathName
Act:
    Set PathCheck = LiteFSCheck(FilePathName, False)
Assert:
    Assert.AreEqual 0, PathCheck.ErrNumber, "Unexpected error occured"
    Assert.AreEqual Expected, PathCheck.DatabasePathName, "Resolved path mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Path resolution")
Private Sub ztcCreate_FailsResolveCreatableWithEmptyPath()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    FilePathName = vbNullString
    ErrNumber = ErrNo.FileNotFoundErr
    ErrSource = "CommonRoutines"
    ErrDescription = "File <> not found!" & vbNewLine & _
                     "Source: " & FilePathName
    ErrStack = "ExistsAccesibleValid" & vbNewLine
Act:
    Set PathCheck = LiteFSCheck(FilePathName, True)
Assert:
    With PathCheck
        Assert.AreEqual 0, Len(.DatabasePathName), "Database should not be set"
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


'@TestMethod("New database")
Private Sub ztcCreate_FailsCreateDbInReadOnlyDir()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    FilePathName = Environ$("ALLUSERSPROFILE") & PATH_SEP & "Dummy.db"
    ErrNumber = ErrNo.PermissionDeniedErr
    ErrSource = "LiteFSCheck"
    ErrDescription = "Permission denied" & vbNewLine & _
                     "Cannot create a new file." & vbNewLine & _
                     "Source: " & FilePathName
    ErrStack = "ExistsAccesibleValid" & vbNewLine
Act:
    Set PathCheck = LiteFSCheck(FilePathName, Null)
Assert:
    With PathCheck
        Assert.AreEqual 0, Len(.DatabasePathName), "Database should not be set"
        Assert.AreEqual ErrNumber, .ErrNumber, "ErrNumber mismatch"
        Assert.IsTrue InStr(.ErrSource, ErrSource) > 0, "ErrSource mismatch"
        Assert.AreEqual ErrDescription, .ErrDescription, "ErrDescription mismatch"
        Assert.AreEqual ErrStack, .ErrStack, "ErrStack mismatch"
    End With

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("New database")
Private Sub ztcCreate_FailsCreateDbNoFileName()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    FilePathName = FixObjAdo.FixPath
    ErrNumber = ErrNo.FileNotFoundErr
    ErrSource = "LiteFSCheck"
    ErrDescription = "Filename is not provided or provided name conflicts " & _
                     "with existing folder." & vbNewLine & _
                     "Source: " & FilePathName
    ErrStack = "ExistsAccesibleValid" & vbNewLine
Act:
    Set PathCheck = LiteFSCheck(FilePathName, Null)
Assert:
    With PathCheck
        Assert.AreEqual 0, Len(.DatabasePathName), "Database should not be set"
        Assert.AreEqual ErrNumber, .ErrNumber, "ErrNumber mismatch"
        Assert.IsTrue InStr(.ErrSource, ErrSource) > 0, "ErrSource mismatch"
        Assert.AreEqual ErrDescription, .ErrDescription, "ErrDescription mismatch"
        Assert.AreEqual ErrStack, .ErrStack, "ErrStack mismatch"
    End With

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Path resolution")
Private Sub ztcCreate_FailsResolvingBlankNoCreatePath()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1
    
Arrange:
    FilePathName = vbNullString
    Dim Expected As String
    Expected = FixObjAdo.DefaultDbPathName
Act:
    Set PathCheck = LiteFSCheck(FilePathName, False)
Assert:
    Assert.AreEqual ErrNo.FileNotFoundErr, PathCheck.ErrNumber, "Unexpected error occured"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Path resolution")
Private Sub ztcCreate_ResolvesNameOnlyNoCreatePath()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1
    
Arrange:
    FilePathName = FixObjAdo.DefaultDbName
    Dim Expected As String
    Expected = FixObjAdo.DefaultDbPathName
Act:
    Set PathCheck = LiteFSCheck(FilePathName, False)
Assert:
    Assert.AreEqual 0, PathCheck.ErrNumber, "Unexpected error occured"
    Assert.AreEqual Expected, PathCheck.DatabasePathName, "Resolved path mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Path resolution")
Private Sub ztcCreate_ResolvesInMemory()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    FilePathName = ":MEmoRy:"
    Dim Expected As String
    Expected = LCase$(FilePathName)
Act:
    Set PathCheck = LiteFSCheck(FilePathName, True)
Assert:
    Assert.AreEqual 0, PathCheck.ErrNumber, "Unexpected error occured"
    Assert.AreEqual Expected, PathCheck.DatabasePathName, "Resolved path mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Path resolution")
Private Sub ztcCreate_ResolvesAnonPath()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    FilePathName = ":blank:"
    Dim Expected As String
    Expected = vbNullString
Act:
    Set PathCheck = LiteFSCheck(FilePathName, True)
Assert:
    Assert.AreEqual 0, PathCheck.ErrNumber, "Unexpected error occured"
    Assert.AreEqual Expected, PathCheck.DatabasePathName, "Resolved path mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Path resolution")
Private Sub ztcCreate_ResolvesTemp()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    FilePathName = ":temp:"
    Dim Actual As String
    Dim Prefix As String
    Dim SuffixPattern As String
Act:
    Set PathCheck = LiteFSCheck(FilePathName, True)
    Prefix = Environ$("TEMP") & PATH_SEP & Format$(Now, "yyyy_mm_dd-hh_mm_")
    SuffixPattern = "ss-12345678.db"
    Actual = PathCheck.DatabasePathName
Assert:
    Assert.AreEqual 0, PathCheck.ErrNumber, "Unexpected error occured"
    Assert.AreEqual Prefix, Left$(Actual, Len(Prefix)), "Resolved path mismatch"
    Assert.AreEqual ".db", Right$(Actual, 3), "Resolved path mismatch"
    Assert.AreEqual Len(Prefix) + Len(SuffixPattern), Len(Actual), "Resolved path mismatch"
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
