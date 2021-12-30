Attribute VB_Name = "DllManagerTests"
'@Folder "DllTools.Manager"
'@TestModule
'@IgnoreModule IndexedDefaultMemberAccess, UnhandledOnErrorResumeNext
'@IgnoreModule LineLabelNotUsed, VariableNotUsed, AssignmentNotUsed
Option Explicit
Option Private Module

Private Const MODULE_NAME As String = "DllManagerTests"
Private TestCounter As Long

Private Const LIB_NAME As String = "DllTools"
Private Const PATH_SEP As String = "\"
Private Const LIB_RPREFIX As String = _
    "Library" & PATH_SEP & LIB_NAME & PATH_SEP & "dll" & PATH_SEP

#Const LateBind = 1     '''' RubberDuck Tests
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


Private Function zfxGetDefaultAppDllPath() As String
    zfxGetDefaultAppDllPath = ThisWorkbook.Path & "\Library\" & _
                              ThisWorkbook.VBProject.Name & "\dll\" & CStr(ARCH)
End Function


Private Function zfxGetLibraryArray() As Variant
    #If Win64 Then
        zfxGetLibraryArray = Array("sqlite3.dll", "libicudt68.dll", _
                                   "libstdc++-6.dll", "libwinpthread-1.dll", _
                                   "libicuuc68.dll", "libicuin68.dll")
    #Else
        zfxGetLibraryArray = Array("icudt68.dll", "icuuc68.dll", "icuin68.dll", _
                                   "icuio68.dll", "icutu68.dll", "sqlite3.dll")
    #End If
End Function


Private Function zfxGetDefaultManager() As DllManager
    Dim DllPath As String
    DllPath = LIB_RPREFIX & CStr(ARCH)
    Dim DllNames As Variant
    DllNames = zfxGetLibraryArray()
    Dim DllMan As DllManager
    Set DllMan = DllManager.Create(DllPath, DllNames)
    If DllMan Is Nothing Then Err.Raise ErrNo.UnknownClassErr, _
        "DllManagerTests", "Failed to create a DllManager instance."
    Set zfxGetDefaultManager = DllMan
End Function


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("Factory")
Private Sub ztcCreate_VerifiesEmptyPath()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim DefaultPath As String
    DefaultPath = vbNullString
Act:
    Dim DllMan As DllManager
    Set DllMan = DllManager.Create(DefaultPath)
    Dim Expected As String
    Expected = zfxGetDefaultAppDllPath()
    Dim Actual As String
    Actual = DllMan.DefaultPath
Assert:
    Assert.AreEqual Expected, Actual, "Empty default path mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Factory")
Private Sub ztcCreate_VerifiesRelativePath()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim DefaultPath As String
    DefaultPath = "Project"
Act:
    Dim DllMan As DllManager
    Set DllMan = DllManager.Create(DefaultPath)
Assert:
    Assert.AreEqual ThisWorkbook.Path & "\" & "Project", DllMan.DefaultPath, "Relative default path mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Factory")
Private Sub ztcCreate_VerifiesAbsolutePath()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim DefaultPath As String
    DefaultPath = ThisWorkbook.Path & "\" & "Library"
Act:
    Dim DllMan As DllManager
    Set DllMan = DllManager.Create(DefaultPath)
Assert:
    Assert.AreEqual DefaultPath, DllMan.DefaultPath, "Absolute default path mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Factory")
Private Sub ztcCreate_ThrowsOnInvalidPath()
    On Error Resume Next
    TestCounter = TestCounter + 1
    Dim DllMan As DllManager
    Set DllMan = DllManager.Create("____INVALID PATH____")
    Guard.AssertExpectedError Assert, ErrNo.FileNotFoundErr
End Sub


'@TestMethod("DefaultPath")
Private Sub ztcDefaultPath_VerifiesRelativePath()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim DefaultPath As String
    DefaultPath = "Project"
Act:
    Dim DllMan As DllManager
    Set DllMan = DllManager.Create(vbNullString)
    DllMan.DefaultPath = DefaultPath
Assert:
    Assert.AreEqual ThisWorkbook.Path & "\" & "Project", DllMan.DefaultPath, "Relative default path mismatch"
    Assert.AreEqual 0, DllMan.Dlls.Count, "Dlls.Count mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DefaultPath")
Private Sub ztcDefaultPath_ThrowsOnInvalidPath()
    On Error Resume Next
    TestCounter = TestCounter + 1
    Dim DllMan As DllManager
    Set DllMan = DllManager.Create(vbNullString)
    DllMan.DefaultPath = "____INVALID PATH____"
    Guard.AssertExpectedError Assert, ErrNo.FileNotFoundErr
End Sub


'@TestMethod("Load")
Private Sub ztcLoad_ThrowsOnBitnessMismatch()
    On Error Resume Next
    TestCounter = TestCounter + 1
    Err.Clear
    '''' Set mismatched path to test for error
    Dim DllPath As String
    If ARCH = "x32" Then
        DllPath = LIB_RPREFIX & "x64"
    Else
        DllPath = LIB_RPREFIX & "x32"
    End If
    Dim DllName As String
    DllName = "sqlite3.dll"
    
    Dim DllMan As DllManager
    Set DllMan = DllManager.Create(DllPath)
    Dim ResultCode As DllLoadStatus
    ResultCode = DllMan.Load(DllName)
    Guard.AssertExpectedError Assert, LoadingDllErr
End Sub


'@TestMethod("Load")
Private Sub ztcLoad_VerifiesLoad()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim DllPath As String
    DllPath = LIB_RPREFIX & CStr(ARCH)
    Dim DllNames As Variant
    #If Win64 Then
        DllNames = "sqlite3.dll"
    #Else
        DllNames = "icudt68.dll"
    #End If
Act:
    Dim DllMan As DllManager
    Set DllMan = DllManager.Create(DllPath)
    Dim ResultCode As DllLoadStatus
    ResultCode = DllMan.Load(DllNames)
Assert:
    Assert.AreEqual LOAD_OK, ResultCode, "Unexpected loading result code."
    Assert.AreEqual TextCompare, DllMan.Dlls.CompareMode, "CompareMode mismatch"
    Assert.AreEqual 1, DllMan.Dlls.Count, "Dlls.Count mismatch"
    Assert.IsTrue DllMan.Dlls.Exists(DllNames), "Dll is not in DllMan"
    
    ResultCode = DllMan.Load(DllNames)
    Assert.AreEqual LOAD_ALREADY_LOADED, ResultCode, "Unexpected loading result code."
    Assert.AreEqual 1, DllMan.Dlls.Count, "Dlls.Count mismatch"
    Assert.IsTrue DllMan.Dlls.Exists(DllNames), "Dll is not in DllMan"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Load")
Private Sub ztcLoadMultiple_VerifiesLoadOne()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim DllPath As String
    DllPath = LIB_RPREFIX & CStr(ARCH)
    Dim DllNames As Variant
    #If Win64 Then
        DllNames = "sqlite3.dll"
    #Else
        DllNames = "icudt68.dll"
    #End If
    Dim DllMan As DllManager
    Set DllMan = DllManager.Create(DllPath)
Act:
    Dim ResultCode As DllLoadStatus
    ResultCode = DllMan.LoadMultiple(DllNames)
Assert:
    Assert.AreEqual LOAD_OK, ResultCode, "Unexpected loading result code."
    Assert.AreEqual TextCompare, DllMan.Dlls.CompareMode, "CompareMode mismatch"
    Assert.AreEqual 1, DllMan.Dlls.Count, "Dlls.Count mismatch"
    Assert.IsTrue DllMan.Dlls.Exists(DllNames), "Dll is not in DllMan"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Load")
Private Sub ztcLoadMultiple_VerifiesLoadArray()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
Act:
    Dim DllMan As DllManager
    Set DllMan = zfxGetDefaultManager
Assert:
    Assert.AreEqual TextCompare, DllMan.Dlls.CompareMode, "CompareMode mismatch"
    Assert.AreEqual 6, DllMan.Dlls.Count, "Dlls.Count mismatch"
    Assert.IsTrue DllMan.Dlls.Exists("sqlite3.dll"), "sqlite3.dll is not in DllMan"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Load")
Private Sub ztcLoadMultiple_VerifiesLoadParamArray()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim DllPath As String
    DllPath = LIB_RPREFIX & CStr(ARCH)
    Dim DllMan As DllManager
    Set DllMan = DllManager.Create(DllPath)
Act:
    Dim ResultCode As DllLoadStatus
    #If Win64 Then
        ResultCode = DllMan.LoadMultiple("sqlite3.dll", "libicudt68.dll", "libstdc++-6.dll", "libwinpthread-1.dll", "libicuuc68.dll", "libicuin68.dll")
    #Else
        ResultCode = DllMan.LoadMultiple("icudt68.dll", "icuuc68.dll", "icuin68.dll", "icuio68.dll", "icutu68.dll", "sqlite3.dll")
    #End If
    
Assert:
    Assert.AreEqual LOAD_OK, ResultCode, "Unexpected loading result code."
    Assert.AreEqual TextCompare, DllMan.Dlls.CompareMode, "CompareMode mismatch"
    Assert.AreEqual 6, DllMan.Dlls.Count, "Dlls.Count mismatch"
    Assert.IsTrue DllMan.Dlls.Exists("sqlite3.dll"), "sqlite3.dll is not in DllMan"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Free")
Private Sub ztcFree_VerifiesFree()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim DllMan As DllManager
    Set DllMan = zfxGetDefaultManager
Act:
    Dim ResultCode As DllLoadStatus
    ResultCode = DllMan.Free("sqlite3.dll")
Assert:
    Assert.AreEqual LOAD_OK, ResultCode, "Unexpected loading result code."
    Assert.AreEqual 5, DllMan.Dlls.Count, "Dlls.Count mismatch"
    Assert.IsFalse DllMan.Dlls.Exists("sqlite3.dll"), "sqlite3.dll should not be in DllMan"

    ResultCode = DllMan.Free("sqlite3.dll")
    Assert.AreEqual LOAD_OK, ResultCode, "Unexpected loading result code."
    Assert.AreEqual 5, DllMan.Dlls.Count, "Dlls.Count mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Free")
Private Sub ztcFreeMultiple_VerifiesFreeOne()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim DllMan As DllManager
    Set DllMan = zfxGetDefaultManager
Act:
    Dim ResultCode As DllLoadStatus
    ResultCode = DllMan.FreeMultiple("sqlite3.dll")
Assert:
    Assert.AreEqual LOAD_OK, ResultCode, "Unexpected loading result code."
    Assert.AreEqual 5, DllMan.Dlls.Count, "Dlls.Count mismatch"
    Assert.IsFalse DllMan.Dlls.Exists("sqlite3.dll"), "sqlite3.dll should not be in DllMan"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Free")
Private Sub ztcFreeMultiple_VerifiesFreeTwoParamArray()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim DllICUName As String
    #If Win64 Then
        DllICUName = "libicudt68.dll"
    #Else
        DllICUName = "icudt68.dll"
    #End If
    Dim DllMan As DllManager
    Set DllMan = zfxGetDefaultManager
Act:
    Dim ResultCode As DllLoadStatus
    ResultCode = DllMan.FreeMultiple("sqlite3.dll", DllICUName)
Assert:
    Assert.AreEqual LOAD_OK, ResultCode, "Unexpected loading result code."
    Assert.AreEqual 4, DllMan.Dlls.Count, "Dlls.Count mismatch"
    Assert.IsFalse DllMan.Dlls.Exists("sqlite3.dll"), "sqlite3.dll should not be in DllMan"
    Assert.IsFalse DllMan.Dlls.Exists(DllICUName), DllICUName & " should not be in DllMan"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Free")
Private Sub ztcFreeMultiple_VerifiesFreeTwoArray()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim DllICUName As String
    #If Win64 Then
        DllICUName = "libicudt68.dll"
    #Else
        DllICUName = "icudt68.dll"
    #End If
    Dim DllMan As DllManager
    Set DllMan = zfxGetDefaultManager
Act:
    Dim ResultCode As DllLoadStatus
    ResultCode = DllMan.FreeMultiple(Array("sqlite3.dll", DllICUName))
Assert:
    Assert.AreEqual LOAD_OK, ResultCode, "Unexpected loading result code."
    Assert.AreEqual 4, DllMan.Dlls.Count, "Dlls.Count mismatch"
    Assert.IsFalse DllMan.Dlls.Exists("sqlite3.dll"), "sqlite3.dll should not be in DllMan"
    Assert.IsFalse DllMan.Dlls.Exists(DllICUName), DllICUName & " should not be in DllMan"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Free")
Private Sub ztcFreeMultiple_VerifiesFreeAll()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim DllMan As DllManager
    Set DllMan = zfxGetDefaultManager
Act:
    Dim ResultCode As DllLoadStatus
    ResultCode = DllMan.FreeMultiple
Assert:
    Assert.AreEqual LOAD_OK, ResultCode, "Unexpected loading result code."
    Assert.AreEqual 0, DllMan.Dlls.Count, "Dlls.Count mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ProcAddress")
Private Sub ztcProcAddressGet_VerifiesProcAddress()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim DllMan As DllManager
    Set DllMan = zfxGetDefaultManager
    Dim LoadResult As DllLoadStatus
    LoadResult = DllMan.Load("kernel32", , False)
Assert:
    Assert.AreNotEqual 0, DllMan.ProcAddressGet("kernel32", "GetProcAddress"), "Failed to get an address."

CleanExit:
    Exit Sub
TestFail:
    If Not Assert Is Nothing Then
        Assert.Fail "Error: " & Err.Number & " - " & Err.Description
    Else
        Debug.Print "Assert is Nothing. ## Error: " & Err.Number & " - " & Err.Description
    End If
End Sub


'@TestMethod("IndirectCall")
Private Sub ztcIndirectCall_VerifiesFunc0ArgsReturnLong()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim DllMan As DllManager
    Set DllMan = zfxGetDefaultManager
Act:
    Dim Result As Long
    Result = DllMan.IndirectCall("SQLite3", "sqlite3_libversion_number", CC_STDCALL, vbLong, Empty)
Assert:
    Assert.IsTrue Result > 3 * 10 ^ 6, "Failed to call dll function/args-0/ret-Long."

CleanExit:
    Exit Sub
TestFail:
    If Not Assert Is Nothing Then
        Assert.Fail "Error: " & Err.Number & " - " & Err.Description
    Else
        Debug.Print "Assert is Nothing. ## Error: " & Err.Number & " - " & Err.Description
    End If
End Sub
