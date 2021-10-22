Attribute VB_Name = "DllManagerTests"
'@Folder "DllManager"
''@TestModule
'@IgnoreModule IndexedDefaultMemberAccess, UnhandledOnErrorResumeNext
'@IgnoreModule LineLabelNotUsed, VariableNotUsed, AssignmentNotUsed
Option Explicit
Option Private Module

Private Const LITE_LIB As String = "DllManager"
Private Const PATH_SEP As String = "\"
Private Const LITE_RPREFIX As String = "Library" & PATH_SEP & LITE_LIB & PATH_SEP

Public Const LoadingDllErr As Long = 48

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


Private Function zfxGetDefaultDllPath() As String
    #If WIN64 Then
        zfxGetDefaultDllPath = LITE_RPREFIX & "dll\x64"
    #Else
        zfxGetDefaultDllPath = LITE_RPREFIX & "dll\x32"
    #End If
End Function


Private Function zfxGetDefaultManager() As DllManager
    Dim DllPath As String
    DllPath = zfxGetDefaultDllPath
    Dim DllNames As Variant
    #If WIN64 Then
        DllNames = Array("sqlite3.dll", "libicudt68.dll", "libstdc++-6.dll", "libwinpthread-1.dll", "libicuuc68.dll", "libicuin68.dll")
    #Else
        DllNames = Array("icudt68.dll", "icuuc68.dll", "icuin68.dll", "icuio68.dll", "icutu68.dll", "sqlite3.dll")
    #End If
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

Arrange:
    Dim DefaultPath As String
    DefaultPath = vbNullString
Act:
    Dim DllMan As DllManager
    Set DllMan = DllManager.Create(DefaultPath)
Assert:
    Assert.AreEqual ThisWorkbook.Path, DllMan.DefaultPath, "Empty default path mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Factory")
Private Sub ztcCreate_VerifiesRelativePath()
    On Error GoTo TestFail

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
    Dim DllMan As DllManager
    Set DllMan = DllManager.Create("____INVALID PATH____")
    Guard.AssertExpectedError Assert, ErrNo.FileNotFoundErr
End Sub


'@TestMethod("DefaultPath")
Private Sub ztcDefaultPath_VerifiesRelativePath()
    On Error GoTo TestFail

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
    Dim DllMan As DllManager
    Set DllMan = DllManager.Create(vbNullString)
    DllMan.DefaultPath = "____INVALID PATH____"
    Guard.AssertExpectedError Assert, ErrNo.FileNotFoundErr
End Sub


'@TestMethod("Load")
Private Sub ztcLoad_ThrowsOnBitnessMismatch()
    On Error Resume Next
    '''' Set mismatched path to test for error
    Dim DllPath As String
    #If WIN64 Then
        DllPath = LITE_RPREFIX & "dll\x32"
    #Else
        DllPath = LITE_RPREFIX & "dll\x64"
    #End If
    Dim DllName As String
    DllName = "sqlite3.dll"
    
    Dim DllMan As DllManager
    Set DllMan = DllManager.Create(DllPath)
    Dim ResultCode As DllLoadStatus
    ResultCode = LOAD_ALREADY_LOADED
    ResultCode = DllMan.Load(DllName)
    Assert.AreEqual LOAD_ALREADY_LOADED, ResultCode, "Unexpected result code."
    Guard.AssertExpectedError Assert, LoadingDllErr
End Sub


'@TestMethod("Load")
Private Sub ztcLoad_VerifiesLoad()
    On Error GoTo TestFail

Arrange:
    Dim DllPath As String
    DllPath = zfxGetDefaultDllPath
    Dim DllNames As Variant
    #If WIN64 Then
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

Arrange:
    Dim DllPath As String
    DllPath = zfxGetDefaultDllPath
    Dim DllNames As Variant
    #If WIN64 Then
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

Arrange:
    Dim DllPath As String
    DllPath = zfxGetDefaultDllPath
    Dim DllMan As DllManager
    Set DllMan = DllManager.Create(DllPath)
Act:
    Dim ResultCode As DllLoadStatus
    #If WIN64 Then
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

Arrange:
    Dim DllICUName As String
    #If WIN64 Then
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

Arrange:
    Dim DllICUName As String
    #If WIN64 Then
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

