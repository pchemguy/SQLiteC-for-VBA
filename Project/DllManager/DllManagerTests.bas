Attribute VB_Name = "DllManagerTests"
'@Folder "DllManager"
'@TestModule
'@IgnoreModule IndexedDefaultMemberAccess, LineLabelNotUsed
Option Explicit
Option Private Module

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
    Set DllMan = DllManager(DefaultPath)
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
    Set DllMan = DllManager(DefaultPath)
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
    Set DllMan = DllManager(DefaultPath)
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
    Set DllMan = DllManager("____INVALID PATH____")
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
    Set DllMan = DllManager(vbNullString)
    DllMan.DefaultPath = DefaultPath
Assert:
    Assert.AreEqual ThisWorkbook.Path & "\" & "Project", DllMan.DefaultPath, "Relative default path mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("DefaultPath")
Private Sub ztcDefaultPath_ThrowsOnInvalidPath()
    On Error Resume Next
    Dim DllMan As DllManager
    Set DllMan = DllManager(vbNullString)
    DllMan.DefaultPath = "____INVALID PATH____"
    Guard.AssertExpectedError Assert, ErrNo.FileNotFoundErr
End Sub

