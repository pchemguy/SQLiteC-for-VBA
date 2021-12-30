Attribute VB_Name = "SQLUtilsTests"
'@Folder "SQLiteDBdev.Drafts.Helper - Working"
'@TestModule
'@IgnoreModule LineLabelNotUsed, UnhandledOnErrorResumeNext, FunctionReturnValueDiscarded
'@IgnoreModule IndexedDefaultMemberAccess
Option Explicit
Option Private Module

#Const LateBind = 0     '''' RubberDuck Tests
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
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("Helpers")
Private Sub ztcFieldsQA_ValidatesSingleField()
    On Error GoTo TestFail
    Assert.AreEqual "[A] AS [A]", SQLUtils.FieldsQA("A"), "Single field list mismatch"
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Helpers")
Private Sub ztcFieldsQA_ValidatesTwoFields()
    On Error GoTo TestFail
    Assert.AreEqual "[A] AS [A], [B] AS [B]", SQLUtils.FieldsQA("A", "B"), "Two-field list mismatch"
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Helpers")
Private Sub ztcFieldsQA_ThrowsOnEmptyArgumentList()
    On Error Resume Next
    SQLUtils.FieldsQA
    Guard.AssertExpectedError Assert, ErrNo.CustomErr
End Sub


'@TestMethod("Helpers")
Private Sub ztcFieldsQ_ValidatesSingleField()
    On Error GoTo TestFail
    Assert.AreEqual "[A]", SQLUtils.FieldsQ("A"), "Single field list mismatch"
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Helpers")
Private Sub ztcFieldsQ_ValidatesTwoFields()
    On Error GoTo TestFail
    Assert.AreEqual "[A], [B]", SQLUtils.FieldsQ("A", "B"), "Two-field list mismatch"
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Helpers")
Private Sub ztcFieldsQ_ThrowsOnEmptyArgumentList()
    On Error Resume Next
    SQLUtils.FieldsQ
    Guard.AssertExpectedError Assert, ErrNo.CustomErr
End Sub
