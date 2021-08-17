Attribute VB_Name = "ConstraintCKTests"
'@Folder "SQLiteDB.DB Objects.Table Constraint"
'@TestModule
'@IgnoreModule LineLabelNotUsed, IndexedDefaultMemberAccess
'@IgnoreModule UnhandledOnErrorResumeNext, VariableNotUsed, AssignmentNotUsed
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


'This method runs after every test in the module.
'@TestCleanup
Private Sub TestCleanup()
    Err.Clear
End Sub


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("Input Validation")
Private Sub ztcCheckName_ThrowsIfNameWithSpace()
    On Error Resume Next
    Debug.Assert Not ConstraintCK("id > 5", "ck id") Is Nothing
    AssertExpectedError Assert, ErrNo.InvalidCharacterErr
End Sub


'@TestMethod("Input Validation")
Private Sub ztcCheckName_ThrowsIfNameWithQuote()
    On Error Resume Next
    Debug.Assert Not ConstraintCK("id > 5", "ck'id") Is Nothing
    AssertExpectedError Assert, ErrNo.InvalidCharacterErr
End Sub


'@TestMethod("Input Validation")
Private Sub ztcCheckName_ThrowsIfNameWithDash()
    On Error Resume Next
    Debug.Assert Not ConstraintCK("id > 5", "ck-id") Is Nothing
    AssertExpectedError Assert, ErrNo.InvalidCharacterErr
End Sub


'@TestMethod("SQL")
Private Sub ztcSQL_ValidatesSimpleConstraint()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = "    CHECK(id > 5)"
Act:
    Dim Actual As String
    Actual = ConstraintCK("id > 5").SQL
Assert:
    Assert.AreEqual Expected, Actual, "Named constraint mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQL")
Private Sub ztcSQL_ValidatesNamedConstraint()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = "    CONSTRAINT ""ck_id"" CHECK(id > 5)"
Act:
    Dim Actual As String
    Actual = ConstraintCK("id > 5", "ck_id").SQL
Assert:
    Assert.AreEqual Expected, Actual, "Named constraint mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
