Attribute VB_Name = "ConstraintPKTests"
'@Folder "SQLiteDBdev.DB Objects.Table Constraint"
'@TestModule
'@IgnoreModule LineLabelNotUsed, IndexedDefaultMemberAccess
'@IgnoreModule UnhandledOnErrorResumeNext, VariableNotUsed, AssignmentNotUsed
Option Explicit
Option Private Module

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
    Debug.Assert Not ConstraintPK(Array("id", "name"), "pk log", True) Is Nothing
    Guard.AssertExpectedError Assert, ErrNo.InvalidCharacterErr
End Sub


'@TestMethod("Input Validation")
Private Sub ztcCheckName_ThrowsIfNameWithQuote()
    On Error Resume Next
    Debug.Assert Not ConstraintPK(Array("id", "name"), "pk'log", True) Is Nothing
    Guard.AssertExpectedError Assert, ErrNo.InvalidCharacterErr
End Sub


'@TestMethod("Input Validation")
Private Sub ztcCheckName_ThrowsIfNameWithDash()
    On Error Resume Next
    Debug.Assert Not ConstraintPK(Array("id", "name"), "pk-log", True) Is Nothing
    Guard.AssertExpectedError Assert, ErrNo.InvalidCharacterErr
End Sub


'@TestMethod("Input Validation")
Private Sub ztcCheckFieldNames_ThrowsIfNameWithDash()
    On Error Resume Next
    Debug.Assert Not ConstraintPK("i-d", "pk_log", True) Is Nothing
    Guard.AssertExpectedError Assert, ErrNo.InvalidCharacterErr
End Sub


'@TestMethod("Input Validation")
Private Sub ztcCheckFieldNames_ThrowsIfNameWithDashInArray()
    On Error Resume Next
    Debug.Assert Not ConstraintPK(Array("i-d"), "pk_log", True) Is Nothing
    Guard.AssertExpectedError Assert, ErrNo.InvalidCharacterErr
End Sub


'@TestMethod("Input Validation")
Private Sub ztcCheckFieldNames_ThrowsIfNameNotStringOrArray()
    On Error Resume Next
    Debug.Assert Not ConstraintPK(1, "pk_log", True) Is Nothing
    Guard.AssertExpectedError Assert, ErrNo.CustomErr
End Sub


'@TestMethod("Input Validation")
Private Sub ztcCheckFieldNames_ThrowsIfNameNotArrayOfStrings()
    On Error Resume Next
    Debug.Assert Not ConstraintPK(Array("name", 1), "pk_log", True) Is Nothing
    Guard.AssertExpectedError Assert, ErrNo.TypeMismatchErr
End Sub


'@TestMethod("SQL")
Private Sub ztcSQL_ValidatesSingleFieldName()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = "    PRIMARY KEY(""id"")"
Act:
    Dim Actual As String
    Actual = ConstraintPK("id").SQL
Assert:
    Assert.AreEqual Expected, Actual, "Single field mismatch"

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
    Expected = "    CONSTRAINT ""pk_log"" PRIMARY KEY(""id"")"
Act:
    Dim Actual As String
    Actual = ConstraintPK("id", "pk_log").SQL
Assert:
    Assert.AreEqual Expected, Actual, "Named constraint mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQL")
Private Sub ztcSQL_ValidatesAutoincrement()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = "    CONSTRAINT ""pk_log"" PRIMARY KEY(""id"" AUTOINCREMENT)"
Act:
    Dim Actual As String
    Actual = ConstraintPK("id", "pk_log", True).SQL
Assert:
    Assert.AreEqual Expected, Actual, "Autoincrement mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQL")
Private Sub ztcSQL_ValidatesTwoFields()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = "    CONSTRAINT ""pk_log"" PRIMARY KEY(""user"",""email"")"
Act:
    Dim Actual As String
    Actual = ConstraintPK(Array("user", "email"), "pk_log").SQL
Assert:
    Assert.AreEqual Expected, Actual, "Two fields mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
