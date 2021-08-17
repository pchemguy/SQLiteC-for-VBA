Attribute VB_Name = "ConstraintFKTests"
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
    Debug.Assert Not ConstraintFK("log_id", "logs", "id", , , "fk actions_log_id_logs_id") Is Nothing
    AssertExpectedError Assert, ErrNo.InvalidCharacterErr
End Sub


'@TestMethod("Input Validation")
Private Sub ztcCheckName_ThrowsIfNameWithQuote()
    On Error Resume Next
    Debug.Assert Not ConstraintFK("log_id", "logs", "id", , , "fk'actions_log_id_logs_id") Is Nothing
    AssertExpectedError Assert, ErrNo.InvalidCharacterErr
End Sub


'@TestMethod("Input Validation")
Private Sub ztcCheckName_ThrowsIfNameWithDash()
    On Error Resume Next
    Debug.Assert Not ConstraintFK("log_id", "logs", "id", , , "fk-actions_log_id_logs_id") Is Nothing
    AssertExpectedError Assert, ErrNo.InvalidCharacterErr
End Sub


'@TestMethod("Input Validation")
Private Sub ztcCheckName_ThrowsIfForeingTableNameWithSpace()
    On Error Resume Next
    Debug.Assert Not ConstraintFK("log_id", "log s", "id") Is Nothing
    AssertExpectedError Assert, ErrNo.InvalidCharacterErr
End Sub


'@TestMethod("Input Validation")
Private Sub ztcCheckName_ThrowsIfForeingTableNameWithQuote()
    On Error Resume Next
    Debug.Assert Not ConstraintFK("log_id", "log's", "id") Is Nothing
    AssertExpectedError Assert, ErrNo.InvalidCharacterErr
End Sub


'@TestMethod("Input Validation")
Private Sub ztcCheckName_ThrowsIfForeingTableNameWithDash()
    On Error Resume Next
    Debug.Assert Not ConstraintFK("log_id", "log-s", "id") Is Nothing
    AssertExpectedError Assert, ErrNo.InvalidCharacterErr
End Sub


'@TestMethod("Input Validation")
Private Sub ztcCheckFieldNames_ThrowsIfNameWithDash()
    On Error Resume Next
    Debug.Assert Not ConstraintFK("log_i-d", "logs", "id") Is Nothing
    AssertExpectedError Assert, ErrNo.InvalidCharacterErr
End Sub


'@TestMethod("Input Validation")
Private Sub ztcCheckFieldNames_ThrowsIfNameWithDashInArray()
    On Error Resume Next
    Debug.Assert Not ConstraintFK("log_id", "logs", "i-d") Is Nothing
    AssertExpectedError Assert, ErrNo.InvalidCharacterErr
End Sub


'@TestMethod("Input Validation")
Private Sub ztcCheckFieldNames_ThrowsIfNameNotStringOrArray()
    On Error Resume Next
    Debug.Assert Not ConstraintFK("log_id", "logs", 1) Is Nothing
    AssertExpectedError Assert, ErrNo.CustomErr
End Sub


'@TestMethod("Input Validation")
Private Sub ztcCheckFieldNames_ThrowsIfNameNotArrayOfStrings()
    On Error Resume Next
    Debug.Assert Not ConstraintFK(Array("name", 1), "logs", "id") Is Nothing
    AssertExpectedError Assert, ErrNo.TypeMismatchErr
End Sub


'@TestMethod("Input Validation")
Private Sub ztcCheckFieldNames_ThrowsIfInvalidOnDelete()
    On Error Resume Next
    Debug.Assert Not ConstraintFK("log_id", "logs", "id", "SET 5") Is Nothing
    AssertExpectedError Assert, ErrNo.ActionNotSupportedErr
End Sub


'@TestMethod("Input Validation")
Private Sub ztcCheckFieldNames_ThrowsIfInvalidOnUpdate()
    On Error Resume Next
    Debug.Assert Not ConstraintFK("log_id", "logs", "id", , "SET 5") Is Nothing
    AssertExpectedError Assert, ErrNo.ActionNotSupportedErr
End Sub


'@TestMethod("SQL")
Private Sub ztcSQL_ValidatesSingleFieldName()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = "    FOREIGN KEY(""log_id"") REFERENCES ""logs""(""id"")"
Act:
    Dim Actual As String
    Actual = ConstraintFK("log_id", "logs", "id").SQL
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
    Expected = "    CONSTRAINT ""fk_actions_log_id_logs_id"" FOREIGN KEY(""log_id"") REFERENCES ""logs""(""id"")"
Act:
    Dim Actual As String
    Actual = ConstraintFK("log_id", "logs", "id", , , "fk_actions_log_id_logs_id").SQL
Assert:
    Assert.AreEqual Expected, Actual, "Named constraint mismatch"

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
    Expected = "    FOREIGN KEY(""log_type"",""log_date"") REFERENCES ""logs""(""type"",""date"")"
Act:
    Dim Actual As String
    Actual = ConstraintFK(Array("log_type", "log_date"), "logs", Array("type", "date")).SQL
Assert:
    Assert.AreEqual Expected, Actual, "Two fields mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQL")
Private Sub ztcSQL_ValidatesOnDelete()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = "    FOREIGN KEY(""log_id"") REFERENCES ""logs""(""id"") ON DELETE NO ACTION"
Act:
    Dim Actual As String
    Actual = ConstraintFK("log_id", "logs", "id", "no action").SQL
Assert:
    Assert.AreEqual Expected, Actual, "OnDelete mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQL")
Private Sub ztcSQL_ValidatesOnUpdate()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = "    FOREIGN KEY(""log_id"") REFERENCES ""logs""(""id"") ON UPDATE SET NULL"
Act:
    Dim Actual As String
    Actual = ConstraintFK("log_id", "logs", "id", , "set null").SQL
Assert:
    Assert.AreEqual Expected, Actual, "OnUpdate mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQL")
Private Sub ztcSQL_ValidatesFull()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = "    CONSTRAINT ""fk_actions_log_id_logs_id"" FOREIGN KEY(""log_id"") " _
                 & "REFERENCES ""logs""(""id"") ON DELETE RESTRICT ON UPDATE CASCADE"
Act:
    Dim Actual As String
    Actual = ConstraintFK("log_id", "logs", "id", "restrict", "cascade", "fk_actions_log_id_logs_id").SQL
Assert:
    Assert.AreEqual Expected, Actual, "Full mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


