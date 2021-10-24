Attribute VB_Name = "IndexOTests"
'@Folder "SQLite.SQLiteDBdev.DB Objects"
'@TestModule
'@IgnoreModule LineLabelNotUsed, IndexedDefaultMemberAccess
'@IgnoreModule UnhandledOnErrorResumeNext, VariableNotUsed, AssignmentNotUsed
Option Explicit
Option Private Module

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
    Debug.Assert Not IndexO("idx contacts_email", "contacts", "email") Is Nothing
    Guard.AssertExpectedError Assert, ErrNo.InvalidCharacterErr
End Sub


'@TestMethod("Input Validation")
Private Sub ztcCheckName_ThrowsIfNameWithQuote()
    On Error Resume Next
    Debug.Assert Not IndexO("idx'contacts_email", "contacts", "email") Is Nothing
    Guard.AssertExpectedError Assert, ErrNo.InvalidCharacterErr
End Sub


'@TestMethod("Input Validation")
Private Sub ztcCheckName_ThrowsIfTableNameWithDash()
    On Error Resume Next
    Debug.Assert Not IndexO("idx_contacts_email", "contacts-", "email") Is Nothing
    Guard.AssertExpectedError Assert, ErrNo.InvalidCharacterErr
End Sub


'@TestMethod("SQL")
Private Sub ztcSQL_ValidatesSingleFieldName()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = "CREATE INDEX ""idx_contacts_email"" ON ""contacts""(""email"")"
Act:
    Dim Actual As String
    Actual = IndexO("idx_contacts_email", "contacts", "email").SQL
Assert:
    Assert.AreEqual Expected, Actual, "Single field mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQL")
Private Sub ztcSQL_ValidatesSingleFieldUnique()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = "CREATE UNIQUE INDEX ""idx_contacts_email"" ON ""contacts""(""email"")"
Act:
    Dim Actual As String
    Actual = IndexO("idx_contacts_email", "contacts", "email", True).SQL
Assert:
    Assert.AreEqual Expected, Actual, "Single field_unique mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQL")
Private Sub ztcSQL_ValidatesSingleFieldCollateOrder()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = "CREATE INDEX ""idx_contacts_email"" ON ""contacts""(""email"" COLLATE NOCASE ASC)"
Act:
    Dim Actual As String
    Actual = IndexO("idx_contacts_email", "contacts", Array("email", "ASC", "NOCASE")).SQL
Assert:
    Assert.AreEqual Expected, Actual, "Single field_collate_order mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQL")
Private Sub ztcSQL_ValidatesSingleTwoFields()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = "CREATE INDEX ""idx_contacts_email"" ON ""contacts""(""email"", ""domain"")"
Act:
    Dim Actual As String
    Actual = IndexO("idx_contacts_email", "contacts", Array(Array("email"), Array("domain"))).SQL
Assert:
    Assert.AreEqual Expected, Actual, "Two fields mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQL")
Private Sub ztcSQL_ValidatesSingleTwoFieldsOrder()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = "CREATE INDEX ""idx_contacts_email"" ON ""contacts""(""email"" ASC, ""domain"" DESC)"
Act:
    Dim Actual As String
    Actual = IndexO("idx_contacts_email", "contacts", Array(Array("email", "ASC"), Array("domain", "DESC"))).SQL
Assert:
    Assert.AreEqual Expected, Actual, "Two fields with order mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub

