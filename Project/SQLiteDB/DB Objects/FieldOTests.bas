Attribute VB_Name = "FieldOTests"
'@Folder "SQLiteDB.DB Objects"
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
Private Sub ztcCheckDefault_ValidatesNumeric()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = "(3.14)"
Act:
    Dim Actual As String
    Actual = FieldO("id").CheckDefault(3.14)
Assert:
    Assert.AreEqual Expected, Actual, "Numeric value mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Input Validation")
Private Sub ztcCheckDefault_ValidatesText()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = "('3.14')"
Act:
    Dim Actual As String
    Actual = FieldO("id").CheckDefault("3.14")
Assert:
    Assert.AreEqual Expected, Actual, "Text value mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Input Validation")
Private Sub ztcCheckDefault_ValidatesFormula()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = "(asin(0.5) * 6 AS pi)"
Act:
    Dim Actual As String
    Actual = FieldO("id").CheckDefault("(asin(0.5) * 6 AS pi)")
Assert:
    Assert.AreEqual Expected, Actual, "Formula value mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Input Validation")
Private Sub ztcCheckDefault_ThrowsIfTextWithQuote()
    On Error Resume Next
    Debug.Assert Not FieldO("id").CheckDefault("can't") Is Nothing
    AssertExpectedError Assert, ErrNo.InvalidCharacterErr
End Sub


'@TestMethod("Input Validation")
Private Sub ztcCheckName_ThrowsIfNameEmpty()
    On Error Resume Next
    Dim Field As FieldO: Set Field = FieldO(vbNullString)
    AssertExpectedError Assert, ErrNo.EmptyStringErr
End Sub


'@TestMethod("Input Validation")
Private Sub ztcCheckName_ThrowsIfNameWithSpace()
    On Error Resume Next
    Dim Field As FieldO: Set Field = FieldO("i d")
    AssertExpectedError Assert, ErrNo.InvalidCharacterErr
End Sub


'@TestMethod("Input Validation")
Private Sub ztcCheckName_ThrowsIfNameWithQuote()
    On Error Resume Next
    Dim Field As FieldO: Set Field = FieldO("i'd")
    AssertExpectedError Assert, ErrNo.InvalidCharacterErr
End Sub


'@TestMethod("Input Validation")
Private Sub ztcCheckName_ThrowsIfNameWithDash()
    On Error Resume Next
    Dim Field As FieldO: Set Field = FieldO("i-d")
    AssertExpectedError Assert, ErrNo.InvalidCharacterErr
End Sub


'@TestMethod("Input Validation")
Private Sub ztcCheckName_ValidatesName()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = "First_name_1"
Act:
    Dim Actual As String
    Actual = FieldO("id").CheckName("First_name_1")
Assert:
    Assert.AreEqual Expected, Actual, "Field name mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQL")
Private Sub ztcSQL_ValidatesName()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = "    ""id"""
Act:
    Dim Actual As String
    Actual = FieldO("id").SQL
Assert:
    Assert.AreEqual Expected, Actual, "Field name mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQL")
Private Sub ztcSQL_ValidatesNameType()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = "    ""id""    VARCHAR(50)"
Act:
    Dim Actual As String
    Actual = FieldO("id", "VARCHAR(50)").SQL
Assert:
    Assert.AreEqual Expected, Actual, "Field name_type mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQL")
Private Sub ztcSQL_ValidatesNameTypeNull()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = "    ""id""    TEXT NOT NULL"
Act:
    Dim Actual As String
    Actual = FieldO("id", "TEXT", True).SQL
Assert:
    Assert.AreEqual Expected, Actual, "Field name_type_null mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQL")
Private Sub ztcSQL_ValidatesNameTypeDefaultNumber()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = "    ""id""    TEXT DEFAULT (3.14)"
Act:
    Dim Actual As String
    Actual = FieldO("id", "TEXT", , 3.14).SQL
Assert:
    Assert.AreEqual Expected, Actual, "Field name_type_default_number mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQL")
Private Sub ztcSQL_ValidatesNameTypeDefaultText()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = "    ""id""    TEXT DEFAULT ('3.14')"
Act:
    Dim Actual As String
    Actual = FieldO("id", "TEXT", , "3.14").SQL
Assert:
    Assert.AreEqual Expected, Actual, "Field name_type_default_text mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQL")
Private Sub ztcSQL_ValidatesNameTypeDefaultFormula()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = "    ""id""    TEXT DEFAULT (asin(0.5) * 6)"
Act:
    Dim Actual As String
    Actual = FieldO("id", "TEXT", , "(asin(0.5) * 6)").SQL
Assert:
    Assert.AreEqual Expected, Actual, "Field name_type_default_formula mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQL")
Private Sub ztcSQL_ValidatesNameTypeCheck()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = "    ""id""    TEXT CHECK(id > 0)"
Act:
    Dim Actual As String
    Actual = FieldO("id", "TEXT", , , "id > 0").SQL
Assert:
    Assert.AreEqual Expected, Actual, "Field name_type_check mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQL")
Private Sub ztcSQL_ValidatesNameTypeUniqueCollate()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = "    ""id""    TEXT UNIQUE COLLATE NOCASE"
Act:
    Dim Actual As String
    Actual = FieldO("id", "TEXT", , , , True, "NOCASE").SQL
Assert:
    Assert.AreEqual Expected, Actual, "Field name_type_unique_collate mismatch"

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
    Expected = "    ""id""    TEXT NOT NULL DEFAULT ('3.14') CHECK(id > 0) UNIQUE COLLATE NOCASE"
Act:
    Dim Actual As String
    Actual = FieldO("id", "TEXT", True, "3.14", "id > 0", True, "NOCASE").SQL
Assert:
    Assert.AreEqual Expected, Actual, "Field name_type_unique_collate mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
