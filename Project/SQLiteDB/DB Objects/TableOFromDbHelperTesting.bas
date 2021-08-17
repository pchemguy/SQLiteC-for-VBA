Attribute VB_Name = "TableOFromDbHelperTesting"
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


'@TestMethod("SQL")
Private Sub ztcParseFKClause_ValidatesOnDelete()
    On Error GoTo TestFail

Arrange:
    Dim ExpectedOnDelete As String
    ExpectedOnDelete = "SET DEFAULT"
    Dim ExpectedOnUpdate As String
    ExpectedOnUpdate = "NO ACTION"
Act:
    Dim Actual As Scripting.Dictionary
    Set Actual = TableOFromDbHelper.ParseFKClause("oN dELETE SeT dEFAULT On upDATE no action")
Assert:
    Assert.AreEqual ExpectedOnDelete, Actual("OnDelete"), "OnDelete mismatch"
    Assert.AreEqual ExpectedOnUpdate, Actual("OnUpdate"), "OnUpdate mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub

