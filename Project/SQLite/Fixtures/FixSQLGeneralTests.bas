Attribute VB_Name = "FixSQLGeneralTests"
'@Folder "SQLite.Fixtures"
'@TestModule
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed, UnhandledOnErrorResumeNext
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


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("Select literal parameter")
Private Sub ztcSelectLiteralAtParam_VerifiesSelectLiteralQuery()
    On Error GoTo TestFail

Arrange:
    Dim ExpQuery As String
    Dim ActQuery As String
    Dim Literal As Variant
    
    Dim QueryLong As String
    QueryLong = "SELECT 10241024;"
    Dim QueryStr As String
    QueryStr = "SELECT 'ABC';"
Act:
Assert:
    ExpQuery = "SELECT @Literal;"
    ActQuery = FixSQLMisc.SelectLiteralAtParam
    Assert.AreEqual ActQuery, ExpQuery, "Template query mismatch"
    
    Literal = 1024&
    ExpQuery = "SELECT 1024;"
    ActQuery = FixSQLMisc.SelectLiteralAtParam(Literal)
    Assert.AreEqual ActQuery, ExpQuery, "Long literal query mismatch"

    Literal = "ABC"
    ExpQuery = "SELECT 'ABC';"
    ActQuery = FixSQLMisc.SelectLiteralAtParam(Literal)
    Assert.AreEqual ActQuery, ExpQuery, "String literal query mismatch"

    Literal = 3.14
    ExpQuery = "SELECT 3.14;"
    ActQuery = FixSQLMisc.SelectLiteralAtParam(Literal)
    Assert.AreEqual ActQuery, ExpQuery, "Double literal query mismatch"

    Literal = 102410241024102@
    ExpQuery = "SELECT 102410241024102;"
    ActQuery = FixSQLMisc.SelectLiteralAtParam(Literal)
    Assert.AreEqual ActQuery, ExpQuery, "Currency literal query mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
