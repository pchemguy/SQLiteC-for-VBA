Attribute VB_Name = "SQLiteCTestFixSQLTests"
'@Folder "SQLite.SQLiteC For VBA.Test Fixtures"
'@TestModule
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed, UnhandledOnErrorResumeNext
Option Explicit
Option Private Module

#If LateBind Then
    Private Assert As Object
#Else
    Private Assert As Rubberduck.PermissiveAssertClass
#End If
Private FixObj As SQLiteCTestFixObj
Private FixSQL As SQLiteCTestFixSQL


'This method runs once per module.
'@ModuleInitialize
Private Sub ModuleInitialize()
    #If LateBind Then
        Set Assert = CreateObject("Rubberduck.PermissiveAssertClass")
    #Else
        Set Assert = New Rubberduck.PermissiveAssertClass
    #End If
    Set FixObj = New SQLiteCTestFixObj
    Set FixSQL = New SQLiteCTestFixSQL
End Sub


'This method runs once per module.
'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
    Set FixObj = Nothing
End Sub


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("Select literal parameter")
Private Sub ztcSELECTLiteralAtParam_VerifiesSelectLiteralQuery()
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
    ActQuery = FixSQL.SELECTLiteralAtParam
    Assert.AreEqual ActQuery, ExpQuery, "Template query mismatch"
    
    Literal = 1024&
    ExpQuery = "SELECT 1024;"
    ActQuery = FixSQL.SELECTLiteralAtParam(Literal)
    Assert.AreEqual ActQuery, ExpQuery, "Long literal query mismatch"

    Literal = "ABC"
    ExpQuery = "SELECT 'ABC';"
    ActQuery = FixSQL.SELECTLiteralAtParam(Literal)
    Assert.AreEqual ActQuery, ExpQuery, "String literal query mismatch"

    Literal = 3.14
    ExpQuery = "SELECT 3.14;"
    ActQuery = FixSQL.SELECTLiteralAtParam(Literal)
    Assert.AreEqual ActQuery, ExpQuery, "Double literal query mismatch"

    Literal = 102410241024102@
    ExpQuery = "SELECT 102410241024102;"
    ActQuery = FixSQL.SELECTLiteralAtParam(Literal)
    Assert.AreEqual ActQuery, ExpQuery, "Currency literal query mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
