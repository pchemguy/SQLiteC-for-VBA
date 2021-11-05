Attribute VB_Name = "ILiteADOTests"
'@Folder "SQLite.Abstract"
'@TestModule
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed
'@IgnoreModule IndexedDefaultMemberAccess, UnhandledOnErrorResumeNext
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


'@TestMethod("INSERT")
Private Sub ztcExecuteNonQuery_VerifiesInsertPlainITRB()
    On Error GoTo TestFail

Arrange:
    Dim dbm As ILiteADO
    Set dbm = FixObjAdo.GetDBMMemITRB
    Assert.IsNotNothing dbm, "FixObjAdo.GetDBMMemITRB returned Nothing."
Act:
    Dim SQLQuery As String
    SQLQuery = FixSQLITRB.InsertPlain()
    Dim AffectedRecords As Long
    AffectedRecords = dbm.ExecuteNonQuery(SQLQuery)
    Dim ExpectedChanges As Long
    ExpectedChanges = Len(SQLQuery) - Len(Replace(SQLQuery, "(", vbNullString)) - 1
Assert:
    Assert.AreEqual ExpectedChanges, AffectedRecords, "AffectedRecords mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
