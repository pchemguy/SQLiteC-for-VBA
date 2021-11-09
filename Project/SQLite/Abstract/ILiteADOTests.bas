Attribute VB_Name = "ILiteADOTests"
'@Folder "SQLite.Abstract"
'@TestModule
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed
'@IgnoreModule IndexedDefaultMemberAccess, UnhandledOnErrorResumeNext
Option Explicit
Option Private Module

Private Const MODULE_NAME As String = "ILiteADOTests"
Private TestCounter As Long

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
    With Logger
        .ClearLog
        .DebugLevelDatabase = DEBUGLEVEL_MAX
        .DebugLevelImmediate = DEBUGLEVEL_NONE
        .UseIdPadding = True
        .UseTimeStamp = False
        .RecordIdDigits 3
        .TimerSet MODULE_NAME
    End With
    TestCounter = 0
End Sub


'This method runs once per module.
'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
    Logger.TimerLogClear MODULE_NAME, TestCounter
    Logger.PrintLog
End Sub


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("INSERT")
Private Sub ztcExecuteNonQuery_VerifiesInsertPlainITRB()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

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
