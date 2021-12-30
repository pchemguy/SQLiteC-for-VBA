Attribute VB_Name = "ILiteADOTests"
'@Folder "SQLite.Abstract"
'@TestModule
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed
'@IgnoreModule IndexedDefaultMemberAccess, UnhandledOnErrorResumeNext
Option Explicit
Option Private Module

Private Const MODULE_NAME As String = "ILiteADOTests"
Private TestCounter As Long

#Const LateBind = 0     '''' RubberDuck Tests
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


'@TestMethod("Introspection")
Private Sub ztcClassNameGet_VerifiesClassNameGetter()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbq As ILiteADO
Act:
Assert:
    Set dbq = FixObjAdo.GetDbMem
    Assert.AreEqual ":memory:", dbq.MainDB, "In-memory database path mismatch."
    Assert.AreEqual "LiteADO", dbq.ClassName, "Implementing class name mismatch."
    Set dbq = FixObjC.GetDBCMem.CreateStatement(vbNullString)
    Assert.AreEqual ":memory:", dbq.MainDB, "In-memory database path mismatch."
    Assert.AreEqual "SQLiteCStatement", dbq.ClassName, "Implementing class name mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("INSERT")
Private Sub ztcExecuteNonQuery_VerifiesInsertPlainITRB()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbq As ILiteADO
    Set dbq = FixObjAdo.GetDbMemITRB
    Assert.IsFalse dbq Is Nothing, "FixObjAdo.GetDBMMemITRB returned Nothing."
Act:
    Dim SQLQuery As String
    SQLQuery = FixSQLITRB.InsertPlain()
    Dim AffectedRecords As Long
    AffectedRecords = dbq.ExecuteNonQuery(SQLQuery)
    Dim ExpectedChanges As Long
    ExpectedChanges = Len(SQLQuery) - Len(Replace(SQLQuery, "(", vbNullString)) - 1
Assert:
    Assert.AreEqual ExpectedChanges, AffectedRecords, "AffectedRecords mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
