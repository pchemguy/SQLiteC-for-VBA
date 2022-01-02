Attribute VB_Name = "FixLimitsTests"
'@Folder "SQLite.Fixtures"
'@TestModule
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed
'@IgnoreModule IndexedDefaultMemberAccess, UnhandledOnErrorResumeNext
Option Explicit
Option Private Module

Private Const MODULE_NAME As String = "FixLimitsTests"
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


'@TestMethod("Fixture")
Private Sub ztcLiteADOCreateRowidWithExtraData_VerifiesLiteADODbTmpCreateRowidWithExtraData()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbq As ILiteADO
    Set dbq = FixObjAdo.GetDbTmp
    Dim AffectedRows As Long
    AffectedRows = dbq.ExecuteNonQuery(FixSQLITRB.CreateRowidWithExtraData)
Act:
    Dim SQLQuery As String
    SQLQuery = FixSQLITRB.SelectNoRowid
    Dim Values As Variant
    Values = dbq.GetAdoRecordset(SQLQuery).GetRows
Assert:
    Assert.IsTrue IsArray(Values), "Query result not set."
    Assert.AreEqual 0, LBound(Values, 1), "Columns base mismatch."
    Assert.AreEqual 4, UBound(Values, 1), "Columns count mismatch."
    Assert.AreEqual 0, LBound(Values, 1), "Records base mismatch."
    Assert.AreEqual 4, UBound(Values, 1), "Records count mismatch."
    Assert.AreEqual 2147483647, Values(1, 0), "Record #1 xi mismatch."
    Assert.AreEqual 4294967296@, Values(1, 1), "Record #2 xi mismatch."
    Assert.AreEqual 900000000000000@, Values(1, 2), "Record #3 xi mismatch."
    Assert.AreEqual 922337203685477@, Values(1, 3), "Record #4 xi mismatch."
    Assert.AreEqual 900000000000000@, Values(1, 4), "Record #5 xi mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Fixture")
Private Sub ztcSQLiteCCreateRowidWithExtraData_VerifiesSQLiteCDbTmpCreateRowidWithExtraData()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbs As SQLiteCStatement
    Set dbs = FixObjC.GetDBSMemITRBWithExtraData
    Dim SQLQuery As String
    SQLQuery = FixSQLITRB.SelectNoRowid
    Dim Values As Variant
    Values = dbs.GetRowSet2D(SQLQuery)
Act:
    SQLQuery = SQLlib.CountSelect(FixSQLFunc.SelectNoRowid)
Assert:
    'Assert.AreEqual Expected, dbq.GetScalar(SQLQuery), "Row count mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
