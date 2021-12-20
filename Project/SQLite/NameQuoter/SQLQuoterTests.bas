Attribute VB_Name = "SQLQuoterTests"
'@Folder "SQLite.NameQuoter"
'@TestModule
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed, UnhandledOnErrorResumeNext
Option Explicit
Option Private Module

Private Const MODULE_NAME As String = "SQLQuoterTests"
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
    FixObjC.CleanUp
End Sub


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("Quoting")
Private Sub ztcQuoteSQLName_VerifiesQLNameQuoting()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbm As SQLiteC
    Set dbm = FixObjC.GetDBM
Act:
Assert:
    Assert.AreEqual "FirstName", SQLQuoter.QuoteSQLName("FirstName"), "Basic name quoting mismatch."
    Assert.AreEqual "FirstName", SQLQuoter.QuoteSQLName("[FirstName]"), "Quoted basic name quoting mismatch."
    Assert.AreEqual "FirstName", SQLQuoter.QuoteSQLName("""FirstName"""), "Quoted basic name quoting mismatch."
    Assert.AreEqual "First_Name", SQLQuoter.QuoteSQLName("First_Name"), "Basic name quoting mismatch."
    Assert.AreEqual """_First_Name""", SQLQuoter.QuoteSQLName("_First_Name"), "First char non-alpha quoting mismatch."
    Assert.AreEqual """1First_Name""", SQLQuoter.QuoteSQLName("1First_Name"), "First char non-alpha quoting mismatch."
    Assert.AreEqual """First Name""", SQLQuoter.QuoteSQLName("First Name"), "Space quoting mismatch."
    Assert.AreEqual """Group""", SQLQuoter.QuoteSQLName("Group"), "Keyword quoting mismatch."
CleanUp:

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
