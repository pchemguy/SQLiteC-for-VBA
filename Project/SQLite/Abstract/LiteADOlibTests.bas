Attribute VB_Name = "LiteADOlibTests"
'@Folder "SQLite.Abstract"
'@TestModule
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed
'@IgnoreModule IndexedDefaultMemberAccess, UnhandledOnErrorResumeNext
Option Explicit
Option Private Module

Private Const MODULE_NAME As String = "LiteADOlibTests"
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


'@TestMethod("Utils")
Private Sub ztcMapFields_ValidatesFieldMap()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
Act:
    Dim FieldMap As Scripting.Dictionary
    Set FieldMap = LiteADOlib.MapFields(FixUtils.People2D)
Assert:
    Assert.AreEqual 8, FieldMap.Count, "FieldMap size mismatch."
    Assert.IsTrue FieldMap.Exists("gender"), "Missing field."
    Assert.AreEqual 4, FieldMap("gender"), "Wrong field index."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
