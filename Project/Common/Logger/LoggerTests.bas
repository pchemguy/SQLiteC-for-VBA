Attribute VB_Name = "LoggerTests"
Attribute VB_Description = "Tests for the Logger class."
'@Folder "Common.Logger"
'@TestModule
'@ModuleDescription("Tests for the Logger class.")
'@IgnoreModule VariableNotUsed, AssignmentNotUsed, LineLabelNotUsed, UnhandledOnErrorResumeNext
Option Explicit
Option Private Module

#Const LateBind = 1     '''' RubberDuck Tests
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


'@TestMethod("Log Database")
Private Sub ztcLog_VerifyItemCount()
    On Error GoTo TestFail
    
Arrange:
    Dim LoggerInstance As Logger
    Set LoggerInstance = New Logger
Act:
    LoggerInstance.Logg "AAA"
    LoggerInstance.Logg "AAA"
Assert:
    Assert.AreEqual 2, LoggerInstance.LogDatabase.Count

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Log Database")
Private Sub ztcLog_VerifyItemCountWithClear()
    On Error GoTo TestFail
    
Arrange:
    Dim LoggerInstance As Logger
    Set LoggerInstance = New Logger
Act:
    With LoggerInstance
        .DebugLevelImmediate = DEBUGLEVEL_NONE
        .DebugLevelDatabase = DEBUGLEVEL_ERROR
        .RecordIdDigits 3
        .UseIdPadding = True
        .UseTimeStamp = True
        
        .Logg "AAA"
        .Logg "AAA"
        .ClearLog
        .Logg "AAA"
        .Logg "AAA"
        .Logg "AAA"
        
        .PrintLog
    End With
Assert:
    Assert.AreEqual 3, LoggerInstance.LogDatabase.Count

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Log Database")
Private Sub ztcLog_VerifyItemCountOnGlobalWithCustomDb()
    On Error GoTo TestFail
    
Arrange:
    Dim LogDb As Scripting.Dictionary
    Set LogDb = New Scripting.Dictionary
    LogDb.CompareMode = TextCompare
Act:
    Logger.Logg "AAA", LogDb
    Logger.Logg "AAA", LogDb
Assert:
    Assert.AreEqual 2, LogDb.Count

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Log Database")
Private Sub ztcLog_VerifyItemCountOnGlobalWithClearWithCustomDb()
    On Error GoTo TestFail
    
Arrange:
    Dim LogDb As Scripting.Dictionary
    Set LogDb = New Scripting.Dictionary
    LogDb.CompareMode = TextCompare
Act:
    With Logger
        .Logg "AAA", LogDb
        .Logg "AAA", LogDb
        .ClearLog LogDb
        .Logg "AAA", LogDb
        .Logg "AAA", LogDb
        .Logg "AAA", LogDb
    End With
Assert:
    Assert.AreEqual 3, LogDb.Count

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
