Attribute VB_Name = "DllCallTests"
'@Folder "SQLite.C.Config"
'@TestModule
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed
'@IgnoreModule UnhandledOnErrorResumeNext, StopKeyword
'@IgnoreModule IndexedDefaultMemberAccess, FunctionReturnValueDiscarded, UseMeaningfulName
Option Explicit
Option Private Module

Private Const MODULE_NAME As String = "DllCallTests"
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


'@TestMethod("ProcAddress")
Private Sub ztcProcAddressGet_VerifiesProcAddress()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbm As SQLiteC
    Set dbm = FixObjC.GetDBM
    Dim DllMan As DllManager
    Set DllMan = dbm.DllMan
    Dim LoadResult As DllLoadStatus
    LoadResult = DllMan.Load("kernel32", , False)
    Dim dbConf As DllCall
    Set dbConf = DllCall(DllMan)
Assert:
    Assert.AreNotEqual 0, dbConf.ProcAddressGet("kernel32", "GetProcAddress"), "Failed to get an address."

CleanExit:
    Exit Sub
TestFail:
    If Not Assert Is Nothing Then
        Assert.Fail "Error: " & Err.Number & " - " & Err.Description
    Else
        Debug.Print "Assert is Nothing. ## Error: " & Err.Number & " - " & Err.Description
    End If
End Sub


'@TestMethod("IndirectCall")
Private Sub ztcIndirectCall_VerifiesFunc0ArgsReturnLong()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim dbm As SQLiteC
    Set dbm = FixObjC.GetDBM
    Dim DllMan As DllManager
    Set DllMan = dbm.DllMan
    Dim dbConf As DllCall
    Set dbConf = DllCall(DllMan)
Act:
    Dim Result As Long
    Result = dbConf.IndirectCall("SQLite3", "sqlite3_libversion_number", CC_STDCALL, vbLong, Empty)
Assert:
    Assert.IsTrue Result > 3 * 10 ^ 6, "Failed to call dll function/args-0/ret-Long."

CleanExit:
    Exit Sub
TestFail:
    If Not Assert Is Nothing Then
        Assert.Fail "Error: " & Err.Number & " - " & Err.Description
    Else
        Debug.Print "Assert is Nothing. ## Error: " & Err.Number & " - " & Err.Description
    End If
End Sub

