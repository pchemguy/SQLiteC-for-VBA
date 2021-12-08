Attribute VB_Name = "GuardTests"
Attribute VB_Description = "Tests for the Guard class."
'@Folder "Common.Guard"
'@TestModule
'@ModuleDescription "Tests for the Guard class."
'@IgnoreModule LineLabelNotUsed, UnhandledOnErrorResumeNext
Option Explicit
Option Private Module

Private Const MODULE_NAME As String = "GuardTests"
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


'@TestMethod("Guard.EmptyString")
Private Sub EmptyString_Pass()
    On Error Resume Next
    TestCounter = TestCounter + 1
    Guard.EmptyString "Non-empty string"
    Guard.AssertExpectedError Assert, ErrNo.PassedNoErr
End Sub

'@TestMethod("Guard.EmptyString")
Private Sub EmptyString_ThrowsIfNotString()
    On Error Resume Next
    TestCounter = TestCounter + 1
    Guard.EmptyString True
    Guard.AssertExpectedError Assert, ErrNo.TypeMismatchErr
End Sub

'@TestMethod("Guard.EmptyString")
Private Sub EmptyString_ThrowsIfEmptyString()
    On Error Resume Next
    TestCounter = TestCounter + 1
    Guard.EmptyString vbNullString
    Guard.AssertExpectedError Assert, ErrNo.EmptyStringErr
End Sub

'@TestMethod("Guard.ObjectNotSet")
Private Sub ObjectNotSet_Pass()
    On Error Resume Next
    TestCounter = TestCounter + 1
    Guard.NullReference Guard
    Guard.AssertExpectedError Assert, ErrNo.PassedNoErr
End Sub

'@TestMethod("Guard.ObjectNotSet")
Private Sub ObjectNotSet_ThrowsIfNotObject()
    On Error Resume Next
    TestCounter = TestCounter + 1
    Guard.NullReference Empty
    Guard.AssertExpectedError Assert, ErrNo.ObjectRequiredErr
End Sub

'@TestMethod("Guard.ObjectNotSet")
Private Sub ObjectNotSet_ThrowsIfNothing()
    On Error Resume Next
    TestCounter = TestCounter + 1
    Guard.NullReference Nothing
    Guard.AssertExpectedError Assert, ErrNo.ObjectNotSetErr
End Sub

'@TestMethod("Guard.ObjectSet")
Private Sub ObjectSet_Pass()
    On Error Resume Next
    TestCounter = TestCounter + 1
    Guard.NonNullReference Nothing
    Guard.AssertExpectedError Assert, ErrNo.PassedNoErr
End Sub

'@TestMethod("Guard.ObjectSet")
Private Sub ObjectSet_ThrowsIfNotObject()
    On Error Resume Next
    TestCounter = TestCounter + 1
    Guard.NonNullReference Empty
    Guard.AssertExpectedError Assert, ErrNo.ObjectRequiredErr
End Sub

'@TestMethod("Guard.ObjectSet")
Private Sub ObjectSet_ThrowsIfNotNothing()
    On Error Resume Next
    TestCounter = TestCounter + 1
    Guard.NonNullReference Guard
    Guard.AssertExpectedError Assert, ErrNo.ObjectSetErr
End Sub

'@TestMethod("Guard.NonDefaultInstance")
Private Sub NonDefaultInstance_Pass()
    On Error Resume Next
    TestCounter = TestCounter + 1
    Guard.NonDefaultInstance Guard
    Guard.AssertExpectedError Assert, ErrNo.PassedNoErr
End Sub

'@TestMethod("Guard.NonDefaultInstance")
Private Sub NonDefaultInstance_ThrowsIfNothing()
    On Error Resume Next
    TestCounter = TestCounter + 1
    Guard.NonDefaultInstance Nothing
    Guard.AssertExpectedError Assert, ErrNo.ObjectNotSetErr
End Sub

'@TestMethod("Guard.DefaultInstance")
Private Sub DefaultInstance_ThrowsIfDefaultInstance()
    On Error Resume Next
    TestCounter = TestCounter + 1
    Guard.DefaultInstance Guard
    Guard.AssertExpectedError Assert, ErrNo.DefaultInstanceErr
End Sub


'@TestMethod("Guard.Self")
Private Sub Self_CheckAvailability()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim instanceVar As Object
    Set instanceVar = Guard.Create("Dummy")
Act:
    Dim selfVar As Object
    Set selfVar = instanceVar.Self
Assert:
    Assert.AreEqual TypeName(instanceVar), TypeName(selfVar), "Error: type mismatch: " & TypeName(selfVar) & " type."
    Assert.AreSame instanceVar, selfVar, "Error: bad Self pointer"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Guard.Class")
Private Sub Class_CheckAvailability()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim classVar As Object
    Set classVar = Guard
Act:
    Dim classVarReturned As Object
    Set classVarReturned = classVar.Create("Dummy").Class
Assert:
    Assert.AreEqual TypeName(classVar), TypeName(classVarReturned), "Error: type mismatch: " & TypeName(classVarReturned) & " type."
    Assert.AreSame classVar, classVarReturned, "Error: bad Class pointer"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
