Attribute VB_Name = "QSortTests"
'@Folder "Common.QuickSort"
'@TestModule
'@IgnoreModule AssignmentNotUsed, VariableNotUsed, LineLabelNotUsed
'@IgnoreModule UnhandledOnErrorResumeNext, IndexedDefaultMemberAccess, FunctionReturnValueDiscarded
Option Explicit
Option Private Module

Private Const MODULE_NAME As String = "QSortTests"
Private TestCounter As Long

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


'@TestMethod("QSort")
Private Sub ztcNumericArrayFullSort()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim Sample() As Variant
    Sample = Array(45, 30, 25, 15, 10, 5, 40, 20, 35, 50, 75, 85, 60, 80, 55, 65, 70, 75)
Act:
    QSort.Vector Sample
Assert:
    Assert.AreEqual "51015202530354045505560657075758085", Join(Sample, vbNullString)

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("QSort")
Private Sub ztcNumericArrayPartialSort()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim Sample() As Variant
    Sample = Array(45, 30, 25, 15, 10, 5, 40, 20, 35, 50, 75, 85, 60, 80, 55, 65, 70, 75)
Act:
    QSort.Vector Sample, 5, 12
Assert:
    Assert.AreEqual "45302515105203540506075858055657075", Join(Sample, vbNullString)

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("QSort")
Private Sub ztcNumericArray1FullSort()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim Sample() As Variant
    Sample = Array(45)
Act:
    QSort.Vector Sample
Assert:
    Assert.AreEqual 45, Sample(0)

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("QSort")
Private Sub ztcNumericArray2FullSort()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim Sample() As Variant
    Sample = Array(45, 15)
Act:
    QSort.Vector Sample
Assert:
    Assert.AreEqual 45, Sample(1)

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("QSort")
Private Sub ztcTextArrayFullSort()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim Sample() As Variant
    Sample = Array("Kas", "Qman", "Cs", "Ib", "Zd", "Csg", "bs", "afeee", "i", "Oddd")
Act:
    QSort.Vector Sample
Assert:
    Assert.AreEqual "afeeebsCsCsgiIbKasOdddQmanZd", Join(Sample, vbNullString)

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("QSort")
Private Sub ztcTextArrayPartialSort()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim Sample() As Variant
    Sample = Array("Kas", "Qman", "Cs", "Ib", "Zd", "Csg", "bs", "afeee", "i", "Oddd")
Act:
    QSort.Vector Sample, 2, 7
Assert:
    Assert.AreEqual "KasQmanafeeebsCsCsgIbZdiOddd", Join(Sample, vbNullString)

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("QSort")
Private Sub ztcNumericArrayFullSortReturn()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim Sample() As Variant
    Sample = Array(45, 30, 25, 15, 10, 5, 40, 20, 35, 50, 75, 85, 60, 80, 55, 65, 70, 75)
Act:
    Dim Returned As Variant
    Returned = QSort.Vector(Sample)
Assert:
    Assert.AreEqual "51015202530354045505560657075758085", Join(Returned, vbNullString)

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
