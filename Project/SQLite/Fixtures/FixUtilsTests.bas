Attribute VB_Name = "FixUtilsTests"
'@Folder "SQLite.Fixtures"
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


'@TestMethod("ExpectedError")
Private Sub ztcByteArray_ThrowsOnMultipleStringArgs()
    On Error Resume Next
    Dim TargetArray() As Byte
    TargetArray = FixUtils.ByteArray("AA", "AB")
    Guard.AssertExpectedError Assert, ErrNo.TypeMismatchErr
End Sub


'@TestMethod("ExpectedError")
Private Sub ztcByteArray_ThrowsOnOutOfRangePositiveNumber()
    On Error Resume Next
    Dim TargetArray() As Byte
    TargetArray = FixUtils.ByteArray(Array(1, 256))
    Guard.AssertExpectedError Assert, ErrNo.TypeMismatchErr
End Sub


'@TestMethod("ExpectedError")
Private Sub ztcByteArray_ThrowsOnOutOfRangeNegativeNumber()
    On Error Resume Next
    Dim TargetArray() As Byte
    TargetArray = FixUtils.ByteArray(Array(1, -1))
    Guard.AssertExpectedError Assert, ErrNo.TypeMismatchErr
End Sub


'@TestMethod("ExpectedError")
Private Sub ztcByteArray_ThrowsOnOneBasedArray()
    On Error Resume Next
    Dim SourceArray(1 To 1) As Byte
    SourceArray(1) = Asc("A")
    Dim TargetArray() As Byte
    TargetArray = FixUtils.ByteArray(SourceArray)
    Guard.AssertExpectedError Assert, ErrNo.TypeMismatchErr
End Sub


'@TestMethod("ValidInput")
Private Sub ztcByteArray_VerifiesArrayFromListOfBytes()
    On Error GoTo TestFail

Arrange:
Act:
    Dim TargetArray() As Byte
    TargetArray = FixUtils.ByteArray(Asc("A"), Asc("B"), Asc("C"), Asc("D"))
Assert:
    Assert.AreEqual 0, LBound(TargetArray), "TargetArray base mismatch."
    Assert.AreEqual 3, UBound(TargetArray), "TargetArray size mismatch."
    Assert.AreEqual Asc("A"), TargetArray(0), "TargetArray(0) element mismatch."
    Assert.AreEqual Asc("B"), TargetArray(1), "TargetArray(1) element mismatch."
    Assert.AreEqual Asc("C"), TargetArray(2), "TargetArray(2) element mismatch."
    Assert.AreEqual Asc("D"), TargetArray(3), "TargetArray(3) element mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ValidInput")
Private Sub ztcByteArray_VerifiesArrayFromArrayOfBytes()
    On Error GoTo TestFail

Arrange:
    Dim SourceArray(0 To 3) As Byte
    SourceArray(0) = Asc("A")
    SourceArray(1) = Asc("B")
    SourceArray(2) = Asc("C")
    SourceArray(3) = Asc("D")
Act:
    Dim TargetArray() As Byte
    TargetArray = FixUtils.ByteArray(SourceArray)
Assert:
    Assert.AreEqual 0, LBound(TargetArray), "TargetArray base mismatch."
    Assert.AreEqual 3, UBound(TargetArray), "TargetArray size mismatch."
    Assert.AreEqual Asc("A"), TargetArray(0), "TargetArray(0) element mismatch."
    Assert.AreEqual Asc("B"), TargetArray(1), "TargetArray(1) element mismatch."
    Assert.AreEqual Asc("C"), TargetArray(2), "TargetArray(2) element mismatch."
    Assert.AreEqual Asc("D"), TargetArray(3), "TargetArray(3) element mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ValidInput")
Private Sub ztcByteArray_VerifiesArrayFromVariantArrayNumeric()
    On Error GoTo TestFail

Arrange:
    Dim SourceArray As Variant
    SourceArray = Array(Asc("A"), Asc("B"), Asc("C"), Asc("D"))
Act:
    Dim TargetArray() As Byte
    TargetArray = FixUtils.ByteArray(SourceArray)
Assert:
    Assert.AreEqual 0, LBound(TargetArray), "TargetArray base mismatch."
    Assert.AreEqual 3, UBound(TargetArray), "TargetArray size mismatch."
    Assert.AreEqual Asc("A"), TargetArray(0), "TargetArray(0) element mismatch."
    Assert.AreEqual Asc("B"), TargetArray(1), "TargetArray(1) element mismatch."
    Assert.AreEqual Asc("C"), TargetArray(2), "TargetArray(2) element mismatch."
    Assert.AreEqual Asc("D"), TargetArray(3), "TargetArray(3) element mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ValidInput")
Private Sub ztcByteArray_VerifiesArrayFromArrayOfChars()
    On Error GoTo TestFail

Arrange:
    Dim SourceArray As Variant
    SourceArray = Array("A", "B", "C", "D")
Act:
    Dim TargetArray() As Byte
    TargetArray = FixUtils.ByteArray(SourceArray)
Assert:
    Assert.AreEqual 0, LBound(TargetArray), "TargetArray base mismatch."
    Assert.AreEqual 3, UBound(TargetArray), "TargetArray size mismatch."
    Assert.AreEqual Asc("A"), TargetArray(0), "TargetArray(0) element mismatch."
    Assert.AreEqual Asc("B"), TargetArray(1), "TargetArray(1) element mismatch."
    Assert.AreEqual Asc("C"), TargetArray(2), "TargetArray(2) element mismatch."
    Assert.AreEqual Asc("D"), TargetArray(3), "TargetArray(3) element mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ValidInput")
Private Sub ztcByteArray_VerifiesArrayFromVariant()
    On Error GoTo TestFail

Arrange:
Act:
    Dim TargetArray() As Byte
    TargetArray = FixUtils.ByteArray("A", 66.2@, 66.9, 68&)
Assert:
    Assert.AreEqual 0, LBound(TargetArray), "TargetArray base mismatch."
    Assert.AreEqual 3, UBound(TargetArray), "TargetArray size mismatch."
    Assert.AreEqual Asc("A"), TargetArray(0), "TargetArray(0) element mismatch."
    Assert.AreEqual Asc("B"), TargetArray(1), "TargetArray(1) element mismatch."
    Assert.AreEqual Asc("C"), TargetArray(2), "TargetArray(2) element mismatch."
    Assert.AreEqual Asc("D"), TargetArray(3), "TargetArray(3) element mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ValidInput")
Private Sub ztcByteArray_VerifiesArrayFromString()
    On Error GoTo TestFail

Arrange:
Act:
    Dim TargetArray() As Byte
    TargetArray = FixUtils.ByteArray("ABCD")
Assert:
    Assert.AreEqual 0, LBound(TargetArray), "TargetArray base mismatch."
    Assert.AreEqual 3, UBound(TargetArray), "TargetArray size mismatch."
    Assert.AreEqual Asc("A"), TargetArray(0), "TargetArray(0) element mismatch."
    Assert.AreEqual Asc("B"), TargetArray(1), "TargetArray(1) element mismatch."
    Assert.AreEqual Asc("C"), TargetArray(2), "TargetArray(2) element mismatch."
    Assert.AreEqual Asc("D"), TargetArray(3), "TargetArray(3) element mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ByteArrayToString")
Private Sub ztcAsciiByteArrayToString_VerifiesStringFromArray()
    On Error GoTo TestFail

Arrange:
    Const Expected As String = "ABCD"
Act:
    Dim Actual As String
    Actual = FixUtils.AsciiByteArrayToString(FixUtils.ByteArray(Expected))
Assert:
    Assert.AreEqual Expected, Actual, "String mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ByteToHex")
Private Sub ztcByteToHex_VerifiesByteConversion()
    On Error GoTo TestFail

Arrange:
Act:
Assert:
    Assert.AreEqual "41", FixUtils.ByteToHex(&H41), "Code mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ValidInput")
Private Sub ztcKeysValuesToDict_VerifiesBasicDict()
    On Error GoTo TestFail

Arrange:
    Dim KeyValMap As Scripting.Dictionary
Act:
    Set KeyValMap = FixUtils.KeysValuesToDict(Array("Zero", "One"), Array(0, 1))
Assert:
    Assert.IsNotNothing KeyValMap, "KeyValMap is not set mismatch."
    Assert.AreEqual KeyValMap.CompareMode, TextCompare, "CompareMode mismatch."
    Assert.AreEqual 2, KeyValMap.Count, "Item count mismatch."
    Assert.IsTrue KeyValMap.Exists("Zero") And KeyValMap.Exists("One"), "Keys mismatch."
    Assert.IsFalse KeyValMap.Exists("Three"), "Unexpected key mismatch."
    Assert.IsTrue KeyValMap("Zero") = 0 And KeyValMap("One") = 1, "Values mismatch."
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ValidInput")
Private Sub ztcKeysValuesToDict_VerifiesExtendedDict()
    On Error GoTo TestFail

Arrange:
    Dim KeyValMap As Scripting.Dictionary
Act:
    Set KeyValMap = FixUtils.KeysValuesToDict( _
        Array("Long", "String", "Empty", "Null", "Object", "Array", "Error"), _
        Array(0&, "Text", Empty, Null, ThisWorkbook, Array(1), CVErr(1)))
Assert:
    Assert.IsNotNothing KeyValMap, "KeyValMap is not set mismatch."
    Assert.AreEqual KeyValMap.CompareMode, TextCompare, "CompareMode mismatch."
    Assert.AreEqual 7, KeyValMap.Count, "Item count mismatch."
    Assert.IsTrue KeyValMap.Exists("Long"), "'Long' key is missing."
    Assert.IsTrue KeyValMap.Exists("String"), "'String' key is missing."
    Assert.IsTrue KeyValMap.Exists("Empty"), "'Empty' key is missing."
    Assert.IsTrue KeyValMap.Exists("Null"), "'Null' key is missing."
    Assert.IsTrue KeyValMap.Exists("Object"), "'Object' key is missing."
    Assert.IsTrue KeyValMap.Exists("Array"), "'Array' key is missing."
    Assert.IsTrue KeyValMap.Exists("Error"), "'Error' key is missing."
    
    Assert.AreEqual vbLong, VarType(KeyValMap("Long")), "'Long' type mismatch."
    Assert.AreEqual vbString, VarType(KeyValMap("String")), "'String' type mismatch."
    Assert.AreEqual vbEmpty, VarType(KeyValMap("Empty")), "'Empty' type mismatch."
    Assert.AreEqual vbNull, VarType(KeyValMap("Null")), "'Null' type mismatch."
    Assert.AreEqual vbObject, VarType(KeyValMap("Object")), "'Object' type mismatch."
    Assert.AreEqual vbArray, VarType(KeyValMap("Array")) And vbArray, "'Array' type mismatch."
    Assert.AreEqual vbError, VarType(KeyValMap("Error")), "'Error' type mismatch."
    
    Assert.AreEqual 0, KeyValMap("Long"), "'Long' value mismatch."
    Assert.AreEqual "Text", KeyValMap("String"), "'String' value mismatch."
    Assert.IsTrue IsEmpty(KeyValMap("Empty")), "'Empty' value mismatch."
    Assert.IsTrue IsNull(KeyValMap("Null")), "'Null' value mismatch."
    Assert.IsTrue IsObject(KeyValMap("Object")), "'Object' value mismatch."
    Assert.AreSame ThisWorkbook, KeyValMap("Object"), "'Object' value mismatch."
    Assert.IsTrue IsArray(KeyValMap("Array")), "'Array' value mismatch."
    Assert.AreEqual 1, KeyValMap("Array")(0), "'Array' value mismatch."
    Assert.IsTrue IsError(KeyValMap("Error")), "'Error' value mismatch."
    Assert.AreEqual CVErr(1), KeyValMap("Error"), "'Error' value mismatch."
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ExpectedError")
Private Sub ztcKeysValuesToDict_ThrowsOnNotArrayKeys()
    On Error Resume Next
    Dim KeyValMap As Scripting.Dictionary
    Set KeyValMap = FixUtils.KeysValuesToDict("One", Array(0, 1))
    Guard.AssertExpectedError Assert, ErrNo.ExpectedArrayErr
End Sub


'@TestMethod("ExpectedError")
Private Sub ztcKeysValuesToDict_ThrowsOnNotArrayValues()
    On Error Resume Next
    Dim KeyValMap As Scripting.Dictionary
    Set KeyValMap = FixUtils.KeysValuesToDict(Array("Zero", "One"), 0)
    Guard.AssertExpectedError Assert, ErrNo.ExpectedArrayErr
End Sub


'@TestMethod("ExpectedError")
Private Sub ztcKeysValuesToDict_ThrowsOnArrayShapeSizeMismatch()
    On Error Resume Next
    Dim KeyValMap As Scripting.Dictionary
    Set KeyValMap = FixUtils.KeysValuesToDict(Array("Zero", "One"), Array(0, 1, 2))
    Guard.AssertExpectedError Assert, ErrNo.IncompatibleArraysErr
End Sub


'@TestMethod("ExpectedError")
Private Sub ztcXorElements_ThrowsOnTypeMismatch()
    On Error Resume Next
    Dim XorHash As Long
    XorHash = Array("A", 1, 2)
    Guard.AssertExpectedError Assert, ErrNo.TypeMismatchErr
End Sub


'@TestMethod("ValidInput")
Private Sub ztcXorElements_VerifiesHashes()
    On Error GoTo TestFail

Arrange:
Act:
Assert:
    Assert.AreEqual 3, FixUtils.XorElements(Array(0, 1, 2))
    Assert.AreEqual &HFFFFFFFF, FixUtils.XorElements(Array(&HFF&, &HFF00&, &HFF0000, &HFF000000))
    Assert.AreEqual 23, FixUtils.XorElements(Array(31, 7, 15))
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ExpectedError")
Private Sub ztcArrayXB_ThrowsOnTypeMismatch()
    On Error Resume Next
    Dim Result As Variant
    Result = FixUtils.ArrayXB("A", 1, 1)
    Guard.AssertExpectedError Assert, ErrNo.ExpectedArrayErr
End Sub


'@TestMethod("ValidInput")
Private Sub ztcArrayXB_VerifiesArrayFromArray()
    On Error GoTo TestFail

Arrange:
Act:
    Dim TargetArray() As Byte
    Dim Result As Variant
    Result = FixUtils.ArrayXB(TargetArray, 2, Array(1, 2, 3))
Assert:
    Assert.AreEqual 2, LBound(TargetArray), "TargetArray base mismatch."
    Assert.AreEqual 4, UBound(TargetArray), "TargetArray size mismatch."
    Assert.AreEqual 1, TargetArray(2), "TargetArray(2) element mismatch."
    Assert.AreEqual 2, TargetArray(3), "TargetArray(3) element mismatch."
    Assert.AreEqual 3, TargetArray(4), "TargetArray(4) element mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ValidInput")
Private Sub ztcArrayXB_VerifiesArrayFromParamArray()
    On Error GoTo TestFail

Arrange:
Act:
    Dim TargetArray() As Byte
    Dim Result As Variant
    Result = FixUtils.ArrayXB(TargetArray, 2, 1, 2, 3)
Assert:
    Assert.AreEqual 2, LBound(TargetArray), "TargetArray base mismatch."
    Assert.AreEqual 4, UBound(TargetArray), "TargetArray size mismatch."
    Assert.AreEqual 1, TargetArray(2), "TargetArray(2) element mismatch."
    Assert.AreEqual 2, TargetArray(3), "TargetArray(3) element mismatch."
    Assert.AreEqual 3, TargetArray(4), "TargetArray(4) element mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Decode byte array to native type")
Private Sub ztcSInt16bFromBytesBE_VerifiesIntegerDecoding()
    On Error GoTo TestFail

Arrange:
Act:
Assert:
    Assert.AreEqual -1, FixUtils.SInt16bFromBytesBE(FixUtils.ByteArray(&HFF, &HFF)), "-1 mismatch."
    Assert.AreEqual &HFF, FixUtils.SInt16bFromBytesBE(FixUtils.ByteArray(0, &HFF)), "&HFF mismatch."
    Assert.AreEqual &HFF00, FixUtils.SInt16bFromBytesBE(FixUtils.ByteArray(&HFF, 0)), "&HFF00 mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Decode byte array to native type")
Private Sub ztcSLong32bFromBytesBE_VerifiesIntegerDecoding()
    On Error GoTo TestFail

Arrange:
Act:
Assert:
    Assert.AreEqual -1, FixUtils.SLong32bFromBytesBE(FixUtils.ByteArray(&HFF, &HFF, &HFF, &HFF)), "-1 mismatch."
    Assert.AreEqual &HFF&, FixUtils.SLong32bFromBytesBE(FixUtils.ByteArray(0, 0, 0, &HFF)), "&HFF mismatch."
    Assert.AreEqual &HFF00&, FixUtils.SLong32bFromBytesBE(FixUtils.ByteArray(0, 0, &HFF, 0)), "&HFF00 mismatch."
    Assert.AreEqual &HFF0000, FixUtils.SLong32bFromBytesBE(FixUtils.ByteArray(0, &HFF, 0, 0)), "&HFF0000 mismatch."
    Assert.AreEqual &HFF000000, FixUtils.SLong32bFromBytesBE(FixUtils.ByteArray(&HFF, 0, 0, 0)), "&HFF000000 mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
