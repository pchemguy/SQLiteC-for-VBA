Attribute VB_Name = "SQLiteCHeaderTests"
'@Folder "SQLiteDBdev.Extra"
'@TestModule
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed, IndexedDefaultMemberAccess
Option Explicit
Option Private Module

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
End Sub


'This method runs once per module.
'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
    FixObjC.CleanUp
End Sub


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("Decode byte array to native type")
Private Sub ztcLoadHeader_VerifiesLoadedHeaderData()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Dim dbh As SQLiteCHeader
    Set dbc = FixObjC.GetDBCTmpFuncWithData
Act:
    Set dbh = SQLiteCHeader.Create(dbc.DbPathName)
    dbh.LoadHeader
Assert:
    Assert.AreEqual "SQLite format 3" & vbNullChar, dbh.Header.MagicHeaderString, "MagicHeaderString mismatch."
    Assert.AreEqual 4096, dbh.Header.PageSizeInBytes, "PageSizeInBytes mismatch"
    Assert.AreEqual 2, dbh.Header.SchemaCookie, "SchemaCookie mismatch"
    Assert.AreEqual 4, dbh.Header.SchemaFormat, "SchemaFormat mismatch"
    Assert.AreEqual 0, dbh.Header.AppId, "AppId mismatch"
    Assert.AreEqual 3, dbh.Header.ChangeCounter, "ChangeCounter mismatch"
    Assert.AreEqual dbc.VersionNumber, dbh.Header.SQLiteVersion, "SqliteVersion mismatch"
    Assert.AreEqual 72, LBound(dbh.Header.Reserved), "Reserved LBound mismatch"
    Assert.AreEqual 91, UBound(dbh.Header.Reserved), "Reserved UBound mismatch"
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Decode byte array to native type")
Private Sub ztcSInt16bFromBytesBE_VerifiesIntegerDecoding()
    On Error GoTo TestFail

Arrange:
    Dim TestArray(0 To 1) As Byte
Act:
Assert:
    TestArray(0) = &HFF
    TestArray(1) = &HFF
    Assert.AreEqual -1, SQLiteCHeader.SInt16bFromBytesBE(TestArray), "-1 mismatch."
    TestArray(0) = 0
    TestArray(1) = &HFF
    Assert.AreEqual &HFF, SQLiteCHeader.SInt16bFromBytesBE(TestArray), "&HFF mismatch."
    TestArray(0) = &HFF
    TestArray(1) = 0
    Assert.AreEqual &HFF00, SQLiteCHeader.SInt16bFromBytesBE(TestArray), "&HFF00 mismatch."

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
    Dim TestArray(0 To 3) As Byte
Assert:
    TestArray(0) = &HFF
    TestArray(1) = &HFF
    TestArray(2) = &HFF
    TestArray(3) = &HFF
    Assert.AreEqual -1, SQLiteCHeader.SLong32bFromBytesBE(TestArray), "-1 mismatch."
    TestArray(0) = 0
    TestArray(1) = 0
    TestArray(2) = 0
    TestArray(3) = &HFF
    Assert.AreEqual &HFF&, SQLiteCHeader.SLong32bFromBytesBE(TestArray), "&HFF mismatch."
    TestArray(0) = 0
    TestArray(1) = 0
    TestArray(2) = &HFF
    TestArray(3) = 0
    Assert.AreEqual &HFF00&, SQLiteCHeader.SLong32bFromBytesBE(TestArray), "&HFF00 mismatch."
    TestArray(0) = 0
    TestArray(1) = &HFF
    TestArray(2) = 0
    TestArray(3) = 0
    Assert.AreEqual &HFF0000, SQLiteCHeader.SLong32bFromBytesBE(TestArray), "&HFF0000 mismatch."
    TestArray(0) = &HFF
    TestArray(1) = 0
    TestArray(2) = 0
    TestArray(3) = 0
    Assert.AreEqual &HFF000000, SQLiteCHeader.SLong32bFromBytesBE(TestArray), "&HFF000000 mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Flip endianness")
Private Sub ztcSLong32bFlipBytes_VerifiesFlippedLong()
    On Error GoTo TestFail

Arrange:
Act:
Assert:
    Assert.AreEqual -1, SQLiteCHeader.SLong32bFlipBytes(&HFFFFFFFF), "-1 mismatch."
    Assert.AreEqual &H11223344, SQLiteCHeader.SLong32bFlipBytes(&H44332211), "&H44332211 mismatch."
    Assert.AreEqual &HDDCCBBAA, SQLiteCHeader.SLong32bFlipBytes(&HAABBCCDD), "&HAABBCCDD mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
