Attribute VB_Name = "SQLiteCHeaderTests"
'@Folder "SQLite.SQLiteC.Connection"
'@TestModule
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed, IndexedDefaultMemberAccess
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


'@TestMethod("Decode byte array to native type")
Private Sub ztcLoadHeader_VerifiesLoadedHeaderData()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Dim dbh As SQLiteCHeader
    Set dbc = FixObjC.GetDBCTempFuncWithData
Act:
    Set dbh = SQLiteCHeader(dbc)
    dbh.LoadHeader
Assert:
    Assert.AreEqual "SQLite format 3" & vbNullChar, dbh.Header.MagicHeaderString, "MagicHeaderString mismatch."
    Assert.AreEqual 4096, dbh.Header.PageSizeInBytes, "PageSizeInBytes mismatch"
    Assert.AreEqual 1, dbh.Header.SchemaCookie, "SchemaCookie mismatch"
    Assert.AreEqual 4, dbh.Header.SchemaFormat, "SchemaFormat mismatch"
    Assert.AreEqual 0, dbh.Header.AppId, "AppId mismatch"
    Assert.AreEqual 2, dbh.Header.ChangeCounter, "ChangeCounter mismatch"
    Assert.AreEqual dbc.VersionNumber, dbh.Header.SqliteVersion, "SqliteVersion mismatch"
    Assert.AreEqual 72, LBound(dbh.Header.Reserved), "Reserved LBound mismatch"
    Assert.AreEqual 91, UBound(dbh.Header.Reserved), "Reserved UBound mismatch"
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
