Attribute VB_Name = "SQLlibTests"
'@Folder "SQLite.SQLlib"
'@TestModule
'@IgnoreModule LineLabelNotUsed, IndexedDefaultMemberAccess
Option Explicit
Option Private Module

Private Const MODULE_NAME As String = "SQLlibTests"
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
'===================== FIXTURES ====================='
'===================================================='


Private Function zfxGetSQL() As SQLlib
    Dim TableName As String
    TableName = "people"
    Set zfxGetSQL = SQLlib.Create(TableName)
End Function


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("SQL")
Private Sub ztcSelectAll_ValidatesWildcardQuery()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim SQL As SQLlib
    Set SQL = zfxGetSQL
    Dim Expected As String
    Expected = "SELECT * FROM [" & SQL.TableName & "]"
Act:
    Dim Actual As String
    Actual = SQL.SelectAll
Assert:
    Assert.AreEqual Expected, Actual, "Wildcard query mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQL")
Private Sub ztcSelectAll_ValidatesFieldsQuery()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim SQL As SQLlib
    Set SQL = zfxGetSQL
    Dim Expected As String
    Expected = "SELECT [id], [FirstName], [LastName] FROM [" & SQL.TableName & "]"
Act:
    Dim Actual As String
    Actual = SQL.SelectAll(Array("id", "FirstName", "LastName"))
Assert:
    Assert.AreEqual Expected, Actual, "Wildcard query mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQL")
Private Sub ztcSelectOne_ValidatesQuery()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim SQL As SQLlib
    Set SQL = zfxGetSQL
    Dim Expected As String
    Expected = "SELECT * FROM [" & SQL.TableName & "] LIMIT 1"
Act:
    Dim Actual As String
    SQL.Limit = 1
    Actual = SQL.SelectAll
    SQL.Limit = 0
Assert:
    Assert.AreEqual Expected, Actual, "SelectOne query mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQL")
Private Sub ztcAsText_ValidatesQuery()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim SQL As SQLlib
    Set SQL = zfxGetSQL
    Dim Expected As String
    Expected = "CAST([id] AS TEXT) AS [id]"
Act:
    Dim Actual As String
    Actual = SQL.AsText("id")
Assert:
    Assert.AreEqual Expected, Actual, "SQLAsText query mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQL")
Private Sub ztcSelectIdAsText_ValidatesQuery()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim SQL As SQLlib
    Set SQL = zfxGetSQL
    Dim Expected As String
    Expected = "SELECT CAST([id] AS TEXT) AS [id], [FirstName], [LastName], [Age] FROM [" & SQL.TableName & "]"
                
Act:
    Dim Actual As String
    Actual = SQL.SelectIdAsText(Array("id", "FirstName", "LastName", "Age"))
Assert:
    Assert.AreEqual Expected, Actual, "SelectIdAsText query mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQL")
Private Sub ztcSelectAllAsText_ValidatesQuery()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim SQL As SQLlib
    Set SQL = zfxGetSQL
    Dim Expected As String
    Expected = "SELECT CAST([id] AS TEXT) AS [id], [FirstName], [LastName], CAST([Age] AS TEXT) AS [Age], [Gender] FROM [" & SQL.TableName & "]"
Act:
    Dim Actual As String
    Actual = SQL.SelectAllAsText(Array("id", "FirstName", "LastName", "Age", "Gender"), _
                                 Array(adInteger, adVarWChar, adVarWChar, adInteger, adVarWChar))
Assert:
    Assert.AreEqual Expected, Actual, "SelectIdAsText query mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQL")
Private Sub ztcUpdateSingleRecord_ValidatesQuery()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim SQL As SQLlib
    Set SQL = zfxGetSQL
    Dim Expected As String
    Expected = "UPDATE [" & SQL.TableName & "] SET ([FirstName], [LastName], [Age], [Gender], [Email]) = (?, ?, ?, ?, ?) WHERE [id] = ?"
Act:
    Dim Actual As String
    Actual = SQL.UpdateSingleRecord(Array("id", "FirstName", "LastName", "Age", "Gender", "Email"))
Assert:
    Assert.AreEqual Expected, Actual, "UpdateSingleRecord query mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQL")
Private Sub ztcAttach_ValidatesQueryWithMissingAlias()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim SQL As SQLlib
    Set SQL = zfxGetSQL
    Dim DbPathName As String
    DbPathName = ThisWorkbook.Path & Application.PathSeparator & _
                 ThisWorkbook.VBProject.Name & ".db"
    Dim Expected As String
    Expected = "ATTACH '" & DbPathName & "' AS [" & ThisWorkbook.VBProject.Name & "]"
Act:
    Dim Actual As String
    Actual = SQL.Attach(DbPathName)
Assert:
    Assert.AreEqual Expected, Actual, "ATTACH query mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("SQL")
Private Sub ztcVacuum_ValidatesQueries()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
    Dim Expected As String
    Dim Actual As String
Act:
Assert:
    Expected = "VACUUM"
    Actual = SQLlib.Vacuum()
    Assert.AreEqual Expected, Actual, "VACUUM bare query mismatch"
    Expected = "VACUUM"
    Actual = SQLlib.Vacuum(vbNullString, vbNullString)
    Assert.AreEqual Expected, Actual, "VACUUM bare query mismatch"
    Expected = "VACUUM [memory]"
    Actual = SQLlib.Vacuum("memory")
    Assert.AreEqual Expected, Actual, "VACUUM query with alias mismatch"
    Expected = "VACUUM INTO 'C:\TEMP\qqq.db'"
    Actual = SQLlib.Vacuum(vbNullString, "C:\TEMP\qqq.db")
    Assert.AreEqual Expected, Actual, "VACUUM query with INTO mismatch"
    Expected = "VACUUM [main] INTO 'C:\TEMP\qq''q.db'"
    Actual = SQLlib.Vacuum("main", "C:\TEMP\qq'q.db")
    Assert.AreEqual Expected, Actual, "VACUUM query with alias and INTO mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Utils")
Private Sub ztcRowsToTable_ValidatesArray()
    On Error GoTo TestFail
    TestCounter = TestCounter + 1

Arrange:
Act:
    Dim Values As Variant
    Values = SQLlib.RowsToTable(Array( _
        Array("A", "B", "C"), _
        Array("1", "1", "1"), _
        Array("2", "2", "2"), _
        Array("3", "3", "3") _
    ))
Assert:
    Assert.AreEqual 0, LBound(Values, 1), "Row base mismatch."
    Assert.AreEqual 0, LBound(Values, 2), "Col base mismatch."
    Assert.AreEqual 3, UBound(Values, 1), "Row count mismatch."
    Assert.AreEqual 2, UBound(Values, 2), "Col count mismatch."
    Assert.AreEqual "ABC", Values(0, 0) & Values(0, 1) & Values(0, 2), "Header row mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
