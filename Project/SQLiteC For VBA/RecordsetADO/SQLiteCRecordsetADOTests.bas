Attribute VB_Name = "SQLiteCRecordsetADOTests"
'@Folder "SQLiteC For VBA.RecordsetADO"
'@TestModule
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed, UnhandledOnErrorResumeNext
Option Explicit
Option Private Module

#Const LateBind = LateBindTests
#If LateBind Then
    Private Assert As Object
#Else
    Private Assert As Rubberduck.PermissiveAssertClass
#End If
Private FixObj As SQLiteCTestFixObj
Private FixSQL As SQLiteCTestFixSQL


'This method runs once per module.
'@ModuleInitialize
Private Sub ModuleInitialize()
    #If LateBind Then
        Set Assert = CreateObject("Rubberduck.PermissiveAssertClass")
    #Else
        Set Assert = New Rubberduck.PermissiveAssertClass
    #End If
    Set FixObj = New SQLiteCTestFixObj
    Set FixSQL = New SQLiteCTestFixSQL
End Sub


'This method runs once per module.
'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
    Set FixObj = Nothing
End Sub


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("Query ADO Recordset")
Private Sub ztcAddMeta_InsertPlainSelectFromITRBTableRowid()
    On Error GoTo TestFail

    Set FixObj = New SQLiteCTestFixObj
    Set FixSQL = New SQLiteCTestFixSQL
Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObj.GetConnDbMemory
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)

    Dim ResultCode As SQLiteResultCodes
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    Dim AffectedRows As Long
Act:
    Dim SQLQuery As String
    SQLQuery = FixSQL.CREATETableITRBrowid
    ResultCode = dbc.ExecuteNonQueryPlain(SQLQuery, AffectedRows)
    SQLQuery = FixSQL.SELECTTestTable
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    ResultCode = dbs.DbExecutor.TableMetaCollect
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected GetTableMeta error."
    Assert.AreEqual 0, AffectedRows, "AffectedRows mismatch"
    
    Dim dbr As SQLiteCRecordsetADO
    Set dbr = SQLiteCRecordsetADO(dbs)
    dbr.AddMeta
Assert:
'    Assert.IsNotNothing dbr, "Unexpected error from FabRecordset."
'    Assert.AreEqual SQLQuery, dbs.SQLQueryOriginal, "Original query mismatch"
'    Assert.AreEqual 5, dbr.AdoRecordset.RecordCount, "Recordset.RecordCount mismatch"
'    Assert.AreEqual 6, dbr.AdoRecordset.Fields.Count, "Fields.Count mismatch"
Cleanup:
    ResultCode = dbs.Finalize
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Finalize error."
    ResultCode = dbc.CloseDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected CloseDb error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
