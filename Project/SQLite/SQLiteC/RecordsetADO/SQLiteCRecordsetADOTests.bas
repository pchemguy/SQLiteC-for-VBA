Attribute VB_Name = "SQLiteCRecordsetADOTests"
'@Folder "SQLite.SQLiteC.RecordsetADO"
'@TestModule
'@IgnoreModule AssignmentNotUsed, LineLabelNotUsed, VariableNotUsed, ProcedureNotUsed, UnhandledOnErrorResumeNext, IndexedDefaultMemberAccess
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


'@TestMethod("Query ADO Recordset")
Private Sub ztcAddMeta_InsertPlainSelectFromITRBTableRowid()
    On Error GoTo TestFail

Arrange:
    Dim dbc As SQLiteCConnection
    Set dbc = FixObjC.GetDBCMem
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(vbNullString)

    Dim ResultCode As SQLiteResultCodes
    ResultCode = dbc.OpenDb
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected OpenDb error."
    Dim AffectedRows As Long
    Dim Attr As ADODB.FieldAttributeEnum
Act:
    Dim SQLQuery As String
    SQLQuery = FixSQLITRB.CreateRowid
    ResultCode = dbc.ExecuteNonQueryPlain(SQLQuery, AffectedRows)
    SQLQuery = FixSQLITRB.SelectRowid
    ResultCode = dbs.Prepare16V2(SQLQuery)
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected Prepare16V2 error."
    ResultCode = dbs.DbExecutor.TableMetaCollect
    Assert.AreEqual SQLITE_OK, ResultCode, "Unexpected GetTableMeta error."
    Assert.AreEqual 0, AffectedRows, "AffectedRows mismatch"
    
    Dim dbr As SQLiteCRecordsetADO
    Set dbr = SQLiteCRecordsetADO(dbs)
    dbr.AddMeta
    Dim TableMeta() As SQLiteCColumnMeta
    TableMeta = dbs.DbExecutor.TableMeta
    Dim Rst As ADODB.Recordset
    Set Rst = dbr.AdoRecordset
    Rst.Open
Assert:
    With TableMeta(0)
        Assert.AreEqual adInteger, .AdoType, "AdoType mismatch."
        Assert.AreEqual 0, .AdoSize, "AdoSize should be 0."
        Attr = adFldIsNullable + adFldMayBeNull + adFldKeyColumn + adFldRowID + adFldUpdatable
        Assert.AreEqual Attr, .AdoAttr, "AdoAttr mismatch"
    End With

    With TableMeta(3)
        Assert.AreEqual adVarWChar, .AdoType, "AdoType mismatch."
        Assert.AreEqual 8192, .AdoSize, "AdoSize should be 8192."
        Attr = adFldIsNullable + adFldMayBeNull + adFldUpdatable
        Assert.AreEqual Attr, .AdoAttr, "AdoAttr mismatch"
    End With

    With TableMeta(4)
        Assert.AreEqual adDouble, .AdoType, "AdoType mismatch."
        Assert.AreEqual 0, .AdoSize, "AdoSize should be 0."
        Attr = adFldUpdatable
        Assert.AreEqual Attr, .AdoAttr, "AdoAttr mismatch"
    End With

    With TableMeta(5)
        Assert.AreEqual adLongVarBinary, .AdoType, "AdoType mismatch."
        Assert.AreEqual 65535, .AdoSize, "AdoSize should be 65535."
        Attr = adFldIsNullable + adFldMayBeNull + adFldUpdatable + adFldLong
        Assert.AreEqual Attr, .AdoAttr, "AdoAttr mismatch"
    End With
    
    With Rst.Fields(0)
        Assert.AreEqual "rowid", .Name, "Rst field name mismatch"
        Assert.AreEqual 4, .DefinedSize, "Rst field DefinedSize mismatch"
        Assert.AreEqual adInteger, .Type, "Rst field type mismatch"
        Attr = adFldIsNullable + adFldMayBeNull + adFldKeyColumn + adFldRowID + adFldUpdatable + adFldFixed
        Assert.AreEqual Attr, .Attributes, "Rst field Attributes mismatch"
    End With
    
    With Rst.Fields(3)
        Assert.AreEqual "xt", .Name, "Rst field name mismatch"
        Assert.AreEqual 8192, .DefinedSize, "Rst field DefinedSize mismatch"
        Assert.AreEqual adVarWChar, .Type, "Rst field type mismatch"
        Attr = adFldIsNullable + adFldMayBeNull + adFldUpdatable
        Assert.AreEqual Attr, .Attributes, "Rst field Attributes mismatch"
    End With

    With Rst.Fields(4)
        Assert.AreEqual "xr", .Name, "Rst field name mismatch"
        Assert.AreEqual 8, .DefinedSize, "Rst field DefinedSize mismatch"
        Assert.AreEqual adDouble, .Type, "Rst field type mismatch"
        Attr = adFldFixed + adFldUpdatable
        Assert.AreEqual Attr, .Attributes, "Rst field Attributes mismatch"
    End With

    With Rst.Fields(5)
        Assert.AreEqual "xb", .Name, "Rst field name mismatch"
        Assert.AreEqual 65535, .DefinedSize, "Rst field DefinedSize mismatch"
        Assert.AreEqual adLongVarBinary, .Type, "Rst field type mismatch"
        Attr = adFldIsNullable + adFldMayBeNull + adFldUpdatable + adFldLong
        Assert.AreEqual Attr, .Attributes, "Rst field Attributes mismatch"
    End With
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
