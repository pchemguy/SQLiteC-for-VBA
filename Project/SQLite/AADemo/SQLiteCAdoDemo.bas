Attribute VB_Name = "SQLiteCAdoDemo"
Attribute VB_Description = "Illustrates typical workflows for SQLiteCAdo"
'@Folder "SQLite.AADemo"
'@ModuleDescription "Illustrates typical workflows for SQLiteCAdo"
'@IgnoreModule
Option Explicit
Option Private Module

Private Const LITE_LIB As String = "SQLiteCAdo"
Private Const PATH_SEP As String = "\"
Private Const LITE_RPREFIX As String = "Library" & PATH_SEP & LITE_LIB & PATH_SEP

Private Type TSQLiteCAdoDemo
    DbPathName As String
    dbmC As SQLiteC
    dbmADO As LiteMan
    dbs As SQLiteCStatement
    dbq As ILiteADO
End Type
Private this As TSQLiteCAdoDemo


'''' ILiteADO/SQLiteC demo
Private Sub MainC()
    this.DbPathName = FixObjAdo.RandomTempFileName
    '''' The shortcut version:
    ''''     Set this.dbmC = SQLiteC("")
    ''''     Set this.dbq = this.dbmC.CreateConnection(this.DbPathName, True).ExecADO
    InitDBQC
    Debug.Print "Created blank db: " & this.dbq.MainDB
    
    DemoDBQ "C"
    CleanUp
End Sub


'''' ILiteADO/SQLiteAdo demo
Private Sub MainADO()
    this.DbPathName = FixObjAdo.RandomTempFileName
    Set this.dbmADO = LiteMan(this.DbPathName, True)
    Set this.dbq = this.dbmADO.ExecADO
    Debug.Print "Created blank db: " & this.dbq.MainDB
    
    DemoDBQ "ADO"
    CleanUp
End Sub


Private Sub CleanUp()
    With this
        Set .dbq = Nothing
        Set .dbs = Nothing
        Set .dbmADO = Nothing
        Set .dbmC = Nothing
    End With
End Sub


Private Sub DemoDBQ(Optional ByVal Subpackage As String = "C")
    Dim dbq As ILiteADO
    Set dbq = this.dbq
    
    Dim SQLQuery As String
    Dim AffectedRows As Long
    
    '''' ===== CREATE Functions table ===== ''''
    SQLQuery = FixSQLFunc.Create
    AffectedRows = dbq.ExecuteNonQuery(SQLQuery)
    '''' ========= INSERT records ========= ''''
    SQLQuery = FixSQLFunc.InsertData
    AffectedRows = dbq.ExecuteNonQuery(SQLQuery)
    
    Debug.Print "Number of inserted rows: " & CStr(AffectedRows)
    
    '''' ========= SELECT records ========= ''''
    Dim QueryParams As Scripting.Dictionary
    If Subpackage = "C" Then
        '''' ========= SQLiteC/ILiteADO supports PARAMS ========= ''''
        SQLQuery = FixSQLFunc.SelectFilteredParamName
        Set QueryParams = FixSQLFunc.SelectFilteredParamNameValues
    Else
        '''' ==== SQLiteADO/ILiteADO does not support PARAMS ==== ''''
        SQLQuery = FixSQLFunc.SelectFilteredPlain
        Set QueryParams = Nothing
    End If
    
    '''' ============== Get Recordset ============= ''''
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = dbq.GetAdoRecordset(SQLQuery, QueryParams)
    
    '''' ========= Conert into a 2D array ========= ''''
    Dim RowSet2D As Variant
    RowSet2D = ArrayLib.TransposeArray(AdoRecordset.GetRows)
        
    Debug.Print "Number of selected rows: " & CStr(AdoRecordset.RecordCount)
    
    '''' ================================================================= ''''
    '''' ================================================================= ''''
    
    '''' ===== CREATE ITRB table ===== ''''
    SQLQuery = FixSQLITRB.Create
    AffectedRows = dbq.ExecuteNonQuery(SQLQuery)
    '''' ========= INSERT records ========= ''''
    SQLQuery = FixSQLITRB.InsertPlain
    AffectedRows = dbq.ExecuteNonQuery(SQLQuery)

    Debug.Print "Number of inserted rows: " & CStr(AffectedRows)


    '''' ========= UPDATE records ========= ''''
    If Subpackage = "C" Then
        '''' ========= SQLiteC/ILiteADO supports PARAMS ========= ''''
        SQLQuery = FixSQLITRB.UpdateParamName
        Set QueryParams = FixSQLITRB.UpdateParamValueDict
    Else
        '''' ==== SQLiteADO/ILiteADO does not support PARAMS ==== ''''
        SQLQuery = FixSQLITRB.UpdatePlain
        Set QueryParams = Nothing
    End If
    AffectedRows = dbq.ExecuteNonQuery(SQLQuery, QueryParams)
    Debug.Print "Number of updated rows: " & CStr(AffectedRows)
End Sub


Private Sub InitDBQC()
    '------------------------'
    '===== INIT MANAGER ====='
    '------------------------'
    Dim DllPath As String
    DllPath = LITE_RPREFIX & "dll\" & ARCH
    Dim DllNames As Variant
    #If Win64 Then
        DllNames = "sqlite3.dll"
    #Else
        DllNames = Array( _
            "icudt" & SQL_ICU_V & ".dll", "icuuc" & SQL_ICU_V & ".dll", _
            "icuin" & SQL_ICU_V & ".dll", "icuio" & SQL_ICU_V & ".dll", _
            "icutu" & SQL_ICU_V & ".dll", "sqlite3.dll")
    #End If
    Dim dbm As SQLiteC
    '@Ignore IndexedDefaultMemberAccess
    Set dbm = SQLiteC(DllPath, DllNames)
    If dbm Is Nothing Then
        Err.Raise ErrNo.ObjectCreateErr, "SQLiteCExamples", _
                  "Failed to create an SQLiteC instance."
    Else
        Debug.Print "Database manager instance (SQLiteC class) is ready"
    End If
    
    '''' Test SQLite3.dll
    If Replace(dbm.Version(False), ".", "0") & "0" = CStr(dbm.Version) Then
        Debug.Print "Database engine version functionality test passed."
    Else
        Debug.Print "Database engine version functionality test failed."
    End If
    Set this.dbmC = dbm

    '---------------------------'
    '===== INIT CONNECTION ====='
    '---------------------------'
    Dim dbc As SQLiteCConnection
    Set dbc = dbm.CreateConnection(this.DbPathName, AllowNonExistent:=True)
    If dbc Is Nothing Then
        Err.Raise ErrNo.ObjectCreateErr, "SQLiteCExamples", _
                  "Failed to create an SQLiteCConnection instance."
    Else
        Debug.Print "Database SQLiteCConnection instance is ready."
    End If

    '--------------------------'
    '===== INIT STATEMENT ====='
    '--------------------------'
    Dim DbStmtName As String
    DbStmtName = vbNullString
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(DbStmtName)
    Set this.dbs = dbs
    Dim dbq As ILiteADO
    Set dbq = dbs
    If dbq Is Nothing Then
        Err.Raise ErrNo.ObjectCreateErr, "SQLiteCExamples", _
                  "Failed to create an SQLiteCStatement instance."
    Else
        Debug.Print "Database SQLiteCStatement instance is ready."
    End If
    '''' Maximum capapacity of 100x10 = 1000 rows
    dbs.DbExecutor.PageCount = 10
    dbs.DbExecutor.PageSize = 100
    Set this.dbq = dbq
End Sub
