Attribute VB_Name = "SQLiteCAdoTutorial"
'@Folder "SQLite.AADemo"
'@IgnoreModule
Option Explicit

Private Type TSQLiteCAdoTutorial
    DbPathName As String
    dbmADO As LiteMan
    dbq As ILiteADO
    dbl As LiteADOlib
    dbmC As SQLiteC
    dbc As SQLiteCConnection
    dbs As SQLiteCStatement
End Type
Private this As TSQLiteCAdoTutorial


Private Sub CleanUp()
    With this
        Set .dbl = Nothing
        Set .dbq = Nothing
        Set .dbs = Nothing
        Set .dbmADO = Nothing
        Set .dbmC = Nothing
    End With
End Sub


'''' ILiteADO/SQLiteAdo demo
Private Sub MainADO()
    this.DbPathName = FixObjAdo.RandomTempFileName
    Set this.dbmADO = LiteMan(this.DbPathName, True)
    Set this.dbq = this.dbmADO.ExecADO
    Set this.dbl = LiteADOlib(this.dbq)
    Debug.Print "Created blank db: " & this.dbq.MainDB
    
    Dim SQLQuery As String
    SQLQuery = SQLCreateTablePeople()
    Dim AffectedRows As Long
    AffectedRows = this.dbq.ExecuteNonQuery(SQLQuery, Null)
    
    Dim TableName As String
    TableName = "main.people"
    Dim TableData As Variant
    TableData = ThisWorkbook.Worksheets("FixPeopleData").UsedRange.Value2
    AffectedRows = this.dbl.InsertSkipExistingFrom2D(TableName, TableData)
    AffectedRows = this.dbl.InsertUpdateExistingFrom2D(TableName, TableData, Array(2, 3, 5, 7, 99))
    
    CleanUp
End Sub


'''' ILiteADO/SQLiteC demo
Private Sub MainC()
    this.DbPathName = FixObjC.RandomTempFileName
    InitDBQC
    
    Dim SQLQuery As String
    SQLQuery = SQLCreateTablePeople()
    Dim AffectedRows As Long
    AffectedRows = this.dbq.ExecuteNonQuery(SQLQuery)
    
    Dim TableName As String
    TableName = "main.people"
    Dim TableData As Variant
    TableData = ThisWorkbook.Worksheets("FixPeopleData").UsedRange.Value2
    AffectedRows = this.dbl.InsertSkipExistingFrom2D(TableName, TableData)
    AffectedRows = this.dbl.InsertUpdateExistingFrom2D(TableName, TableData, Array(2, 3, 5, 7, 99))
    
    CleanUp
End Sub


Private Function SQLCreateTablePeople() As String
    SQLCreateTablePeople = Join(Array( _
        "CREATE TABLE people (", _
        "    id         INTEGER NOT NULL,", _
        "    first_name VARCHAR(255) NOT NULL COLLATE NOCASE,", _
        "    last_name  VARCHAR(255) NOT NULL COLLATE NOCASE,", _
        "    age        INTEGER,", _
        "    gender     VARCHAR(10)  COLLATE NOCASE,", _
        "    email      VARCHAR(255) NOT NULL UNIQUE COLLATE NOCASE,", _
        "    country    VARCHAR(255) COLLATE NOCASE,", _
        "    domain     VARCHAR(255) COLLATE NOCASE,", _
        "    PRIMARY KEY(id AUTOINCREMENT),", _
        "    UNIQUE(last_name, first_name, email),", _
        "    CHECK(18 <= ""Age"" <= 80),", _
        "    CHECK(""gender"" IN ('male', 'female'))", _
        ");", _
        "CREATE UNIQUE INDEX female_names_idx ON people (", _
        "    last_name,", _
        "    first_name", _
        ") WHERE gender = 'female';", _
        "CREATE UNIQUE INDEX male_names_idx ON people (", _
        "    last_name,", _
        "    first_name", _
        ") WHERE gender = 'male'" _
    ), vbNewLine)
End Function


Private Sub InitDBQC()
    '------------------------'
    '===== INIT MANAGER ====='
    '------------------------'
    Dim DllPath As String
    Dim DllNames As Variant
    #If Win64 Then
        DllPath = ThisWorkbook.Path & "\Library\SQLiteCAdo\dll\x64"
        DllNames = "sqlite3.dll"
    #Else
        DllPath = ThisWorkbook.Path & "\Library\SQLiteCAdo\dll\x32"
        DllNames = Array("icudt68.dll", "icuuc68.dll", "icuin68.dll", _
                         "icuio68.dll", "icutu68.dll", "sqlite3.dll")
    #End If
    Dim dbm As SQLiteC
    '@Ignore IndexedDefaultMemberAccess
    Set dbm = SQLiteC(DllPath, DllNames)
    If dbm Is Nothing Then
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
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
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
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
        Err.Raise ErrNo.UnknownClassErr, "SQLiteCExamples", _
                  "Failed to create an SQLiteCStatement instance."
    Else
        Debug.Print "Database SQLiteCStatement instance is ready."
    End If
    '''' Maximum capapacity of 100x10 = 1000 rows
    dbs.DbExecutor.PageCount = 10
    dbs.DbExecutor.PageSize = 100
    Set this.dbq = dbq
    Set this.dbl = LiteADOlib(this.dbq)
End Sub


