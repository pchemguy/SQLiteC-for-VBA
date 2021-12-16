Attribute VB_Name = "SQLiteCAdoTutorial"
'@Folder "SQLite.AADemo"
'@IgnoreModule
Option Explicit

Private Type TSQLiteCAdoTutorial
    DbPathName As String
    dbmC As SQLiteC
    dbmADO As LiteMan
    dbs As SQLiteCStatement
    dbq As ILiteADO
End Type
Private this As TSQLiteCAdoTutorial


Private Sub CleanUp()
    With this
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
    Debug.Print "Created blank db: " & this.dbq.MainDB
    
    Dim SQLQuery As String
    SQLQuery = SQLCreateTablePeople()
    Dim Result As Long
    Result = this.dbq.ExecuteNonQuery(SQLQuery)
    
    Dim TableName As String
    TableName = "main.people"
    Dim SQLTool As SQLlib
    Set SQLTool = SQLlib(TableName)
    Dim TableData As Variant
    TableData = FixPeopleData.UsedRange.Value2
    SQLQuery = SQLTool.INSERTVALUESFrom2DArray(TableData)
    Result = this.dbq.ExecuteNonQuery(SQLQuery)
    
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
