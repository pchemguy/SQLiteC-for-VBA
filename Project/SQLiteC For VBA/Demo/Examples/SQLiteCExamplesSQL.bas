Attribute VB_Name = "SQLiteCExamplesSQL"
'@Folder "SQLiteC For VBA.Demo.Examples"
Option Explicit


Public Function SQLforGetSQLiteVersion() As String
    SQLforGetSQLiteVersion = Join(Array( _
        "SELECT sqlite_version();" _
    ), vbNewLine)
End Function

Public Function SQLforGetDbPath() As String
    SQLforGetDbPath = Join(Array( _
        "SELECT file FROM pragma_database_list;" _
    ), vbNewLine)
End Function

Public Function SQLforGetCollations() As String
    SQLforGetCollations = Join(Array( _
        "SELECT name FROM pragma_collation_list AS collations ORDER BY name;" _
    ), vbNewLine)
End Function

Public Function SQLforFunctionsTable() As String
    SQLforFunctionsTable = Join(Array( _
        "WITH functions AS (SELECT rowid, * FROM pragma_function_list)", _
        "SELECT * FROM functions ORDER BY name;" _
    ), vbNewLine)
End Function

Public Function SQLforFunctionsTableFiltered() As String
    SQLforFunctionsTableFiltered = Join(Array( _
        "WITH functions AS (SELECT rowid, * FROM pragma_function_list)", _
        "SELECT * FROM functions", _
        "WHERE ([builtin] = 1 OR [builtin] = 0 AND [flags] = 0) AND", _
        "      ([enc] = 'utf8' AND [narg] >= 0 AND [type] = 's')", _
        "ORDER BY name;" _
    ), vbNewLine)
End Function

Public Function SQLforFunctionsTableFilteredNamedParams() As String
    SQLforFunctionsTableFilteredNamedParams = Join(Array( _
        "WITH functions AS (SELECT rowid, * FROM pragma_function_list)", _
        "SELECT * FROM functions", _
        "WHERE ([builtin] = @builtinY OR [builtin] = @builtinN AND [flags] = @flags) AND", _
        "      ([enc] = @enc AND [narg] >= @narg AND [type] = @type)", _
        "ORDER BY name;" _
    ), vbNewLine)
End Function

Public Function SQLforFunctionsTableFilteredNamedParamsArray() As Variant
    SQLforFunctionsTableFilteredNamedParamsArray = Array( _
        1, _
        0, _
        0, _
        "utf8", _
        0, _
        "s" _
    )
End Function

Public Function SQLforFunctionsTableFilteredNamedParamsDict() As Scripting.Dictionary
    Dim QueryParams As Scripting.Dictionary
    Set QueryParams = New Scripting.Dictionary
    With QueryParams
        .CompareMode = TextCompare
        .Item("@builtinY") = 1
        .Item("@builtinN") = 0
        .Item("@flags") = 0
        .Item("@enc") = "utf8"
        .Item("@narg") = 0
        .Item("@type") = "s"
    End With
    Set SQLforFunctionsTableFilteredNamedParamsDict = QueryParams
End Function

Public Function SQLforCreateTestTable() As String
    SQLforCreateTestTable = Join(Array( _
        "CREATE TABLE t1(", _
        "    xi INTEGER,", _
        "    xt TEXT,", _
        "    xr REAL,", _
        "    xb BLOB", _
        ");" _
    ), vbNewLine)
End Function

Public Function SQLforInsertTestRows() As String
    SQLforInsertTestRows = Join(Array( _
        "INSERT INTO t1(rowid, xi,    xt,  xr,                  xb) ", _
        "VALUES        (    1, 10, 'AAA', 3.1, X'410A0D0942434445'),", _
        "              (    2, 20, 'BBB', 1.3, X'30310A0D09323334'),", _
        "              (    3,  8, 'AAA', 7.2, X'30310A0D32093334'),", _
        "              (    4, 27, 'DDD', 4.3, X'410A0D0942434445'),", _
        "              (    5,  8, 'BBB', 3.8, X'30310A0D32093334');" _
    ), vbNewLine)
End Function

Public Function SQLforSelectFromTestTable() As String
    SQLforSelectFromTestTable = Join(Array( _
        "SELECT rowid, * FROM t1;" _
    ), vbNewLine)
End Function
