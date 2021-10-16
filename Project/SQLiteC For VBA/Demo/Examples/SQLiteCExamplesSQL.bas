Attribute VB_Name = "SQLiteCExamplesSQL"
'@Folder "SQLiteC For VBA.Demo.Examples"
Option Explicit


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


Public Function SQLforCreateTestTable() As String
    SQLforCreateTestTable = Join(Array( _
        "CREATE TABLE t1(", _
        "    xi INTEGER,", _
        "    xr REAL,", _
        "    xb BLOB,", _
        "    xn NUMERIC,", _
        "    xt Text", _
        ");" _
    ), vbNewLine)
End Function


Public Function SQLforInsertTestRows() As String
    SQLforInsertTestRows = Join(Array( _
        "INSERT INTO t1(rowid, xi, xr, xb, xn,    xt) ", _
        "VALUES        (    1, 10, .1,  1,  7, 'AAA'),", _
        "              (    2, 20, .3,  1, .2, 'BBB'),", _
        "              (    3,  8, .2,  1, 10, 'AAA'),", _
        "              (    4, 27, .3,  1, .9, 'DDD'),", _
        "              (    5,  2, .8,  1,  1, 'BBB');" _
    ), vbNewLine)
End Function

