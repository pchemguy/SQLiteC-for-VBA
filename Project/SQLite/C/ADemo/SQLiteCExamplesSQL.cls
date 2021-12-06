VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLiteCExamplesSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "SQLite.C.ADemo"
'@PredeclaredId
Option Explicit


Public Function GetSQLiteVersion() As String
    GetSQLiteVersion = Join(Array( _
        "SELECT sqlite_version();" _
    ), vbNewLine)
End Function

Public Function GetDbPath() As String
    GetDbPath = Join(Array( _
        "SELECT file FROM pragma_database_list;" _
    ), vbNewLine)
End Function

Public Function FunctionsTableSELECT() As String
    FunctionsTableSELECT = Join(Array( _
        "SELECT * FROM functions ORDER BY name;" _
    ), vbNewLine)
End Function

Public Function FunctionsPragmaTable() As String
    FunctionsPragmaTable = Join(Array( _
        "WITH functions AS (SELECT rowid, * FROM pragma_function_list)", _
        FunctionsTableSELECT _
    ), vbNewLine)
End Function

Public Function FunctionsTableCREATE() As String
    FunctionsTableCREATE = Join(Array( _
        "CREATE TABLE functions(", _
        "    name    TEXT COLLATE NOCASE NOT NULL,", _
        "    builtin INTEGER             NOT NULL,", _
        "    type    TEXT COLLATE NOCASE NOT NULL,", _
        "    enc     TEXT COLLATE NOCASE NOT NULL,", _
        "    narg    INTEGER             NOT NULL,", _
        "    flags   INTEGER             NOT NULL", _
        ");" _
    ), vbNewLine)
End Function

'''' This SQL command is a multi-statement "nonquery".
'''' Use step_exec API.
Public Function FunctionsTableCREATEWithData() As String
    FunctionsTableCREATEWithData = Join(Array( _
        "DROP TABLE IF EXISTS functions;", _
        FunctionsTableCREATE, _
        "INSERT INTO functions (rowid, name, builtin, type, enc, narg, flags)", _
        FunctionsPragmaTable _
    ), vbNewLine)
End Function

Public Function FunctionsPragmaTableFiltered() As String
    FunctionsPragmaTableFiltered = Join(Array( _
        "WITH functions AS (SELECT rowid, * FROM pragma_function_list)", _
        "SELECT * FROM functions", _
        "WHERE ([builtin] = 1 OR [builtin] = 0 AND [flags] = 0) AND", _
        "      ([enc] = 'utf8' AND [narg] >= 0 AND [type] = 's')", _
        "ORDER BY name;" _
    ), vbNewLine)
End Function

Public Function FunctionsTableFiltered() As String
    FunctionsTableFiltered = Join(Array( _
        "SELECT * FROM functions", _
        "WHERE ([builtin] = 1 OR [builtin] = 0 AND [flags] = 0) AND", _
        "      ([enc] = 'utf8' AND [narg] >= 0 AND [type] = 's')", _
        "ORDER BY name;" _
    ), vbNewLine)
End Function

Public Function FunctionsPragmaTableNamedParams() As String
    FunctionsPragmaTableNamedParams = Join(Array( _
        "WITH functions AS (SELECT rowid, * FROM pragma_function_list)", _
        "SELECT * FROM functions", _
        "WHERE ([builtin] = @builtinY OR [builtin] = @builtinN AND [flags] = @flags) AND", _
        "      ([enc] = @enc AND [narg] >= @narg AND [type] = @type)", _
        "ORDER BY name;" _
    ), vbNewLine)
End Function

Public Function FunctionsTableNamedParams() As String
    FunctionsTableNamedParams = Join(Array( _
        "SELECT * FROM functions", _
        "WHERE ([builtin] = @builtinY OR [builtin] = @builtinN AND [flags] = @flags) AND", _
        "      ([enc] = @enc AND [narg] >= @narg AND [type] = @type)", _
        "ORDER BY name;" _
    ), vbNewLine)
End Function

Public Function FunctionsFilteredNamedParamsArray() As Variant
    FunctionsFilteredNamedParamsArray = Array( _
        1, _
        0, _
        0, _
        "utf8", _
        0, _
        "s" _
    )
End Function

Public Function FunctionsFilteredNamedParamsDict() As Scripting.Dictionary
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
    Set FunctionsFilteredNamedParamsDict = QueryParams
End Function

Public Function ITRBTableCREATE() As String
    ITRBTableCREATE = Join(Array( _
        "CREATE TABLE itrb(", _
        "    id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,", _
        "    xi INTEGER,", _
        "    xt TEXT COLLATE NOCASE,", _
        "    xr REAL NOT NULL,", _
        "    xb BLOB", _
        ");" _
    ), vbNewLine)
End Function

Public Function ITRBTableINSERT() As String
    ITRBTableINSERT = Join(Array( _
        "INSERT INTO itrb(id,   xi,    xt,  xr,                  xb) ", _
        "VALUES          ( 1,   10, 'AAA', 3.1, X'410A0D0942434445'),", _
        "                ( 2,   20,  NULL, 1.3, X'30310A0D09323334'),", _
        "                ( 3, NULL, 'AAA', 7.2,                NULL),", _
        "                ( 4,   27, 'DDD', 4.3, X'410A0D0942434445'),", _
        "                ( 5, NULL,  NULL, 3.8, X'30310A0D32093334');" _
    ), vbNewLine)
End Function

Public Function ITRBTableSELECT() As String
    ITRBTableSELECT = Join(Array( _
        "SELECT rowid, * FROM itrb;" _
    ), vbNewLine)
End Function
