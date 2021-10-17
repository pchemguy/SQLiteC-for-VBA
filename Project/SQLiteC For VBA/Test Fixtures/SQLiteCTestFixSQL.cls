VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLiteCTestFixSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Provides SQL snippets for tests"
'@Folder "SQLiteC For VBA.Test Fixtures"
'@ModuleDescription "Provides SQL snippets for tests"
'@PredeclaredId
Option Explicit


Public Function CREATETableIRBNT() As String
    CREATETableIRBNT = Join(Array( _
        "CREATE TABLE t1(", _
        "    xi INTEGER,", _
        "    xr REAL,", _
        "    xb BLOB,", _
        "    xn NUMERIC,", _
        "    xt Text", _
        ");" _
    ), vbNewLine)
End Function

Public Function INSERTValuesIRBNT() As String
    INSERTValuesIRBNT = Join(Array( _
        "INSERT INTO t1(rowid, xi, xr, xb, xn,    xt) ", _
        "VALUES        (    1, 10, .1,  1,  7, 'AAA'),", _
        "              (    2, 20, .3,  1, .2, 'BBB'),", _
        "              (    3,  8, .2,  1, 10, 'AAA'),", _
        "              (    4, 27, .3,  1, .9, 'DDD'),", _
        "              (    5,  2, .8,  1,  1, 'BBB');" _
    ), vbNewLine)
End Function

Public Function SELECTt1() As String
    SELECTt1 = "SELECT * FROM t1"
End Function

Public Function CREATETableINSERTValuesIRBNT() As String
    CREATETableINSERTValuesIRBNT = _
        CREATETableIRBNT & vbNewLine & INSERTValuesIRBNT
End Function

Public Function SELECTSQLiteVersion() As String
    SELECTSQLiteVersion = "SELECT sqlite_version()"
End Function

Public Function SELECTDbPath() As String
    SELECTDbPath = Join(Array( _
        "SELECT file FROM pragma_database_list;" _
    ), vbNewLine)
End Function

Public Function SELECTCollations() As String
    SELECTCollations = Join(Array( _
        "SELECT name FROM pragma_collation_list AS collations ORDER BY name;" _
    ), vbNewLine)
End Function

Public Function SELECTFunctionsTable() As String
    SELECTFunctionsTable = Join(Array( _
        "WITH functions AS (SELECT rowid, * FROM pragma_function_list)", _
        "SELECT * FROM functions ORDER BY name;" _
    ), vbNewLine)
End Function

Public Function SELECTFunctionsTableWHERE() As String
    SELECTFunctionsTableWHERE = Join(Array( _
        "WITH functions AS (SELECT rowid, * FROM pragma_function_list)", _
        "SELECT * FROM functions", _
        "WHERE ([builtin] = 1 OR [builtin] = 0 AND [flags] = 0) AND", _
        "      ([enc] = 'utf8' AND [narg] >= 0 AND [type] = 's')", _
        "ORDER BY name;" _
    ), vbNewLine)
End Function

Public Function SELECTFunctionsTableWHERENamedParams() As String
    SELECTFunctionsTableWHERENamedParams = Join(Array( _
        "WITH functions AS (SELECT rowid, * FROM pragma_function_list)", _
        "SELECT * FROM functions", _
        "WHERE ([builtin] = @builtinY OR [builtin] = @builtinN AND [flags] = @flags) AND", _
        "      ([enc] = @enc AND [narg] >= @narg AND [type] = @type)", _
        "ORDER BY name;" _
    ), vbNewLine)
End Function

Public Function SELECTFunctionsTableWHEREParamsArray() As Variant
    SELECTFunctionsTableWHEREParamsArray = Array( _
        1, _
        0, _
        0, _
        "utf8", _
        0, _
        "s" _
    )
End Function

Public Function SELECTFunctionsTableWHEREParamsDict() As Scripting.Dictionary
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
    Set SELECTFunctionsTableWHEREParamsDict = QueryParams
End Function
