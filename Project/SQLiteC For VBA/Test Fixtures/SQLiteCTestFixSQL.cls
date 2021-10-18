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


Public Function CREATETableITRB() As String
    CREATETableITRB = Join(Array( _
        "CREATE TABLE t1(", _
        "    id INTEGER NOT NULL,", _
        "    xi INTEGER,", _
        "    xt TEXT COLLATE NOCASE,", _
        "    xr REAL NOT NULL,", _
        "    xb BLOB,", _
        "    PRIMARY KEY(""id"" AUTOINCREMENT)", _
        ");" _
    ), vbNewLine)
End Function

Public Function INSERTValuesITRB() As String
    INSERTValuesITRB = Join(Array( _
        "INSERT INTO t1(id,   xi,    xt,  xr,                  xb) ", _
        "VALUES        ( 1,   10, 'AAA', 3.1, X'410A0D0942434445'),", _
        "              ( 2,   20,  NULL, 1.3, X'30310A0D09323334'),", _
        "              ( 3, NULL, 'AAA', 7.2,                NULL),", _
        "              ( 4,   27, 'DDD', 4.3, X'410A0D0942434445'),", _
        "              ( 5, NULL,  NULL, 3.8, X'30310A0D32093334');" _
    ), vbNewLine)
End Function

Public Function SELECTTestTable() As String
    SELECTTestTable = Join(Array( _
        "SELECT rowid, * FROM t1;" _
    ), vbNewLine)
End Function

Public Function CREATETableINSERTValuesIRBNT() As String
    CREATETableINSERTValuesIRBNT = _
        CREATETableITRB & vbNewLine & INSERTValuesITRB
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

Public Function SELECTFunctionsNamedParamsAnon() As String
    SELECTFunctionsNamedParamsAnon = Join(Array( _
        "WITH functions AS (SELECT rowid, * FROM pragma_function_list)", _
        "SELECT * FROM functions", _
        "WHERE ([builtin] = ? OR [builtin] = ? AND [flags] = ?) AND", _
        "      ([enc] = ? AND [narg] >= ? AND [type] = ?)", _
        "ORDER BY name;" _
    ), vbNewLine)
End Function

Public Function SELECTFunctionsNamedParamsAnonValues() As Variant
    SELECTFunctionsNamedParamsAnonValues = Array( _
        1, _
        0, _
        0, _
        "utf8", _
        0, _
        "s" _
    )
End Function

Public Function SELECTFunctionsNamedParamsNo() As String
    SELECTFunctionsNamedParamsNo = Join(Array( _
        "WITH functions AS (SELECT rowid, * FROM pragma_function_list)", _
        "SELECT * FROM functions", _
        "WHERE ([builtin] = ?3 OR [builtin] = ?4 AND [flags] = ?5) AND", _
        "      ([enc] = ?1 AND [narg] >= ?6 AND [type] = ?2)", _
        "ORDER BY name;" _
    ), vbNewLine)
End Function

Public Function SELECTFunctionsNamedParamsNoValues() As Variant
    SELECTFunctionsNamedParamsNoValues = Array( _
        "utf8", _
        "s", _
        1, _
        0, _
        0, _
        0 _
    )
End Function

Public Function SELECTFunctionsNamedParamsAt() As String
    SELECTFunctionsNamedParamsAt = Join(Array( _
        "WITH functions AS (SELECT rowid, * FROM pragma_function_list)", _
        "SELECT * FROM functions", _
        "WHERE ([builtin] = @builtinY OR [builtin] = @builtinN AND [flags] = @flags) AND", _
        "      ([enc] = @enc AND [narg] >= @narg AND [type] = @type)", _
        "ORDER BY name;" _
    ), vbNewLine)
End Function

Public Function SELECTFunctionsNamedParamsAtValues() As Scripting.Dictionary
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
    Set SELECTFunctionsNamedParamsAtValues = QueryParams
End Function

Public Function SELECTFunctionsNamedParamsS() As String
    SELECTFunctionsNamedParamsS = Join(Array( _
        "WITH functions AS (SELECT rowid, * FROM pragma_function_list)", _
        "SELECT * FROM functions", _
        "WHERE ([builtin] = $builtinY OR [builtin] = $builtinN AND [flags] = $flags) AND", _
        "      ([enc] = $enc AND [narg] >= $narg AND [type] = $type)", _
        "ORDER BY name;" _
    ), vbNewLine)
End Function

Public Function SELECTFunctionsNamedParamsSValues() As Scripting.Dictionary
    Dim QueryParams As Scripting.Dictionary
    Set QueryParams = New Scripting.Dictionary
    With QueryParams
        .CompareMode = TextCompare
        .Item("$builtinY") = 1
        .Item("$builtinN") = 0
        .Item("$flags") = 0
        .Item("$enc") = "utf8"
        .Item("$narg") = 0
        .Item("$type") = "s"
    End With
    Set SELECTFunctionsNamedParamsSValues = QueryParams
End Function

Public Function SELECTFunctionsNamedParamsCOL() As String
    SELECTFunctionsNamedParamsCOL = Join(Array( _
        "WITH functions AS (SELECT rowid, * FROM pragma_function_list)", _
        "SELECT * FROM functions", _
        "WHERE ([builtin] = :builtinY OR [builtin] = :builtinN AND [flags] = :flags) AND", _
        "      ([enc] = :enc AND [narg] >= :narg AND [type] = :type)", _
        "ORDER BY name;" _
    ), vbNewLine)
End Function

Public Function SELECTFunctionsNamedParamsCOLValues() As Scripting.Dictionary
    Dim QueryParams As Scripting.Dictionary
    Set QueryParams = New Scripting.Dictionary
    With QueryParams
        .CompareMode = TextCompare
        .Item(":builtinY") = 1
        .Item(":builtinN") = 0
        .Item(":flags") = 0
        .Item(":enc") = "utf8"
        .Item(":narg") = 0
        .Item(":type") = "s"
    End With
    Set SELECTFunctionsNamedParamsCOLValues = QueryParams
End Function
