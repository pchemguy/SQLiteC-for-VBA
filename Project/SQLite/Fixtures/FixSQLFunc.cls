VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FixSQLFunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "SQL snippets focused on the function list provided by the function_list pragma."
'@Folder "SQLite.Fixtures"
'@ModuleDescription "SQL snippets focused on the function list provided by the function_list pragma."
'@PredeclaredId
'@IgnoreModule ProcedureNotUsed
Option Explicit


Public Function Create() As String
    Create = Join(Array( _
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

Public Function SelectRowid() As String
    SelectRowid = "SELECT rowid, * FROM functions ORDER BY name"
End Function

Public Function SelectNoRowid() As String
    SelectNoRowid = "SELECT * FROM functions ORDER BY name"
End Function

Public Function SelectPragmaRowid() As String
    SelectPragmaRowid = "SELECT rowid, * FROM pragma_function_list AS functions"
End Function

Public Function SelectPragmaNoRowid() As String
    SelectPragmaNoRowid = "SELECT * FROM pragma_function_list AS functions"
End Function

Public Function InsertData() As String
    InsertData = Join(Array( _
        "INSERT INTO functions ", _
        SelectPragmaNoRowid _
    ), vbNewLine)
End Function

Public Function CreateWithData() As String
    CreateWithData = Join(Array( _
        "DROP TABLE IF EXISTS functions;", _
        Create, _
        "INSERT INTO functions ", _
        SelectPragmaNoRowid _
    ), vbNewLine)
End Function

Public Function SelectFilteredPlain() As String
    SelectFilteredPlain = Join(Array( _
        SelectPragmaNoRowid, _
        "WHERE ([builtin] = 1 OR [builtin] = 0 AND [flags] = 0) AND", _
        "      ([enc] = 'utf8' AND [narg] >= 0 AND [type] = 's')", _
        "ORDER BY name;" _
    ), vbNewLine)
End Function

Public Function SelectFilteredParamAnon() As String
    SelectFilteredParamAnon = Join(Array( _
        SelectPragmaNoRowid, _
        "WHERE ([builtin] = ? OR [builtin] = ? AND [flags] = ?) AND", _
        "      ([enc] = ? AND [narg] >= ? AND [type] = ?)", _
        "ORDER BY name;" _
    ), vbNewLine)
End Function

Public Function SelectFilteredParamAnonValues() As Variant
    SelectFilteredParamAnonValues = Array(1, 0, 0, "utf8", 0, "s")
End Function

Public Function SelectFilteredParamNo() As String
    SelectFilteredParamNo = Join(Array( _
        SelectPragmaNoRowid, _
        "WHERE ([builtin] = ?3 OR [builtin] = ?4 AND [flags] = ?5) AND", _
        "      ([enc] = ?1 AND [narg] >= ?6 AND [type] = ?2)", _
        "ORDER BY name;" _
    ), vbNewLine)
End Function

Public Function SelectFilteredParamNoValues() As Variant
    SelectFilteredParamNoValues = Array("utf8", "s", 1, 0, 0, 0)
End Function

Public Function SelectFilteredParamName(Optional ByVal Prefix As String = "@") As String
    Dim Query As String
    Query = Join(Array( _
        SelectPragmaNoRowid, _
        "WHERE ([builtin] = @builtinY OR [builtin] = @builtinN AND [flags] = @flags) AND", _
        "      ([enc] = @enc AND [narg] >= @narg AND [type] = @type)", _
        "ORDER BY name;" _
    ), vbNewLine)
    SelectFilteredParamName = Replace(Query, "@", Prefix)
End Function

Public Function SelectFilteredParamNameValues(Optional ByVal Prefix As String = "@" _
                                             ) As Scripting.Dictionary
    Dim QueryParams As Scripting.Dictionary
    Set QueryParams = New Scripting.Dictionary
    With QueryParams
        .CompareMode = TextCompare
        .Item(Prefix & "builtinY") = 1
        .Item(Prefix & "builtinN") = 0
        .Item(Prefix & "flags") = 0
        .Item(Prefix & "enc") = "utf8"
        .Item(Prefix & "narg") = 0
        .Item(Prefix & "type") = "s"
    End With
    Set SelectFilteredParamNameValues = QueryParams
End Function

Public Function SelectMinMaxSubstrTrimPlain() As String
    SelectMinMaxSubstrTrimPlain = Join(Array( _
        "SELECT name, count(*) AS rept", _
        "FROM functions", _
        "WHERE name like 'substr%' OR", _
        "      name like 'm%' AND length(name) = 3 OR", _
        "      name like '%trim' OR", _
        "      name like 'l%' AND length(name) < 5", _
        "GROUP BY name", _
        "HAVING rept > 1", _
        "ORDER BY name DESC;" _
    ), vbNewLine)
End Function

Public Function SelectMinMaxSubstrTrimParamName() As String
    SelectMinMaxSubstrTrimParamName = Join(Array( _
        "SELECT name, count(*) AS rept", _
        "FROM functions", _
        "WHERE name like @substr_pattern OR", _
        "      name like @min_max_pattern AND length(name) = @min_max_len OR", _
        "      name like @trim_pattern OR", _
        "      name like @LogPattern AND length(name) < @LogLen", _
        "GROUP BY name", _
        "HAVING Rept > @Rept", _
        "ORDER BY name DESC;" _
    ), vbNewLine)
End Function

Public Function SelectMinMaxSubstrTrimParamNameValues() As Scripting.Dictionary
    Dim QueryParams As Scripting.Dictionary
    Set QueryParams = New Scripting.Dictionary
    With QueryParams
        .CompareMode = TextCompare
        .Item("@substr_pattern") = "substr%"
        .Item("@min_max_pattern") = "m%"
        .Item("@min_max_len") = 3
        .Item("@trim_pattern") = "%trim"
        .Item("@LogPattern") = "l%"
        .Item("@LogLen") = 5
        .Item("@Rept") = 1
    End With
    Set SelectMinMaxSubstrTrimParamNameValues = QueryParams
End Function

Public Function SelectMinMaxSubstrTrimParamAnon() As String
    SelectMinMaxSubstrTrimParamAnon = Join(Array( _
        "SELECT name, count(*) AS rept", _
        "FROM functions", _
        "WHERE name like ? OR", _
        "      name like ? AND length(name) = ? OR", _
        "      name like ? OR", _
        "      name like ? AND length(name) < ?", _
        "GROUP BY name", _
        "HAVING Rept > ?", _
        "ORDER BY name DESC;" _
    ), vbNewLine)
End Function

Public Function SelectMinMaxSubstrTrimParamAnonValues() As Variant
    SelectMinMaxSubstrTrimParamAnonValues = Array("substr%", "m%", 3, "%trim", "l%", 5, 1)
End Function
