VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FixSQLITRB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Provides SQL snippets related to the ITRB table."
'@Folder "SQLite.Fixtures"
'@ModuleDescription "Provides SQLite SQL snippets related to the ITRB table."
'@PredeclaredId
Option Explicit


'''' Reference: https://www.sqlite.org/withoutrowid.html
'''' Reference: https://www.sqlite.org/lang_createtable.html#rowid
''''
'''' Creates a table with "id" aliased to ROWID
Public Function Create() As String
    Create = Join(Array( _
        "CREATE TABLE itrb(", _
        "    id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,", _
        "    xi INTEGER,", _
        "    xt TEXT COLLATE NOCASE,", _
        "    xr REAL NOT NULL,", _
        "    xb BLOB", _
        ");" _
    ), vbNewLine)
End Function

'''' Reference: https://www.sqlite.org/withoutrowid.html
'''' Reference: https://www.sqlite.org/lang_createtable.html#rowid
''''
'''' Creates a table with separate ROWID
Public Function CreateRowid() As String
    CreateRowid = Join(Array( _
        "CREATE TABLE itrb(", _
        "    id INT PRIMARY KEY NOT NULL,", _
        "    xi INTEGER,", _
        "    xt TEXT COLLATE NOCASE,", _
        "    xr REAL NOT NULL,", _
        "    xb BLOB", _
        ");" _
    ), vbNewLine)
End Function

Public Function Drop() As String
    Drop = "DROP TABLE IF EXISTS itrb;"
End Function

Public Function SelectRowid() As String
    SelectRowid = "SELECT rowid, * FROM itrb;"
End Function

Public Function SelectNoRowid() As String
    SelectNoRowid = "SELECT * FROM itrb;"
End Function

Public Function InsertPlain() As String
    InsertPlain = Join(Array( _
        "INSERT INTO itrb(id,   xi,    xt,  xr,                  xb) ", _
        "VALUES          ( 0,   10, 'AAA', 3.1, X'410A0D0942434445'),", _
        "                ( 6,   20,  NULL, 1.3, X'30310A0D09323334'),", _
        "                ( 7, NULL, 'AAA', 7.2,                NULL),", _
        "                ( 8,   27, 'DDD', 4.3, X'410A0D0942434445'),", _
        "                ( 9, NULL,  NULL, 3.8, X'30310A0D32093334');" _
    ), vbNewLine)
End Function

Public Function CreateWithValues() As String
    CreateWithValues = Create & vbNewLine & InsertPlain
End Function

Public Function CreateRowidWithValues() As String
    CreateRowidWithValues = CreateRowid & vbNewLine & InsertPlain
End Function

Public Function InsertParamName() As String
    InsertParamName = "INSERT INTO itrb VALUES ($id, $xi, $xt, $xr, $xb);"
End Function

Public Function InsertParamNameNames() As Variant
    InsertParamNameNames = Array("$id", "$xi", "$xt", "$xr", "$xb")
End Function

Public Function InsertParamNameValueSets() As Variant
    Const NonPrn As String = vbCr & vbLf & vbTab
    InsertParamNameValueSets = Array( _
        Array(1, 10, "AAA", 3.1, FixUtils.ByteArray("A" & NonPrn & "BCDE")), _
        Array(2, 20, Null, 1.3, FixUtils.ByteArray("01" & NonPrn & "234")), _
        Array(3, Null, "AAA", 7.2, Null), _
        Array(4, 27, "DDD", 4.3, FixUtils.ByteArray("A" & NonPrn & "BCDE")), _
        Array(5, Null, Null, 3.8, FixUtils.ByteArray("01" & vbCr & vbLf & "2" & vbTab & "34")) _
    )
End Function

Public Function UpdatePlain() As String
    UpdatePlain = "UPDATE itrb SET xr = xr * 2 WHERE xt = 'AAA' OR xi > 20.5;"
End Function

Public Function UpdateParamName() As String
    UpdateParamName = "UPDATE itrb SET xr = xr * @xr WHERE xt = @xt OR xi > @xi;"
End Function

Public Function UpdateParamValueDict() As Scripting.Dictionary
    Set UpdateParamValueDict = FixUtils.KeysValuesToDict( _
        Array("@xr", "@xt", "@xi"), _
        Array(2, "AAA", 20.5) _
    )
End Function
