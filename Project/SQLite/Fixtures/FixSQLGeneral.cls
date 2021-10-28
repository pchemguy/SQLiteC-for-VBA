VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FixSQLGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Assorted snippets."
'@Folder "SQLite.Fixtures"
'@ModuleDescription "Assorted snippets."
'@PredeclaredId
'@IgnoreModule ProcedureNotUsed
Option Explicit


'''' Produces 'SELECT <LiteralValue>;' statement with literal from argument.
'''' String literal is quoted, and if now argument is provided, "@Literal"
'''' named parameter is used in place of the value.
''''
'@Description "Produces 'SELECT <LiteralValue>;' statement."
Public Function SelectLiteralAtParam(Optional ByVal Literal As Variant = Empty) As String
Attribute SelectLiteralAtParam.VB_Description = "Produces 'SELECT <LiteralValue>;' statement."
    Dim LiteralParam As String
    If VarType(Literal) = Empty Then
        LiteralParam = "@Literal"
    ElseIf VarType(Literal) = vbString Then
        LiteralParam = "'" & Literal & "'"
    Else
        LiteralParam = CStr(Literal)
    End If
    SelectLiteralAtParam = "SELECT " & LiteralParam & ";"
End Function

Public Function SQLiteVersionScalar() As String
    SQLiteVersionScalar = "SELECT sqlite_version()"
End Function

'@Description "SQL for retrieving a single column containing attached databases paths."
Public Function AttachedDbPaths() As String
Attribute AttachedDbPaths.VB_Description = "SQL for retrieving a single column containing attached databases paths."
    AttachedDbPaths = Join(Array( _
        "SELECT file FROM pragma_database_list;" _
    ), vbNewLine)
End Function

'@Description "SQL for retrieving a single column with SQLite collation names."
Public Function Collations() As String
Attribute Collations.VB_Description = "SQL for retrieving a single column with SQLite collation names."
    Collations = "SELECT name FROM pragma_collation_list AS collations ORDER BY name;"
End Function

Public Function CreateBasicTable() As String
    CreateBasicTable = "CREATE TABLE t(x)"
End Function

'@Description "Wraps a non-CTE SELECT query to yield SQL returning RecordCount for the original query."
Public Function CountSelectNoCTE(ByVal SQLQuery As String) As String
Attribute CountSelectNoCTE.VB_Description = "Wraps a non-CTE SELECT query to yield SQL returning RecordCount for the original query."
    CountSelectNoCTE = _
        "WITH" & vbNewLine & "  subset AS (" & vbNewLine & _
        "    " & Replace(Replace(SQLQuery, ";", vbNullString), vbNewLine, vbNewLine & "    ") & _
        vbNewLine & ")" & vbNewLine & _
        "SELECT count(*) AS counter FROM subset;"
End Function

'@Description "Wraps a SELECT query to yield SQL returning RecordCount for the original query."
Public Function CountSelect(ByVal SQLQuery As String) As String
Attribute CountSelect.VB_Description = "Wraps a SELECT query to yield SQL returning RecordCount for the original query."
    CountSelect = "SELECT count(*) AS counter FROM " & _
                  "(" & Replace(SQLQuery, ";", vbNullString) & ");"
End Function
