VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FixSQLBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Basic SQL statements"
'@Folder "SQLite.Fixtures"
'@ModuleDescription "Basic SQL statements"
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

'''' Examples:
''''   >>>?FixSQLBase.SelectNumericSum()
''''   "SELECT ? + ?;"
''''
''''   >>>?FixSQLBase.SelectNumericSum(1)
''''   "SELECT 1 + ?;"
''''
''''   >>>?FixSQLBase.SelectNumericSum(1, 2)
''''   "SELECT 1 + 2;"
''''
''''   >>>?FixSQLBase.SelectNumericSum(1.5, "A")
''''   "SELECT 1.5 + 0;"
''''
'@Description "Produces 'SELECT <Sum>;' statement."
Public Function SelectNumericSum(Optional ByVal NumA As Variant = Empty, _
                                 Optional ByVal NumB As Variant = Empty) As String
Attribute SelectNumericSum.VB_Description = "Produces 'SELECT <Sum>;' statement."
    Dim TermA As String
    TermA = IIf(IsEmpty(NumA), "?", CStr(Val(NumA)))
    Dim TermB As String
    TermB = IIf(IsEmpty(NumB), "?", CStr(Val(NumB)))
    SelectNumericSum = "SELECT " & TermA & " + " & TermB & ";"
End Function

'''' Examples:
''''   >>>?FixSQLBase.SelectConcat()
''''   "SELECT ? || ?;"
''''
''''   >>>?FixSQLBase.SelectConcat(1)
''''   "SELECT '1' || ?;"
''''
''''   >>>?FixSQLBase.SelectConcat(1, 2)
''''   "SELECT '1' || '2';"
''''
''''   >>>?FixSQLBase.SelectConcat(1.5, "Ab'c")
''''   "SELECT '1.5' || 'Ab''c';"
''''
'@Description "Produces 'SELECT <Concat>;' statement."
Public Function SelectConcat(Optional ByVal TextA As Variant = Empty, _
                             Optional ByVal TextB As Variant = Empty) As String
Attribute SelectConcat.VB_Description = "Produces 'SELECT <Concat>;' statement."
    Dim TermA As String
    TermA = IIf(IsEmpty(TextA), "?", SQLQuoter.QuoteSQLString(CStr(TextA)))
    Dim TermB As String
    TermB = IIf(IsEmpty(TextB), "?", SQLQuoter.QuoteSQLString(CStr(TextB)))
    SelectConcat = "SELECT " & TermA & " || " & TermB & ";"
End Function

Public Function CreateBasicTable() As String
    CreateBasicTable = "CREATE TABLE t(x)"
End Function
