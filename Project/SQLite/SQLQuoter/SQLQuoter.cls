VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLQuoter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Handles SQLite name quoting."
'@Folder "SQLite.SQLQuoter"
'@ModuleDescription "Handles SQLite name quoting."
'@PredeclaredId
'@IgnoreModule IndexedDefaultMemberAccess
Option Explicit

Private Type TObjectState
    SQLiteKeywords As Scripting.Dictionary
    AlphaChar() As Byte
    AlphaExNumChar() As Byte
End Type
Private this As TObjectState


Friend Sub Init()
    InitKeywordMap
    InitCharMap
End Sub


''''
''''
'''' Examples:
''''   >>> ?SQLQuoter.QuoteSQLString("Abc")
''''   "'Abc'"
''''
''''   >>> ?SQLQuoter.QuoteSQLString("A'b'c")
''''   "'A''b''c'"
''''
'@Description "Quotes SQL string, doubling single quotes."
Public Function QuoteSQLString(ByVal SQLString As String) As String
    QuoteSQLString = "'" & Replace(SQLString, "'", "''") & "'"
End Function


'''' Quotes SQL name if necessary. If second parameter is True, quoting is forced.
'''' If a quoted input is provided (using either double quotes or brackets), it
'''' is returned unchanged (so it must be valid overall, no validation is
'''' attempted here). A multipart name must be provided unquoted.
''''
'''' Examples:
''''   >>> ?SQLQuoter.QuoteSQLName("file_name")
''''   file_name
''''
''''   >>> ?SQLQuoter.QuoteSQLName("file name")
''''   "file name"
''''
''''   >>> ?SQLQuoter.QuoteSQLName("group")
''''   "group"
''''
''''   >>> ?SQLQuoter.QuoteSQLName("main.people.id")
''''   main.people.id
''''
''''   >>> ?SQLQuoter.QuoteSQLName("main.users.group")
''''   "main"."users"."group"
''''
'@Description "Quotes SQL name."
Public Function QuoteSQLName(ByVal FieldName As String, _
                    Optional ByVal Force As Boolean = False) As String
    If Len(FieldName) = 0 Then
        QuoteSQLName = vbNullString
        Exit Function
    End If
    If this.SQLiteKeywords Is Nothing Then Init
    Dim Buffer() As Byte
    Buffer = FieldName
    Dim NameLen As Long
    NameLen = Len(FieldName)
    If Buffer(0) = AscB("""") And Buffer(NameLen * 2 - 2) = AscB("""") Or _
       Buffer(0) = AscB("[") And Buffer(NameLen * 2 - 2) = AscB("]") Then
        QuoteSQLName = FieldName
        Exit Function
    End If
    If Force Then GoTo QUOTE_NAME:
        
    Dim DotCount As Long
    DotCount = Len(FieldName) - Len(Replace(FieldName, ".", vbNullString))
    If DotCount >= 1 And DotCount <= 3 And Not InStr(FieldName, "..") Then
        Dim NameParts As Variant
        NameParts = Split(FieldName, ".")
        Dim PartIndex As Long
        For PartIndex = 0 To UBound(NameParts)
            NameParts(PartIndex) = QuoteSQLName(NameParts(PartIndex))
        Next PartIndex
        Dim Result As String
        Result = Join(NameParts, ".")
        If InStr(Result, """") > 0 Then
            For PartIndex = 0 To UBound(NameParts)
                NameParts(PartIndex) = QuoteSQLName(NameParts(PartIndex), True)
            Next PartIndex
        End If
        QuoteSQLName = Join(NameParts, ".")
        Exit Function
    End If
    
    If this.AlphaChar(Buffer(0)) = 1 Then GoTo QUOTE_NAME:
    Dim CharIndex As Long
    For CharIndex = 2 To NameLen * 2 - 2 Step 2
        If this.AlphaExNumChar(Buffer(CharIndex)) = 1 Then GoTo QUOTE_NAME:
    Next CharIndex
    If this.SQLiteKeywords.Exists(UCase$(CStr(Buffer))) Then GoTo QUOTE_NAME:
    QuoteSQLName = CStr(Buffer)
    Exit Function
QUOTE_NAME:
    QuoteSQLName = """" & Replace(CStr(Buffer), """", """""") & """"
End Function


'''' Examples:
''''   >>> ?SQLQuoter.CStrSQL(5.5)
''''   5.5
''''
''''   >>> ?SQLQuoter.CStrSQL("Text")
''''   'Text'
''''
''''   >>> ?SQLQuoter.CStrSQL(Empty)
''''   NULL
''''
''''   >>> ?SQLQuoter.CStrSQL(Null)
''''   NULL
''''
''''   >>> ?SQLQuoter.CStrSQL(True)
''''   'True'
''''
'@Description "Converts a VBA Variant to SQL textual representation."
Public Function CStrSQL(ByVal VariantValue As Variant) As String
Attribute CStrSQL.VB_Description = "Converts a VBA Variant to SQL textual representation."
    Select Case VarType(VariantValue)
        Case vbString
            CStrSQL = QuoteSQLString(VariantValue)
        Case vbLong, vbDouble, vbInteger, vbSingle, _
             vbCurrency, vbByte, vbLongLong
            CStrSQL = CStr(VariantValue)
        Case vbEmpty, vbNull, vbError
            CStrSQL = "NULL"
        Case vbBoolean
            CStrSQL = QuoteSQLString(CStr(VariantValue))
        Case vbDate
            CStrSQL = Format$(VariantValue, "YYYY-MM-DDTHH:MM:SS")
    End Select
End Function


Friend Sub InitKeywordMap()
    If this.SQLiteKeywords Is Nothing Then
        Set this.SQLiteKeywords = New Scripting.Dictionary
        this.SQLiteKeywords.CompareMode = TextCompare
        CachedKeywordsSet
    End If
End Sub


Friend Sub CachedKeywordsSet()
    Dim CachedKeywords As Variant
    CachedKeywords = Array( _
        "ABORT", "ACTION", "ADD", "AFTER", "ALL", "ALTER", "ALWAYS", "ANALYZE", _
        "AND", "AS", "ASC", "ATTACH", "AUTOINCREMENT", "BEFORE", "BEGIN", _
        "BETWEEN", "BY", "CASCADE", "CASE", "CAST", "CHECK", "COLLATE", _
        "COLUMN", "COMMIT", "CONFLICT", "CONSTRAINT", "CREATE", "CROSS", _
        "CURRENT", "CURRENT_DATE", "CURRENT_TIME", "CURRENT_TIMESTAMP", _
        "DATABASE", "DEFAULT", "DEFERRABLE", "DEFERRED", "DELETE", "DESC", _
        "DETACH", "DISTINCT", "DO", "DROP", "EACH", "ELSE", "END", "ESCAPE", _
        "EXCEPT", "EXCLUDE", "EXCLUSIVE", "EXISTS", "EXPLAIN", "FAIL", "FILTER", _
        "FIRST", "FOLLOWING", "FOR", "FOREIGN", "FROM", "FULL", "GENERATED", _
        "GLOB", "GROUP", "GROUPS", "HAVING", "IF", "IGNORE", "IMMEDIATE", "IN", _
        "INDEX", "INDEXED", "INITIALLY", "INNER", "INSERT", "INSTEAD", _
        "INTERSECT", "INTO", "IS", "ISNULL", "JOIN", "KEY", "LAST", "LEFT", _
        "LIKE", "LIMIT", "MATCH", "MATERIALIZED", "NATURAL", "NO", "NOT", _
        "NOTHING", "NOTNULL", "NULL", "NULLS", "OF", "OFFSET", "ON", "OR", _
        "ORDER", "OTHERS", "OUTER", "OVER", "PARTITION", "PLAN", "PRAGMA", _
        "PRECEDING", "PRIMARY", "QUERY", "RAISE", "RANGE", "RECURSIVE", _
        "REFERENCES", "REGEXP", "REINDEX", "RELEASE", "RENAME", "REPLACE", _
        "RESTRICT", "RETURNING", "RIGHT", "ROLLBACK", "ROW", "ROWS", _
        "SAVEPOINT", "SELECT", "SET", "TABLE", "TEMP", "TEMPORARY", "THEN", _
        "TIES", "TO", "TRANSACTION", "TRIGGER", "UNBOUNDED", "UNION", "UNIQUE", _
        "UPDATE", "USING", "VACUUM", "VALUES", "VIEW", "VIRTUAL", "WHEN", _
        "WHERE", "WINDOW", "WITH", "WITHOUT")
    Dim KeywordIndex As Long
    For KeywordIndex = 0 To UBound(CachedKeywords)
        this.SQLiteKeywords(CachedKeywords(KeywordIndex)) = KeywordIndex
    Next KeywordIndex
End Sub


Friend Sub InitCharMap()
    Dim CharMap(0 To 255) As Byte
    Dim CharCode As Byte
    For CharCode = 0 To AscB("A") - 1
        CharMap(CharCode) = 1
    Next CharCode
    For CharCode = AscB("A") To AscB("Z")
        CharMap(CharCode) = 0
    Next CharCode
    For CharCode = AscB("Z") + 1 To AscB("a") - 1
        CharMap(CharCode) = 1
    Next CharCode
    For CharCode = AscB("a") To AscB("z")
        CharMap(CharCode) = 0
    Next CharCode
    For CharCode = AscB("z") + 1 To 254
        CharMap(CharCode) = 1
    Next CharCode
    CharMap(255) = 1
    this.AlphaChar = CharMap
    For CharCode = AscB("0") To AscB("9")
        CharMap(CharCode) = 0
    Next CharCode
    CharMap(AscB("_")) = 0
    this.AlphaExNumChar = CharMap
End Sub


'''' NOTES
''''
''''   TimeZone info:
''''     REGISTRY:
''''       [HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\TimeZoneInformation]
''''       "ActiveTimeBias"=dword:xxxxxxxx
