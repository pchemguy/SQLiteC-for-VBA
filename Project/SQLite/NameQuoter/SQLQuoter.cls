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
'@Folder "SQLite.NameQuoter"
'@ModuleDescription "Handles SQLite name quoting."
'@PredeclaredId
'@IgnoreModule IndexedDefaultMemberAccess
Option Explicit

Private Const SQLITE_N_KEYWORD As Long = 147

#If VBA7 Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As LongPtr)
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal Source As Any, ByVal Length As Long)
#End If

#If VBA7 Then
Private Declare PtrSafe Function sqlite3_keyword_count Lib "SQLite3" () As Long
Private Declare PtrSafe Function sqlite3_keyword_check Lib "SQLite3" (ByVal zName As LongPtr, ByVal nName As Long) As Long
Private Declare PtrSafe Function sqlite3_keyword_name Lib "SQLite3" (ByVal kwi As Long, ByRef pzName As LongPtr, ByRef pnName As Long) As Long
#Else
Private Declare Function sqlite3_keyword_count Lib "SQLite3" () As Long
Private Declare Function sqlite3_keyword_check Lib "SQLite3" (ByVal zName As Long, ByVal nName As Long) As Long
Private Declare Function sqlite3_keyword_name Lib "SQLite3" (ByVal kwi As Long, ByRef pzName As Long, ByRef pnName As Long) As Long
#End If

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
'@Description "Quotes field name."
Public Function QuoteSQLName(ByVal FieldName As String) As String
Attribute QuoteSQLName.VB_Description = "Quotes field name."
    If this.SQLiteKeywords Is Nothing Then Init
    Dim Buffer() As Byte
    Buffer = FieldName
    Dim NameLen As Long
    NameLen = Len(FieldName)
    If Buffer(0) = AscB("""") And Buffer(NameLen * 2 - 2) = AscB("""") Or _
       Buffer(0) = AscB("[") And Buffer(NameLen * 2 - 2) = AscB("]") Then
        NameLen = NameLen - 2
        Buffer = Mid$(FieldName, 2, NameLen)
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
Public Function CStrSQL(ByVal VariantVBA As Variant) As String
Attribute CStrSQL.VB_Description = "Converts a VBA Variant to SQL textual representation."
    Select Case VarType(VariantVBA)
        Case vbString
            CStrSQL = SQLlib.QuoteSQLString(VariantVBA)
        Case vbLong, vbDouble, vbInteger, vbSingle, _
             vbCurrency, vbByte, vbLongLong
            CStrSQL = CStr(VariantVBA)
        Case vbEmpty, vbNull, vbError
            CStrSQL = "NULL"
        Case vbBoolean
            CStrSQL = SQLlib.QuoteSQLString(CStr(VariantVBA))
        Case vbDate
            CStrSQL = Format$(VariantVBA, "YYYY-MM-DDTHH:MM:SS")
    End Select
End Function

Friend Function KeywordCount() As Long
    KeywordCount = sqlite3_keyword_count()
End Function

Friend Function NthKeyword(ByVal KeywordIndex As Long) As String
    #If VBA7 Then
        Dim KeywordPtr As LongPtr
    #Else
        Dim KeywordPtr As Long
    #End If
    KeywordPtr = 0
    Dim KeywordLength As Long
    KeywordLength = 0
    Dim ResultCode As SQLiteResultCodes
    ResultCode = sqlite3_keyword_name(KeywordIndex, KeywordPtr, KeywordLength)
    If ResultCode <> SQLITE_OK Or KeywordPtr = 0 Or KeywordLength = 0 Then
        NthKeyword = vbNullString
    Else
        Dim KeywordBuffer() As Byte
        ReDim KeywordBuffer(0 To KeywordLength - 1)
        CopyMemory KeywordBuffer(0), KeywordPtr, KeywordLength
        NthKeyword = StrConv(KeywordBuffer, vbUnicode)
    End If
End Function

Friend Sub InitKeywordMap()
    If this.SQLiteKeywords Is Nothing Then
        Set this.SQLiteKeywords = New Scripting.Dictionary
        this.SQLiteKeywords.CompareMode = TextCompare
        CachedKeywordsSet
    End If
    If KeywordCount() > SQLITE_N_KEYWORD Then
        Dim KeywordIndex As Long
        KeywordIndex = 0
        Dim Keyword As String
        Keyword = NthKeyword(KeywordIndex)
        Do While Keyword <> vbNullString
            this.SQLiteKeywords(Keyword) = KeywordIndex
            KeywordIndex = KeywordIndex + 1
            Keyword = NthKeyword(KeywordIndex)
        Loop
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
    For KeywordIndex = 0 To SQLITE_N_KEYWORD - 1
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
