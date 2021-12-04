Attribute VB_Name = "SQLiteCConst"
'@Folder "SQLite.C"
'@IgnoreModule IndexedDefaultMemberAccess

''''======================================================================''''
'''' Acknowledgement
'''' Some code from the https://github.com/govert/SQLiteForExcel project.
''''======================================================================''''

Option Explicit

Public Enum SQLiteDataType
    SQLITE_NONE = 0&
    SQLITE_INTEGER = 1&
    SQLITE_FLOAT = 2&
    SQLITE_TEXT = 3&
    SQLITE_BLOB = 4&
    SQLITE_NULL = 5&
End Enum

'''' Reference: https://www.sqlite.org/datatype3.html
''''
'''' ---------------- Mapping SQLiteTypeAffinity -> SQLiteType ------------ ''''
'''' SQLITE_AFF_BLOB -> SQLITE_BLOB
'''' SQLITE_AFF_TEXT -> SQLITE_TEXT
'''' SQLITE_AFF_NUMERIC -> SQLITE_TEXT
'''' SQLITE_AFF_INTEGER -> SQLITE_INTEGER
'''' SQLITE_AFF_REAL -> SQLITE_FLOAT
''''
'''' MappedType = Array(SQLITE_BLOB, SQLITE_TEXT, SQLITE_TEXT, _
''''                    SQLITE_INTEGER, SQLITE_FLOAT)(ColumnAffinity - &H41)
'''' ---------------------------------------------------------------------- ''''
''''
'''' Highest priority at the top.
Public Enum SQLiteTypeAffinity
    SQLITE_AFF_INTEGER = &H44    ' /* 'D': "%INT%" */
    SQLITE_AFF_TEXT = &H42       ' /* 'B': "%CHAR%" | "%CLOB%" | "%TEXT%" */
    SQLITE_AFF_BLOB = &H41       ' /* 'A': "%BLOB%" */
    SQLITE_AFF_REAL = &H45       ' /* 'E': "%REAL%" | "%FLOA%" | "%DOUB%" */
    SQLITE_AFF_NUMERIC = &H43    ' /* 'C': "%%" */
    SQLITE_AFF_NONE = &H40       ' /* '@': */
End Enum

Public Enum SQLiteCTextEncoding
    SQLITE_ENCODING_UTF8 = 1&
    SQLITE_ENCODING_UTF16LE = 2&
    SQLITE_ENCODING_UTF16BE = 3&
End Enum

Public Enum SQLiteCFileFormat
    SQLITE_FORMAT_LEGACY = 1&
    SQLITE_FORMAT_WAL = 2&
End Enum

Public Type SQLiteCErr
    ErrorCode As SQLiteResultCodes
    ErrorCodeName As String
    ErrorCodeEx As SQLiteResultCodes
    ErrorCodeExName As String
    ErrorName As String             ' Alias to ErrorCodeExName
    ErrorMessage As String
    ErrorString As String
End Type

Public Type SQLiteCColumnMeta
    Name As String
    '''' .ColumnIndex must be set by the caller; set .Initialized = -1 flag to confirm
    ColumnIndex As Long
    Initialized As Long
    DbName As String
    TableName As String
    OriginName As String
    DataType As SQLiteDataType
    TableMeta As Boolean '''' Set to True if table meta is available
    DeclaredTypeC As String
    Affinity As SQLiteTypeAffinity
    AffinityType As SQLiteDataType
    DeclaredTypeT As String
    Collation As String
    NotNull As Boolean
    PrimaryKey As Boolean
    AutoIncrement As Boolean
    AdoType As ADODB.DataTypeEnum
    AdoAttr As ADODB.FieldAttributeEnum
    AdoSize As Long
    RowId As Boolean
End Type

'''' https://sqlite.org/fileformat.html
Public Type SQLiteCHeaderData
    MagicHeaderString As String * 16        '''' Bytes  0-15: "SQLite format 3" & vbNullChar
    '@Ignore IntegerDataType
    PageSizeInBytes As Integer              '''' Bytes 16-17: Power of two [2^9, 2^15], 2^0 -> 2^16
    FileFormatWrite As SQLiteCFileFormat    '''' Bytes 18   : 1 - Legacy, 2 - WAL
    FileFormatRead As SQLiteCFileFormat     '''' Bytes 19   : 1 - Legacy, 2 - WAL
    ReservedSpace As Byte                   '''' Bytes 20   : usually 0
    MaxPayload As Byte                      '''' Bytes 21   : must be 64
    MinPayload As Byte                      '''' Bytes 22   : must be 32
    LeafPayload As Byte                     '''' Bytes 23   : must be 32
    ChangeCounter As Long                   '''' Bytes 24-27
    DbFilePageCount As Long                 '''' Bytes 28-31
    FirstFreeListPage As Long               '''' Bytes 32-35
    FreeListPageCount As Long               '''' Bytes 36-39
    SchemaCookie As Long                    '''' Bytes 40-43
    SchemaFormat As Long                    '''' Bytes 44-47: [1,4]
    DefaultPageCacheSize As Long            '''' Bytes 48-51
    LagestBTreeRootPage As Long             '''' Bytes 52-55
    DbTextEncoding As SQLiteCTextEncoding   '''' Bytes 56-59: 1 - UTF8, 2 - UTF16le, 3 - UTF16be
    UserVersion As Long                     '''' Bytes 60-63
    VacuumMode As Long                      '''' Bytes 64-67: True - incremental, False - otherwise.
    AppId As Long                           '''' Bytes 68-71
    Reserved() As Byte                      '''' Bytes 72-91: Must be 0
    VersionValidFor As Long                 '''' Bytes 92-95
    SQLiteVersion As Long                   '''' Bytes 96-99
End Type

'''' https://sqlite.org/fileformat.html
Public Type SQLiteCHeaderPacked
    MagicHeaderString(0 To 15) As Byte      '''' Bytes  0-15: "SQLite format 3" & vbNullChar
    PageSizeInBytes(0 To 1) As Byte         '''' Bytes 16-17: Power of two [2^9, 2^15], 2^0 -> 2^16
    FileFormatWrite As Byte                 '''' Bytes 18   : 1 - Legacy, 2 - WAL
    FileFormatRead As Byte                  '''' Bytes 19   : 1 - Legacy, 2 - WAL
    ReservedSpace As Byte                   '''' Bytes 20   : usually 0
    MaxPayload As Byte                      '''' Bytes 21   : must be 64
    MinPayload As Byte                      '''' Bytes 22   : must be 32
    LeafPayload As Byte                     '''' Bytes 23   : must be 32
    ChangeCounter(0 To 3) As Byte           '''' Bytes 24-27
    DbFilePageCount(0 To 3) As Byte         '''' Bytes 28-31
    FirstFreeListPage(0 To 3) As Byte       '''' Bytes 32-35
    FreeListPageCount(0 To 3) As Byte       '''' Bytes 36-39
    SchemaCookie(0 To 3) As Byte            '''' Bytes 40-43
    SchemaFormat(0 To 3) As Byte            '''' Bytes 44-47: [1,4]
    DefaultPageCacheSize(0 To 3) As Byte    '''' Bytes 48-51
    LagestBTreeRootPage(0 To 3) As Byte     '''' Bytes 52-55
    DbTextEncoding(0 To 3) As Byte          '''' Bytes 56-59: 1 - UTF8, 2 - UTF16le, 3 - UTF16be
    UserVersion(0 To 3) As Byte             '''' Bytes 60-63
    VacuumMode(0 To 3) As Byte              '''' Bytes 64-67: True - incremental, False - otherwise.
    AppId(0 To 3) As Byte                   '''' Bytes 68-71
    Reserved(72 To 91) As Byte              '''' Bytes 72-91: Must be 0
    VersionValidFor(0 To 3) As Byte         '''' Bytes 92-95
    SQLiteVersion(0 To 3) As Byte           '''' Bytes 96-99
End Type
