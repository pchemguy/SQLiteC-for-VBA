Attribute VB_Name = "SQLiteCConst"
'@Folder "SQLiteC For VBA"
'@IgnoreModule IndexedDefaultMemberAccess

''''======================================================================''''
'''' Acknowledgement
'''' Some code from the https://github.com/govert/SQLiteForExcel project.
''''======================================================================''''

Option Explicit

#If WIN64 Then
    Public Const ARCH As String = "x64"
#Else
    Public Const ARCH As String = "x32"
#End If

#If VBA7 <> True Then
    Public Const vbLongLong As Long = 20&
#End If

Public Const KeyAlreadyExistsErr As Long = 457
Public Const OutOfMemoryErr As Long = 7&
Public Const ConnectionNotOpenedErr As Long = vbObjectError + 3000
Public Const StatementNotPreparedErr As Long = vbObjectError + 3001

Public Enum SQLiteType
    SQLITE_NONE = 0&
    SQLITE_INTEGER = 1&
    SQLITE_FLOAT = 2&
    SQLITE_TEXT = 3&
    SQLITE_BLOB = 4&
    SQLITE_NULL = 5&
End Enum

'''' Reference: https://www.sqlite.org/datatype3.html
'''' Highest priority at the top.
Public Enum SQLiteTypeAffinity
    SQLITE_AFF_INTEGER = &H44    ' /* 'D': "%INT%" */
    SQLITE_AFF_TEXT = &H42       ' /* 'B': "%CHAR%" | "%CLOB%" | "%TEXT%" */
    SQLITE_AFF_BLOB = &H41       ' /* 'A': "%BLOB%" */
    SQLITE_AFF_REAL = &H45       ' /* 'E': "%REAL%" | "%FLOA%" | "%DOUB%" */
    SQLITE_AFF_NUMERIC = &H43    ' /* 'C': "%%" */
    SQLITE_AFF_NONE = &H40       ' /* '@': */
End Enum

'''' ====================================================================== ''''
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

Public Type SQLiteCColumnMeta
    Name As String
    '''' .ColumnIndex must be set by the caller; set .Initialized = -1 flag to confirm
    ColumnIndex As Long
    Initialized As Long
    DbName As String
    TableName As String
    OriginName As String
    DataType As SQLiteType
    DeclaredTypeC As String
    Affinity As SQLiteTypeAffinity
    AffinityType As SQLiteType
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
