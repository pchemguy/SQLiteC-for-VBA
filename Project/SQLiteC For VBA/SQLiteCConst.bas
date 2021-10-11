Attribute VB_Name = "SQLiteCConst"
'@Folder "SQLiteC For VBA"
'@IgnoreModule IndexedDefaultMemberAccess
Option Explicit

#If WIN64 Then
    Public Const ARCH As String = "x64"
#Else
    Public Const ARCH As String = "x32"
#End If
Public Const LITE_LIB As String = "SQLiteCforVBA"
Public Const PATH_SEP As String = "\"
Public Const LITE_RPREFIX As String = "Library" & PATH_SEP & LITE_LIB & PATH_SEP

Public Const SQLITE_STATIC      As Long = 0
Public Const SQLITE_TRANSIENT   As Long = -1

Public Enum SQLiteOpenFlags
    SQLITE_OPEN_READONLY = &H1&
    SQLITE_OPEN_READWRITE = &H2&
    SQLITE_OPEN_CREATE = &H4&
    SQLITE_OPEN_DELETEONCLOSE = &H8&
    SQLITE_OPEN_EXCLUSIVE = &H10&
    SQLITE_OPEN_AUTOPROXY = &H20&
    SQLITE_OPEN_URI = &H40&
    SQLITE_OPEN_MEMORY = &H80&
    SQLITE_OPEN_MAIN_DB = &H100&
    SQLITE_OPEN_TEMP_DB = &H200&
    SQLITE_OPEN_TRANSIENT_DB = &H400&
    SQLITE_OPEN_MAIN_JOURNAL = &H800&
    SQLITE_OPEN_TEMP_JOURNAL = &H1000&
    SQLITE_OPEN_SUBJOURNAL = &H2000&
    SQLITE_OPEN_SUPER_JOURNAL = &H4000&
    SQLITE_OPEN_NOMUTEX = &H8000&
    SQLITE_OPEN_FULLMUTEX = &H10000
    SQLITE_OPEN_SHAREDCACHE = &H20000
    SQLITE_OPEN_PRIVATECACHE = &H40000
    SQLITE_OPEN_WAL = &H80000
    SQLITE_OPEN_NOFOLLOW = &H1000000
    SQLITE_OPEN_DEFAULT = SQLITE_OPEN_READWRITE Or SQLITE_OPEN_CREATE
End Enum

Public Enum SQLiteTypes
    SQLITE_INTEGER = 1&
    SQLITE_FLOAT = 2&
    SQLITE_TEXT = 3&
    SQLITE_BLOB = 4&
    SQLITE_NULL = 5&
End Enum


Public Function SQLiteTypeName(ByVal SQLiteType As SQLiteTypes) As String
    SQLiteTypeName = Array("INTEGER", "FLOAT", "TEXT", "BLOB", "NULL")(SQLiteType - 1)
End Function
