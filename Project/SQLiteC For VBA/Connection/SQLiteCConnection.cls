VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLiteCConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Manages SQLite connectivity"
'@Folder "SQLiteC For VBA.Connection"
'@ModuleDescription "Manages SQLite connectivity"
'@PredeclaredId
'@Exposed

''''======================================================================''''
'''' Connection
'''' https://www.sqlite.org/c3ref/open.html
'''' https://www.sqlite.org/c3ref/close.html
'''' https://www.sqlite.org/c3ref/errcode.html
'''' https://www.sqlite.org/vfs.html
''''======================================================================''''

Option Explicit

Private Const VERBOUS As Boolean = True

#If VBA7 Then
'''' Engine test, no db is necessary
Private Declare PtrSafe Function sqlite3_libversion Lib "SQLite3" () As LongPtr ' PtrUtf8String
Private Declare PtrSafe Function sqlite3_libversion_number Lib "SQLite3" () As Long
'''' Open/close connection
Private Declare PtrSafe Function sqlite3_open16 Lib "SQLite3" (ByVal pwsFileName As LongPtr, ByRef hDb As LongPtr) As SQLiteErrors
Private Declare PtrSafe Function sqlite3_open_v2 Lib "SQLite3" (ByVal pwsFileName As LongPtr, ByRef hDb As LongPtr, ByVal iFlags As Long, ByVal zVfs As LongPtr) As SQLiteErrors ' PtrDb
Private Declare PtrSafe Function sqlite3_close Lib "SQLite3" (ByVal hDb As LongPtr) As Long
' Database connection error info
Private Declare PtrSafe Function sqlite3_errmsg Lib "SQLite3" (ByVal hDb As LongPtr) As LongPtr ' PtrUtf8String
Private Declare PtrSafe Function sqlite3_errmsg16 Lib "SQLite3" (ByVal hDb As LongPtr) As LongPtr ' PtrUtf16String
Private Declare PtrSafe Function sqlite3_errstr Lib "SQLite3" (ByVal rc As SQLiteErrors) As LongPtr ' PtrUtf8String
Private Declare PtrSafe Function sqlite3_errcode Lib "SQLite3" (ByVal hDb As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_extended_errcode Lib "SQLite3" (ByVal hDb As LongPtr) As Long
' Database connection change counts
Private Declare PtrSafe Function sqlite3_changes Lib "SQLite3" (ByVal hDb As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_total_changes Lib "SQLite3" (ByVal hDb As LongPtr) As Long
#Else
'''' Engine test, no db is necessary
Private Declare Function sqlite3_libversion Lib "SQLite3" () As Long ' PtrUtf8String
Private Declare Function sqlite3_libversion_number Lib "SQLite3" () As Long
'''' Open/close connection
Private Declare Function sqlite3_open16 Lib "SQLite3" (ByVal pwsFileName As Long, ByRef hDb As Long) As SQLiteErrors ' PtrDb
Private Declare Function sqlite3_open_v2 Lib "SQLite3" (ByVal pwsFileName As Long, ByRef hDb As Long, ByVal iFlags As Long, ByVal zVfs As Long) As SQLiteErrors ' PtrDb
Private Declare Function sqlite3_close Lib "SQLite3" (ByVal hDb As Long) As Long
' Database connection error info
Private Declare Function sqlite3_errmsg Lib "SQLite3" (ByVal hDb As Long) As Long ' PtrUtf8String
Private Declare Function sqlite3_errmsg16 Lib "SQLite3" (ByVal hDb As Long) As Long ' PtrUtf16String
Private Declare Function sqlite3_errstr Lib "SQLite3" (ByVal rc As SQLiteErrors) As Long ' PtrUtf8String
Private Declare Function sqlite3_errcode Lib "SQLite3" (ByVal hDb As Long) As Long
Private Declare Function sqlite3_extended_errcode Lib "SQLite3" (ByVal hDb As Long) As Long
' Database connection change counts
Private Declare Function sqlite3_changes Lib "SQLite3" (ByVal hDb As Long) As Long
Private Declare Function sqlite3_total_changes Lib "SQLite3" (ByVal hDb As Long) As Long
#End If

Private Type TSQLiteCConnection
    DbPathName As String
    #If VBA7 Then
        DbHandle As LongPtr
    #Else
        DbHandle As Long
    #End If
    ErrorStatus As SQLiteErrors
    ErrorStatusEx As SQLiteErrors
    ErrorMessage As String
    ErrorString As String
    ErrorName As String
End Type
Private this As TSQLiteCConnection


'@DefaultMember
Public Function Create(ByVal DbPathName As String) As SQLiteCConnection
Attribute Create.VB_UserMemId = 0
    Dim Instance As SQLiteCConnection
    Set Instance = New SQLiteCConnection
    Instance.Init DbPathName
    Set Create = Instance
End Function

Friend Sub Init(ByVal DbPathName As String)
    Dim FilePathName As String
    FilePathName = DbPathName
    '''' TODO: DbPath validation
    ''''       This task should be performed by SQLiteDBVBA.LiteFSCheck class.
    ''''       This functionality will be integrated after the code is merged
    ''''       into the SQLiteDB VBA Library.
    this.DbPathName = FilePathName
End Sub

#If VBA7 Then
Public Property Get DbHandle() As LongPtr
#Else
Public Property Get DbHandle() As Long
#End If
    DbHandle = this.DbHandle
End Property

Public Property Get DbPathName() As String
    DbPathName = this.DbPathName
End Property

Public Property Get ErrorStatus() As SQLiteErrors
    ErrorStatus = this.ErrorStatus
End Property

Public Property Get ErrorStatusName() As String
    ErrorStatusName = SQLiteErrorName(this.ErrorStatus)
End Property

Public Property Get ErrorStatusEx() As SQLiteErrors
    ErrorStatusEx = this.ErrorStatusEx
End Property

Public Property Get ErrorStatusExName() As String
    ErrorStatusExName = SQLiteErrorName(this.ErrorStatusEx)
End Property

Public Property Get ErrorMessage() As String
    ErrorMessage = this.ErrorMessage
End Property

Public Property Get ErrorString() As String
    ErrorString = this.ErrorString
End Property

Public Property Get ErrorName() As String
    ErrorName = this.ErrorName
End Property


Public Function Version(Optional ByVal Numeric As Boolean = True) As Variant
    If Numeric Then
        Version = sqlite3_libversion_number()
    Else
        Version = UTFlib.Utf8PtrToString(sqlite3_libversion())
    End If
End Function


Public Sub OpenDb(Optional ByVal Flags As SQLiteOpenFlags = SQLITE_OPEN_DEFAULT, _
                  Optional ByVal vfsName As String = vbNullString)
    Dim ErrorStatus As SQLiteErrors
    If Flags = SQLITE_OPEN_DEFAULT And Len(vfsName) = 0 Then
        ErrorStatus = sqlite3_open16(StrPtr(this.DbPathName), this.DbHandle)
    Else
        #If VBA7 Then
            Dim DbPathNamePtr As LongPtr
            Dim vfsNamePtr As LongPtr
        #Else
            Dim DbPathNamePtr As Long
            Dim vfsNamePtr As Long
        #End If
        
        '@Ignore UseMeaningfulName
        Dim DbPathNameUTF8() As Byte
        DbPathNameUTF8 = UTFlib.StringToUtf8Bytes(this.DbPathName)
        DbPathNamePtr = VarPtr(DbPathNameUTF8(0))
        
        If Len(vfsName) = 0 Then
            vfsNamePtr = 0
        Else
            '@Ignore UseMeaningfulName
            Dim vfsNameUTF8() As Byte
            vfsNameUTF8 = UTFlib.StringToUtf8Bytes(vfsName)
            vfsNamePtr = VarPtr(vfsNameUTF8(0))
        End If
        
        ErrorStatus = sqlite3_open_v2(DbPathNamePtr, this.DbHandle, Flags, vfsNamePtr)
    End If
    this.ErrorStatus = ErrorStatus
    If ErrorStatus <> SQLITE_OK Then
        GetErrorInfo
        PrintErrorInfo
    End If
    If VERBOUS Then Debug.Print "Open db status: <" & SQLiteErrorName(ErrorStatus) & ">"
End Sub


Public Sub CloseDb()
    Dim ErrorStatus As SQLiteErrors
    ErrorStatus = sqlite3_close(this.DbHandle)
    If ErrorStatus = SQLITE_OK Then this.DbHandle = 0
    this.ErrorStatus = ErrorStatus
    If ErrorStatus <> SQLITE_OK Then
        GetErrorInfo
        PrintErrorInfo
    End If
    If VERBOUS Then Debug.Print "Close db status: <" & SQLiteErrorName(ErrorStatus) & ">"
End Sub


Public Function ChangesCount(Optional ByVal Total As Boolean = False) As Long
    If Total Then
        ChangesCount = sqlite3_total_changes(this.DbHandle)
    Else
        ChangesCount = sqlite3_changes(this.DbHandle)
    End If
End Function


Public Sub GetErrorInfo()
    With this
        .ErrorStatus = sqlite3_errcode(.DbHandle)
        .ErrorStatusEx = sqlite3_extended_errcode(.DbHandle)
        .ErrorMessage = UTFlib.Utf8PtrToString(sqlite3_errmsg(.DbHandle))
        .ErrorName = SQLiteErrorName(.ErrorStatusEx)
        .ErrorString = UTFlib.Utf8PtrToString(sqlite3_errstr(.ErrorStatus))
    End With
End Sub


Public Sub PrintErrorInfo()
    With this
        Debug.Print "========== Error Details ========== " & vbNewLine & _
                    "ErrCode:          " & CStr(.ErrorStatus) & vbNewLine & _
                    "ErrCode Extended: " & CStr(.ErrorStatusEx) & vbNewLine & _
                    "Error Name:       " & .ErrorName & vbNewLine & _
                    "ErrStr:           " & .ErrorString & vbNewLine & _
                    "ErrMsg:           " & .ErrorMessage & vbNewLine
    End With
End Sub
