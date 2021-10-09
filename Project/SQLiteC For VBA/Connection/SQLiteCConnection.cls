VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLiteCConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Manages SQLite connectivity"
'@Folder "SQLiteC For VBA.Connection"
'@ModuleDescription "Manages SQLite connectivity"
'@PredeclaredId

''''======================================================================''''
'''' Open/close connection
'''' https://www.sqlite.org/c3ref/open.html
'''' https://www.sqlite.org/vfs.html
''''======================================================================''''

Option Explicit

#If VBA7 Then
'''' Engine test, no db is necessary
Private Declare PtrSafe Function sqlite3_libversion Lib "SQLite3" () As LongPtr ' PtrUtf8String
Private Declare PtrSafe Function sqlite3_libversion_number Lib "SQLite3" () As Long
'''' Open/close connection
Private Declare PtrSafe Function sqlite3_open16 Lib "SQLite3" (ByVal pwsFileName As LongPtr, ByRef hDb As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_open_v2 Lib "SQLite3" (ByVal pwsFileName As LongPtr, ByRef hDb As LongPtr, ByVal iFlags As Long, ByVal zVfs As LongPtr) As Long ' PtrDb
Private Declare PtrSafe Function sqlite3_close Lib "SQLite3" (ByVal hDb As LongPtr) As Long
#Else
'''' Engine test, no db is necessary
Private Declare Function sqlite3_libversion Lib "SQLite3" () As Long ' PtrUtf8String
Private Declare Function sqlite3_libversion_number Lib "SQLite3" () As Long
'''' Open/close connection
Private Declare Function sqlite3_open16 Lib "SQLite3" (ByVal pwsFileName As Long, ByRef hDb As Long) As Long ' PtrDb
Private Declare Function sqlite3_open_v2 Lib "SQLite3" (ByVal pwsFileName As Long, ByRef hDb As Long, ByVal iFlags As Long, ByVal zVfs As Long) As Long ' PtrDb
Private Declare Function sqlite3_close Lib "SQLite3" (ByVal hDb As Long) As Long
#End If

Private Type TSQLiteCConnection
    DbPathName As String
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


Public Function SQLite3Version(Optional ByVal Numeric As Boolean = True) As Variant
    If Numeric Then
        SQLite3Version = sqlite3_libversion_number()
    Else
        SQLite3Version = UTFlib.Utf8PtrToString(sqlite3_libversion())
    End If
End Function


#If WIN64 Then
Public Function SQLite3Open(ByVal FileName As String, ByRef dbHandle As LongPtr) As Long
#Else
Public Function SQLite3Open(ByVal FileName As String, ByRef dbHandle As Long) As Long
#End If
    SQLite3Open = sqlite3_open16(StrPtr(FileName), dbHandle)
End Function


#If WIN64 Then
Public Function SQLite3OpenV2(ByVal FileName As String, ByRef dbHandle As LongPtr, ByVal Flags As Long, ByVal vfsName As String) As Long
#Else
Public Function SQLite3OpenV2(ByVal FileName As String, ByRef dbHandle As Long, ByVal Flags As Long, ByVal vfsName As String) As Long
#End If

    Dim bufFileName() As Byte
    Dim bufVfsName() As Byte
    bufFileName = UTFlib.StringToUtf8Bytes(FileName)
    If vfsName = Empty Then
        SQLite3OpenV2 = sqlite3_open_v2(VarPtr(bufFileName(0)), dbHandle, Flags, 0)
    Else
        bufVfsName = UTFlib.StringToUtf8Bytes(vfsName)
        SQLite3OpenV2 = sqlite3_open_v2(VarPtr(bufFileName(0)), dbHandle, Flags, VarPtr(bufVfsName(0)))
    End If

End Function


#If WIN64 Then
Public Function SQLite3Close(ByVal dbHandle As LongPtr) As Long
#Else
Public Function SQLite3Close(ByVal dbHandle As Long) As Long
#End If
    SQLite3Close = sqlite3_close(dbHandle)
End Function

