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
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function sqlite3_libversion Lib "SQLite3" () As LongPtr ' PtrUtf8String
Private Declare PtrSafe Function sqlite3_libversion_number Lib "SQLite3" () As Long
#Else
Private Declare Function sqlite3_libversion Lib "SQLite3" () As Long ' PtrUtf8String
Private Declare Function sqlite3_libversion_number Lib "SQLite3" () As Long
#End If


'@DefaultMember
Public Function Create(ByVal FilePathName As String) As SQLiteCConnection
Attribute Create.VB_UserMemId = 0
    Dim Instance As SQLiteCConnection
    Set Instance = New SQLiteCConnection
    Instance.Init
    Set Create = Instance
End Function


Friend Sub Init()
:
End Sub


Public Function SQLite3Version() As String
    SQLite3Version = UTFlib.Utf8PtrToString(sqlite3_libversion())
End Function


Public Function SQLite3VersionNumber() As Long
    SQLite3VersionNumber = sqlite3_libversion_number()
End Function
