Attribute VB_Name = "DllManagerDemo"
'@Folder "DllManager"
'@IgnoreModule ProcedureNotUsed, IndexedDefaultMemberAccess
Option Explicit
Option Private Module

#If VBA7 Then
'''' System library
Private Declare PtrSafe Function winsqlite3_libversion_number Lib "WinSQLite3" Alias "sqlite3_libversion_number" () As Long
'''' User library
Private Declare PtrSafe Function sqlite3_libversion_number Lib "SQLite3" () As Long
#Else
'''' System library
Private Declare Function winsqlite3_libversion_number Lib "WinSQLite3" Alias "sqlite3_libversion_number" () As Long
'''' User library
Private Declare Function sqlite3_libversion_number Lib "SQLite3" () As Long
#End If


Private Type TDllManagerDemo
    DllMan As DllManager
End Type
Private this As TDllManagerDemo


Private Sub GetWinSQLite3VersionNumber()
    Debug.Print winsqlite3_libversion_number()
End Sub


Private Sub GetSQLite3VersionNumber()
    '''' Absolute or relative to ThisWorkbook.Path
    Dim DllPath As String
    DllPath = "Library\SQLiteCforVBA\dll\x32"
    SQLiteLoadMultipleArray DllPath
    
    Debug.Print sqlite3_libversion_number()
    Set this.DllMan = Nothing
End Sub


Private Sub SQLiteLoadMultipleArray(ByVal DllPath As String)
    Dim DllMan As DllManager
    Set DllMan = DllManager(DllPath)
    Set this.DllMan = DllMan
    Dim DllNames As Variant
    DllNames = Array( _
        "icudt68.dll", _
        "icuuc68.dll", _
        "icuin68.dll", _
        "icuio68.dll", _
        "icutu68.dll", _
        "sqlite3.dll" _
    )
    DllMan.LoadMultiple DllNames
End Sub


' ========================= '
' Additional usage examples '
' ========================= '
Private Sub SQLiteLoadMultipleParamArray()
    Dim RelativePath As String
    RelativePath = "Library\SQLiteCforVBA\dll\x32"
    Dim DllMan As DllManager
    Set DllMan = DllManager(RelativePath)
    DllMan.LoadMultiple _
        "icudt68.dll", _
        "icuuc68.dll", _
        "icuin68.dll", _
        "icuio68.dll", _
        "icutu68.dll", _
        "sqlite3.dll"
End Sub


Private Sub SQLiteLoad()
    Dim RelativePath As String
    RelativePath = "Library\SQLiteCforVBA\dll\x32"
    Dim DllMan As DllManager
    Set DllMan = DllManager(RelativePath)
    Dim DllNames As Variant
    DllNames = Array( _
        "icudt68.dll", _
        "icuuc68.dll", _
        "icuin68.dll", _
        "icuio68.dll", _
        "icutu68.dll", _
        "sqlite3.dll" _
    )
    Dim DllNameIndex As Long
    For DllNameIndex = LBound(DllNames) To UBound(DllNames)
        Dim DllName As String
        DllName = DllNames(DllNameIndex)
        DllMan.Load DllName, RelativePath
    Next DllNameIndex
End Sub
