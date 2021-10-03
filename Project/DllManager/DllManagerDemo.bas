Attribute VB_Name = "DllManagerDemo"
'@Folder "DllManager"
Option Explicit
Option Private Module


Private Sub LoadLibs()
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
