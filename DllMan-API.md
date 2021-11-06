---
layout: default
title: API
nav_order: 1
parent: DLL Manager
permalink: /dllman/api
---

### VBA class managing loading DLL libraries

Before a DLL library placed in a user directory is available from the VBA code, it must be loaded via the Windows API (alternatively, the Declare statement may include the library location, but this approach is ugly and inconvenient). It is also prudent to unload the library when no longer needed. To make the load/unload process more robust, I created the DllManager class wrapping the LoadLibrary, FreeLibrary, and SetDllDirectory [APIs][DLL API]. DllManager can be used for loading/unloading multiple DLLs. It wraps a Scripting.Dictionary object to hold \<DLL name\>&nbsp;&rarr;&nbsp;\<DLL handle\> mapping.

**API**

*DllManager.Create* factory takes one optional parameter, indicating the user's DLL location, and passes it to *DllManager.Init* constructor. Ultimately, the *DefaultPath* setter (Property Let) handles this parameter. The setter checks if the parameter holds a valid absolute or a relative (w.r.t. ThisWorkbook.Path) path. If this check succeeds, SetDllDirectory API sets the default DLL search path. *DllManager.ResetDllSearchPath* can be used to reset the DLL search path to its default value.

*DllManager.Load* loads individual libraries. It takes the target library name and, optionally, path. If the target library has not been loaded, it attempts to resolve the DLL location by checking the provided value and the DefaultPath attribute. If resolution succeeds, the LoadLibrary API is called. *DllManager.Free*, in turn, unloads the previously loaded library.

*DllManager.LoadMultiple* loads a list of libraries. It takes a variable list of arguments (ParamArray) and loads them in the order provided. Alternatively, it also accepts a 0-based array of names as the sole argument. *DllManager.FreeMultiple* is the counterpart of *.LoadMultiple* with the same interface. If no arguments are provided, all loaded libraries are unloaded.

Finally, while *.Free/.FreeMultiple* can be called explicitly, *Class_Terminate* calls  *.FreeMultiple* and *.ResetDllSearchPath* automatically before the object is destroyed.

**Demo**

The *DllManagerDemo* example below illustrates how this class can be used and compares the usage patterns between system and user libraries. In this case, *WinSQLite3* system library is used as a reference (see *GetWinSQLite3VersionNumber*). A call to a custom compiled SQLite library placed in the project folder demos the additional code necessary to make such a call (see *GetSQLite3VersionNumber*). In both cases, *sqlite3_libversion_number* routine, returning the numeric library version, is declared and called.

```vb
'@Folder "DllManager"
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
```


[DLL API]: https://docs.microsoft.com/en-us/windows/win32/dlls/dynamic-link-library-functions
[SQLite VBA]: https://pchemguy.github.io/SQLite-ICU-MinGW/stdcall
