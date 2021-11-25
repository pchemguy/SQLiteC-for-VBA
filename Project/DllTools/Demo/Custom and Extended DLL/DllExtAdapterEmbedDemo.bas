Attribute VB_Name = "DllExtAdapterEmbedDemo"
'@Folder "DllTools.Demo.Custom and Extended DLL"
Option Explicit
Option Private Module

Private Const LIB_NAME As String = "DllTools"
Private Const PATH_SEP As String = "\"
Private Const LIB_RPREFIX As String = _
    "Library" & PATH_SEP & LIB_NAME & PATH_SEP & _
    "Demo - DLL - STDCALL and Adapter" & PATH_SEP

#If Win64 Then
Private Declare PtrSafe Function demo_sqlite3_extension_adapter Lib "SQLite3demo" (ByVal Dummy As Long) As Long
Private Declare PtrSafe Function sqlite3_libversion_number Lib "SQLite3demo" () As Long
#Else
Private Declare Function demo_sqlite3_extension_adapter Lib "SQLite3demo" (ByVal Dummy As Long) As Long
Private Declare Function sqlite3_libversion_number Lib "SQLite3demo" () As Long
#End If

Private Type TDllExtAdapterEmbedDemo
    DllMan As DllManager
End Type
Private this As TDllExtAdapterEmbedDemo


'@Ignore ProcedureNotUsed
Private Sub GetSQLiteVersion()
    '''' Absolute or relative to ThisWorkbook.Path
    Dim DllPath As String
    #If Win64 Then
        '''' TODO
        '''' DllPath = thisworkbook.path & path_sep & LIB_RPREFIX & "SQLite\x64"
        DllPath = vbNullString
    #Else
        DllPath = ThisWorkbook.Path & PATH_SEP & LIB_RPREFIX & "SQLite\x32"
    #End If
    LoadDlls DllPath
    
    '''' Calling CDECL without arguments
    Debug.Print 990000000 + sqlite3_libversion_number()
    '''' Calling STDCALL with arguments
    Debug.Print demo_sqlite3_extension_adapter(990000000)
    Set this.DllMan = Nothing
End Sub


Private Sub LoadDlls(ByVal DllPath As String)
    Dim DllMan As DllManager
    DllManager.Free
    Dim DllName As String
    DllName = "sqlite3demo.dll"
    Set DllMan = DllManager.Create(DllPath, DllName)
    Set this.DllMan = DllMan
End Sub
