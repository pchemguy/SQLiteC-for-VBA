Attribute VB_Name = "DllExtAdapterEmbedDemo"
'@Folder "Custom and Extended DLL Demo"
Option Explicit
Option Private Module

#If WIN64 Then
Private Declare PtrSafe Function demo_sqlite3_extension_adapter Lib "SQLite3" (ByVal Dummy As Long) As Long
Private Declare PtrSafe Function sqlite3_libversion_number Lib "SQLite3" () As Long
#Else
Private Declare Function demo_sqlite3_extension_adapter Lib "SQLite3" (ByVal Dummy As Long) As Long
Private Declare Function sqlite3_libversion_number Lib "SQLite3" () As Long
#End If

Private Type TDllExtAdapterEmbedDemo
    DllMan As DllManager
End Type
Private this As TDllExtAdapterEmbedDemo


Private Sub GetDummySQLiteVersion()
    '''' Absolute or relative to ThisWorkbook.Path
    Dim DllPath As String
    #If WIN64 Then
        '''' TODO
        '''' DllPath = "Library\SQLiteCforVBA\Demo - DLL - STDCALL and Adapter\SQLite\x64"
    #Else
        DllPath = "Library\SQLiteCforVBA\Demo - DLL - STDCALL and Adapter\SQLite\x32"
    #End If
    LoadDlls DllPath
    
    Debug.Print 990000000 + sqlite3_libversion_number()
    Debug.Print demo_sqlite3_extension_adapter(990000000)
    Set this.DllMan = Nothing
End Sub


Private Sub LoadDlls(ByVal DllPath As String)
    Dim DllMan As DllManager
    Set DllMan = DllManager(DllPath)
    Set this.DllMan = DllMan
    Dim DllNames As Variant
    DllNames = Array( _
        "sqlite3.dll" _
    )
    DllMan.LoadMultiple DllNames
End Sub


