Attribute VB_Name = "DllCallDemoSQLite"
'@Folder "DllTools.Demo.DllCall"
''''
'''' WARNING: Dll calls can crash the application. With calls via DispCallFunc,
'''' the VBA compiler cannot perform any correctness checks on the target call.
'''' Make sure your work is saved and be prepared for Excel crashing.
''''
Option Explicit

Private Const LIB_NAME As String = "DllTools"
Private Const PATH_SEP As String = "\"
Private Const LIB_RPREFIX As String = _
    "Library" & PATH_SEP & LIB_NAME & PATH_SEP & "dll" & PATH_SEP

'''' This demo calls two SQLite functions with the following VBA signatures:
''''   Private Declare PtrSafe Function sqlite3_libversion Lib "SQLite3" () As LongPtr     'PtrUtf8String
''''   Private Declare PtrSafe Function sqlite3_libversion_number Lib "SQLite3" () As Long
'''' If successful, this routine should print out numeric and textual forms of
'''' the SQLite library being used and should print "VERSIONS MATCHED" message.
''''
'@EntryPoint
Private Sub Main()
    Dim PtrType As VbVarType
    Dim DllNames As Variant
    #If Win64 Then
        PtrType = vbLongLong
        DllNames = "sqlite3.dll"
    #Else
        PtrType = vbLong
        DllNames = Array("icudt68.dll", "icuuc68.dll", "icuin68.dll", _
                         "icuio68.dll", "icutu68.dll", "sqlite3.dll")
    #End If
    Dim DllPath As String
    DllPath = LIB_RPREFIX & ARCH
    Debug.Print "==================== SQLite ===================="
    Dim DllMan As DllManager
    Set DllMan = DllManager.Create(DllPath, DllNames)
    Dim SQLiteVerLng As Long
    SQLiteVerLng = DllMan.IndirectCall("SQLite3", "sqlite3_libversion_number", CC_STDCALL, vbLong, Empty)
    Debug.Print "SQLite version: " & CStr(SQLiteVerLng)
    Dim SQLiteVerStr As String
    SQLiteVerStr = UTFlib.StrFromUTF8Ptr( _
        DllMan.IndirectCall("SQLite3", "sqlite3_libversion", CC_STDCALL, PtrType, Empty))
    Debug.Print "SQLite version: " & SQLiteVerStr
    If Replace(Replace(SQLiteVerStr, ".", "0"), "0", vbNullString) = _
       Replace(CStr(SQLiteVerLng), "0", vbNullString) Then
        Debug.Print "VERSIONS MATCHED"
    Else
        Debug.Print "VERSIONS MISMATCHED"
    End If
    Debug.Print "-------------------- SQLite --------------------" & vbNewLine
End Sub
