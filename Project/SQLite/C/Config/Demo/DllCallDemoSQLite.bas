Attribute VB_Name = "DllCallDemoSQLite"
'@Folder "SQLite.C.Config.Demo"
''''
'''' WARNING: Dll calls can crash the application. With calls via DispCallFunc,
'''' the VBA compiler cannot perform any correctness checks on the target call.
'''' Make sure your work is saved and be prepared for Excel crashing.
''''
Option Explicit


'''' This demo calls two SQLite functions with the following VBA signatures:
''''   Private Declare PtrSafe Function sqlite3_libversion Lib "SQLite3" () As LongPtr     'PtrUtf8String
''''   Private Declare PtrSafe Function sqlite3_libversion_number Lib "SQLite3" () As Long
'''' If successful, this routine should print out numeric and textual forms of
'''' the SQLite library being used and should print "VERSIONS MATCHED" message.
''''
Private Sub Main()
    Dim TestTarget As String
    Dim PtrType As VbVarType
    #If Win64 Then
        PtrType = vbLongLong
    #Else
        PtrType = vbLong
    #End If
    TestTarget = "==================== SQLite ===================="
    Debug.Print TestTarget
    Dim dbm As SQLiteC
    Set dbm = FixObjC.GetDBM
    Dim DllMan As DllManager
    Set DllMan = dbm.DllMan
    Dim dbConf As DllCall
    Set dbConf = DllCall(DllMan)
    Dim SQLiteVerLng As Long
    SQLiteVerLng = dbConf.IndirectCall("SQLite3", "sqlite3_libversion_number", CC_STDCALL, vbLong, Empty)
    Debug.Print "SQLite version: " & CStr(SQLiteVerLng)
    Dim SQLiteVerStr As String
    SQLiteVerStr = UTFlib.StrFromUTF8Ptr(dbConf.IndirectCall("SQLite3", "sqlite3_libversion", CC_STDCALL, PtrType, Empty))
    Debug.Print "SQLite version: " & SQLiteVerStr
    If Replace(Replace(SQLiteVerStr, ".", "0"), "0", vbNullString) = Replace(CStr(SQLiteVerLng), "0", vbNullString) Then
        Debug.Print "VERSIONS MATCHED"
    Else
        Debug.Print "VERSIONS MISMATCHED"
    End If
    TestTarget = "-------------------- SQLite --------------------" & vbNewLine
    Debug.Print TestTarget
End Sub
