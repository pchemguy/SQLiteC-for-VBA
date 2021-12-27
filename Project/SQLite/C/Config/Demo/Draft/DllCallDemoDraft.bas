Attribute VB_Name = "DllCallDemoDraft"
'@Folder "SQLite.C.Config.Demo.Draft"
'@IgnoreModule AssignedByValParameter, IndexedDefaultMemberAccess, ProcedureNotUsed
Option Explicit

Private Type TModuleState
    LongVal As Long
    LongRef As Long
    ByteVal As Byte
    ByteRef As Byte
    StrVal As String
    StrRef As String
End Type
Private this As TModuleState

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


'''' ====================================================== ''''

Private Sub CallFunctionArgs0Ret1Long()
    Dim dbm As SQLiteC
    Set dbm = FixObjC.GetDBM
    Dim DllMan As DllManager
    Set DllMan = dbm.DllMan
    Dim dbConf As DllCall
    Set dbConf = DllCall(DllMan)
    Dim Result As Long
    Result = dbConf.IndirectCall("SQLite3", "sqlite3_libversion_number", CC_STDCALL, vbLong, Empty)
    Debug.Print "SQLite3 version number: " & CStr(Result)
End Sub

Private Sub CallFunctionArgs0Ret1StrPtr()
    Dim dbm As SQLiteC
    Set dbm = FixObjC.GetDBM
    Dim DllMan As DllManager
    Set DllMan = dbm.DllMan
    Dim dbConf As DllCall
    Set dbConf = DllCall(DllMan)
    Dim Result As String
    #If Win64 Then
        Result = UTFlib.StrFromUTF8Ptr(dbConf.IndirectCall("SQLite3", "sqlite3_libversion", CC_STDCALL, vbLongLong, Empty))
    #Else
        Result = UTFlib.StrFromUTF8Ptr(dbConf.IndirectCall("SQLite3", "sqlite3_libversion", CC_STDCALL, vbLong, Empty))
    #End If
    Debug.Print "SQLite3 version number: " & CStr(Result)
End Sub


'''' ====================================================== ''''

Private Sub CallFunctionArgs6Ret1Long()
    Dim dbm As SQLiteC
    Set dbm = FixObjC.GetDBM
    Dim DllMan As DllManager
    Set DllMan = dbm.DllMan
    Dim dbConf As DllCall
    Set dbConf = DllCall(DllMan)
    dbConf.CacheProcPtr "DllCallDemo", "SixParamOneReturn", _
                        AddressOf SixParamOneReturn
    With this
        Dim Arguments(0 To 5) As Variant
        .LongVal = 30
        Arguments(0) = .LongVal
        .LongRef = 40
        Arguments(1) = VarPtr(.LongRef)
        .ByteVal = 10
        Arguments(2) = .ByteVal
        .ByteRef = 20
        Arguments(3) = VarPtr(.ByteRef)
        .StrVal = "StringVal"
        Arguments(4) = .StrVal
        .StrRef = "StringRef"
        Arguments(5) = VarPtr(.StrRef)
    End With
    
    Dim Result As Long
    Result = dbConf.IndirectCall("DllCallDemo", "SixParamOneReturn", _
        CC_STDCALL, vbLong, Arguments)
    Debug.Print "Result: " & CStr(Result)
End Sub

Private Function SixParamOneReturn( _
            ByVal LongVal As Long, ByRef LongRef As Long, _
            ByVal ByteVal As Byte, ByRef ByteRef As Byte, _
            ByVal StrVal As String, ByRef StrRef As String _
            ) As Long
    Debug.Print "===== SixParamOneReturn ====="
    Debug.Print "LongVal = " & CStr(LongVal)
    Debug.Print "LongRef = " & CStr(LongRef)
    Debug.Print "ByteVal = " & CStr(ByteVal)
    Debug.Print "ByteRef = " & CStr(ByteRef)
    Debug.Print "StrVal  = " & CStr(StrVal)
    Debug.Print "StrRef  = " & CStr(StrRef)
    SixParamOneReturn = LongVal + LongRef
    
    LongVal = 300
    LongRef = 400
    ByteVal = 100
    ByteRef = 200
    StrVal = "StrValNew"
    StrRef = "StrRefNew"
End Function


