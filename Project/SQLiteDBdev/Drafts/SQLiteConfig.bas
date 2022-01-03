Attribute VB_Name = "SQLiteConfig"
'@Folder "SQLiteDBdev.Drafts"
'@IgnoreModule
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
Private Sub Main()
    Dim PtrType As VbVarType
    Dim DllNames As Variant
    #If Win64 Then
        PtrType = vbLongLong
        DllNames = "sqlite3.dll"
    #Else
        PtrType = vbLong
        DllNames = Array("icudt" & SQL_ICU_V & ".dll", "icuuc" & SQL_ICU_V & ".dll", "icuin" & SQL_ICU_V & ".dll", _
                         "icuio" & SQL_ICU_V & ".dll", "icutu" & SQL_ICU_V & ".dll", "sqlite3.dll")
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
    SQLiteVerStr = UTFlib.StrFromUTF8Ptr(DllMan.IndirectCall("SQLite3", "sqlite3_libversion", CC_STDCALL, PtrType, Empty))
    Debug.Print "SQLite version: " & SQLiteVerStr
    If Replace(Replace(SQLiteVerStr, ".", "0"), "0", vbNullString) = Replace(CStr(SQLiteVerLng), "0", vbNullString) Then
        Debug.Print "VERSIONS MATCHED"
    Else
        Debug.Print "VERSIONS MISMATCHED"
    End If
    Debug.Print "-------------------- SQLite --------------------" & vbNewLine
End Sub


Private Sub InitDBQC()
    '------------------------'
    '===== INIT MANAGER ====='
    '------------------------'
    Dim DllPath As String
    Dim DllNames As Variant
    #If Win64 Then
        DllPath = ThisWorkbook.Path & "\Library\SQLiteCAdo\dll\x64"
        DllNames = "sqlite3.dll"
    #Else
        DllPath = ThisWorkbook.Path & "\Library\SQLiteCAdo\dll\x32"
        DllNames = Array( _
            "icudt" & SQL_ICU_V & ".dll", "icuuc" & SQL_ICU_V & ".dll", _
            "icuin" & SQL_ICU_V & ".dll", "icuio" & SQL_ICU_V & ".dll", _
            "icutu" & SQL_ICU_V & ".dll", "sqlite3.dll")
    #End If
    Dim dbm As SQLiteC
    Set dbm = SQLiteC(DllPath, DllNames)
    If dbm Is Nothing Then
        Err.Raise ErrNo.ObjectCreateErr, "SQLiteCExamples", _
                  "Failed to create an SQLiteC instance."
    Else
        Debug.Print "Database manager instance (SQLiteC class) is ready"
    End If
    
    '''' ============================================================= ''''
    
    '---------------------------'
    '===== INIT CONNECTION ====='
    '---------------------------'
    Dim DbPathName As String
    DbPathName = "file::memory:"
    Dim dbc As SQLiteCConnection
    Set dbc = dbm.CreateConnection(DbPathName, AllowNonExistent:=True)
    If dbc Is Nothing Then
        Err.Raise ErrNo.ObjectCreateErr, "SQLiteCExamples", _
                  "Failed to create an SQLiteCConnection instance."
    Else
        Debug.Print "Database SQLiteCConnection instance is ready."
    End If


    '--------------------------'
    '===== INIT STATEMENT ====='
    '--------------------------'
    Dim DbStmtName As String
    DbStmtName = vbNullString
    Dim dbs As SQLiteCStatement
    Set dbs = dbc.CreateStatement(DbStmtName)
    '''' Maximum capapacity of 100x10 = 1000 rows
    dbs.DbExecutor.PageCount = 10
    dbs.DbExecutor.PageSize = 100
    
    Dim dbq As ILiteADO
    Set dbq = dbs
    If dbq Is Nothing Then
        Err.Raise ErrNo.ObjectCreateErr, "SQLiteCExamples", _
                  "Failed to create an SQLiteCStatement instance."
    Else
        Debug.Print "Database SQLiteCStatement instance is ready."
    End If
    Debug.Print "Created blank db: " & dbq.MainDB
End Sub
