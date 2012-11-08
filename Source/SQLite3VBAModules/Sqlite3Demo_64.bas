Attribute VB_Name = "Sqlite3Demo"
Option Explicit

Public Sub AllTests()
    Dim InitReturn As Long
    
    #If Win64 Then
        ' I put the 64-bit version of SQLite.dll under a subdirectory called x64
        InitReturn = SQLite3Initialize(ThisWorkbook.Path + "\x64")
    #Else
        InitReturn = SQLite3Initialize ' Default path is ThisWorkbook.Path but can specify other path where the .dlls reside.
    #End If
    If InitReturn <> SQLITE_INIT_OK Then
        Debug.Print "Error Initializing SQLite. Error: " & Err.LastDllError
        Exit Sub
    End If
    
    TestVersion
    TestOpenClose
    TestOpenCloseV2
    TestError
    TestInsert
    TestSelect
    TestBinding
    TestDates
    TestStrings
    TestBackup
    TestBlob
    TestWriteReadOnly
    SQLite3Free ' Quite optional
End Sub

Public Sub TestVersion()

    Debug.Print SQLite3LibVersion()

End Sub

Public Sub TestApiCallSpeed()
    
    Dim i As Long
    Dim version As String
    Dim start As Date
    
    start = Now()
    For i = 0 To 10000000 ' 10 million
        version = SQLite3LibVersion()
    Next
    
    Debug.Print "ApiCall Elapsed: " & Format(Now() - start, "HH:mm:ss")
    
End Sub

Public Sub TestOpenClose()
    Dim testFile As String
    #If Win64 Then
    Dim myDbHandle As LongPtr
    #Else
    Dim myDbHandle As Long
    #End If
    Dim RetVal As Long
    
    testFile = "C:\TestSqlite3ForExcel.db3"
    RetVal = SQLite3Open(testFile, myDbHandle)
    Debug.Print "SQLite3Open returned " & RetVal
    
    RetVal = SQLite3Close(myDbHandle)
    Debug.Print "SQLite3Close returned " & RetVal
    
    Kill testFile

End Sub

Public Sub TestOpenCloseV2()
    Dim testFile As String
    #If Win64 Then
    Dim myDbHandle As LongPtr
    Dim myDbHandleV2 As LongPtr
    #Else
    Dim myDbHandle As Long
    Dim myDbHandleV2 As Long
    #End If
    Dim RetVal As Long
    
    ' Open the database in Read Write Access
    testFile = "C:\TestSqlite3ForExcel.db3"
    RetVal = SQLite3Open(testFile, myDbHandle)
    Debug.Print "SQLite3Open returned " & RetVal
    
    ' Open the database in Read Only Access
    testFile = "C:\TestSqlite3ForExcel.db3"
    RetVal = SQLite3OpenV2(testFile, myDbHandleV2, SQLITE_OPEN_READONLY, "")
    Debug.Print "SQLite3OpenV2 returned " & RetVal
    
    RetVal = SQLite3Close(myDbHandleV2)
    Debug.Print "SQLite3Close V2 returned " & RetVal
    
    RetVal = SQLite3Close(myDbHandle)
    Debug.Print "SQLite3Close returned " & RetVal
    
    Kill testFile

End Sub

Public Sub TestError()
    #If Win64 Then
    Dim myDbHandle As LongPtr
    #Else
    Dim myDbHandle As Long
    #End If
    Dim RetVal As Long
    
    Dim ErrMsg As String
    
    Debug.Print "----- TestError Start -----"
    
    ' DbHandle is set up even if there is an error !
    RetVal = SQLite3Open("::::", myDbHandle)
    Debug.Print "SQLite3Open returned " & RetVal
    
    ErrMsg = SQLite3ErrMsg(myDbHandle)
    Debug.Print "SQLite3Open error message: " & ErrMsg
  
    RetVal = SQLite3Close(myDbHandle)
    Debug.Print "SQLite3Close returned " & RetVal

    Debug.Print "----- TestError End -----"

End Sub

Public Sub TestStatement()
    Dim testFile As String
    
    #If Win64 Then
    Dim myDbHandle As LongPtr
    Dim myStmtHandle As LongPtr
    #Else
    Dim myDbHandle As Long
    Dim myStmtHandle As Long
    #End If

    Dim RetVal As Long
    
    Dim stepMsg As String
    
    Debug.Print "----- TestStatement Start -----"
    
    ' Open the database - getting a DbHandle back
    testFile = "C:\TestSqlite3ForExcel.db3"
    RetVal = SQLite3Open(testFile, myDbHandle)
    Debug.Print "SQLite3Open returned " & RetVal
    
    ' Create the sql statement - getting a StmtHandle back
    RetVal = SQLite3PrepareV2(myDbHandle, "CREATE TABLE MyFirstTable (TheId INTEGER, TheText TEXT, TheValue REAL)", myStmtHandle)
    Debug.Print "SQLite3PrepareV2 returned " & RetVal
    
    ' Start running the statement
    RetVal = SQLite3Step(myStmtHandle)
    Debug.Print "SQLite3Step returned " & RetVal
    
    ' Finalize (delete) the statement
    RetVal = SQLite3Finalize(myStmtHandle)
    Debug.Print "SQLite3Finalize returned " & RetVal
    
    ' Close the database
    RetVal = SQLite3Close(myDbHandle)
    Kill testFile

    Debug.Print "----- TestStatement End -----"
End Sub

Public Sub TestInsert()
    Dim testFile As String
    
    #If Win64 Then
    Dim myDbHandle As LongPtr
    Dim myStmtHandle As LongPtr
    #Else
    Dim myDbHandle As Long
    Dim myStmtHandle As Long
    #End If
    Dim RetVal As Long
    Dim recordsAffected As Long
    
    Dim stepMsg As String
    
    Debug.Print "----- TestInsert Start -----"
    
    ' Open the database - getting a DbHandle back
    testFile = "C:\TestSqlite3ForExcel.db3"
    RetVal = SQLite3Open(testFile, myDbHandle)
    Debug.Print "SQLite3Open returned " & RetVal
    
    '------------------------
    ' Create the table
    ' ================
    ' Create the sql statement - getting a StmtHandle back
    RetVal = SQLite3PrepareV2(myDbHandle, "CREATE TABLE MySecondTable (TheId INTEGER, TheText TEXT, TheValue REAL)", myStmtHandle)
    Debug.Print "SQLite3PrepareV2 returned " & RetVal
    
    ' Start running the statement
    RetVal = SQLite3Step(myStmtHandle)
    If RetVal = SQLITE_DONE Then
        Debug.Print "SQLite3Step Done"
    Else
        Debug.Print "SQLite3Step returned " & RetVal
    End If
    
    ' Finalize (delete) the statement
    RetVal = SQLite3Finalize(myStmtHandle)
    Debug.Print "SQLite3Finalize returned " & RetVal
    
    '-------------------------
    ' Insert a record
    ' ===============
    ' Create the sql statement - getting a StmtHandle back
    RetVal = SQLite3PrepareV2(myDbHandle, "INSERT INTO MySecondTable Values (123, 'ABC', 42.1)", myStmtHandle)
    Debug.Print "SQLite3PrepareV2 returned " & RetVal
    
    ' Start running the statement
    RetVal = SQLite3Step(myStmtHandle)
    If RetVal = SQLITE_DONE Then
        Debug.Print "SQLite3Step Done"
    Else
        Debug.Print "SQLite3Step returned " & RetVal
    End If
    
    ' Finalize (delete) the statement
    RetVal = SQLite3Finalize(myStmtHandle)
    Debug.Print "SQLite3Finalize returned " & RetVal

    '-------------------------
    ' Insert  using helper
    ' ====================
    recordsAffected = SQLite3ExecuteNonQuery(myDbHandle, "INSERT INTO MySecondTable Values (456, 'DEF', 49.3)")
    Debug.Print "SQLite3Execute - Insert affected " & recordsAffected & " record(s)."
    
    ' Close the database
    RetVal = SQLite3Close(myDbHandle)
    Kill testFile

    Debug.Print "----- TestInsert End -----"
End Sub

Public Sub TestSelect()
    Dim testFile As String
    
    #If Win64 Then
    Dim myDbHandle As LongPtr
    Dim myStmtHandle As LongPtr
    #Else
    Dim myDbHandle As Long
    Dim myStmtHandle As Long
    #End If
    Dim RetVal As Long
    
    Dim stepMsg As String
    
    Debug.Print "----- TestSelect Start -----"
    
    ' Open the database - getting a DbHandle back
    testFile = "C:\TestSqlite3ForExcel.db3"
    RetVal = SQLite3Open(testFile, myDbHandle)
    Debug.Print "SQLite3Open returned " & RetVal
    
    '------------------------
    ' Create the table
    ' ================
    ' Create the sql statement - getting a StmtHandle back
    RetVal = SQLite3PrepareV2(myDbHandle, "CREATE TABLE MyFirstTable (TheId INTEGER, TheText TEXT, TheValue REAL)", myStmtHandle)
    Debug.Print "SQLite3PrepareV2 returned " & RetVal
    
    ' Start running the statement
    RetVal = SQLite3Step(myStmtHandle)
    If RetVal = SQLITE_DONE Then
        Debug.Print "SQLite3Step Done"
    Else
        Debug.Print "SQLite3Step returned " & RetVal
    End If
    
    ' Finalize (delete) the statement
    RetVal = SQLite3Finalize(myStmtHandle)
    Debug.Print "SQLite3Finalize returned " & RetVal
    
    '-------------------------
    ' Insert a record
    ' ===============
    ' Create the sql statement - getting a StmtHandle back
    RetVal = SQLite3PrepareV2(myDbHandle, "INSERT INTO MyFirstTable Values (123, 'ABC', 42.1)", myStmtHandle)
    Debug.Print "SQLite3PrepareV2 returned " & RetVal
    
    ' Start running the statement
    RetVal = SQLite3Step(myStmtHandle)
    If RetVal = SQLITE_DONE Then
        Debug.Print "SQLite3Step Done"
    Else
        Debug.Print "SQLite3Step returned " & RetVal
    End If
    
    ' Finalize (delete) the statement
    RetVal = SQLite3Finalize(myStmtHandle)
    Debug.Print "SQLite3Finalize returned " & RetVal

    '-------------------------
    ' Insert another record
    ' ===============
    ' Create the sql statement - getting a StmtHandle back
    RetVal = SQLite3PrepareV2(myDbHandle, "INSERT INTO MyFirstTable Values (987654, ""ZXCVBNM"", NULL)", myStmtHandle)
    Debug.Print "SQLite3PrepareV2 returned " & RetVal
    
    ' Start running the statement
    RetVal = SQLite3Step(myStmtHandle)
    If RetVal = SQLITE_DONE Then
        Debug.Print "SQLite3Step Done"
    Else
        Debug.Print "SQLite3Step returned " & RetVal
    End If
    
    ' Finalize (delete) the statement
    RetVal = SQLite3Finalize(myStmtHandle)
    Debug.Print "SQLite3Finalize returned " & RetVal

    '-------------------------
    ' Select statement
    ' ===============
    ' Create the sql statement - getting a StmtHandle back
    RetVal = SQLite3PrepareV2(myDbHandle, "SELECT * FROM MyFirstTable", myStmtHandle)
    Debug.Print "SQLite3PrepareV2 returned " & RetVal
    
    ' Start running the statement
    RetVal = SQLite3Step(myStmtHandle)
    If RetVal = SQLITE_ROW Then
        Debug.Print "SQLite3Step Row Ready"
        PrintColumns myStmtHandle
    Else
        Debug.Print "SQLite3Step returned " & RetVal
    End If
    
    ' Move to next row
    RetVal = SQLite3Step(myStmtHandle)
    If RetVal = SQLITE_ROW Then
        Debug.Print "SQLite3Step Row Ready"
        PrintColumns myStmtHandle
    Else
        Debug.Print "SQLite3Step returned " & RetVal
    End If
    
    ' Move on again (now we are done)
    RetVal = SQLite3Step(myStmtHandle)
    If RetVal = SQLITE_DONE Then
        Debug.Print "SQLite3Step Done"
    Else
        Debug.Print "SQLite3Step returned " & RetVal
    End If
    
    ' Finalize (delete) the statement
    RetVal = SQLite3Finalize(myStmtHandle)
    Debug.Print "SQLite3Finalize returned " & RetVal

    
    ' Close the database
    RetVal = SQLite3Close(myDbHandle)
    Kill testFile

    Debug.Print "----- TestSelect End -----"
End Sub

#If Win64 Then
Sub PrintColumns(ByVal stmtHandle As LongPtr)
#Else
Sub PrintColumns(ByVal stmtHandle As Long)
#End If
    Dim colCount As Long
    Dim colName As String
    Dim colType As Long
    Dim colTypeName As String
    Dim colValue As Variant
    
    Dim i As Long
    
    colCount = SQLite3ColumnCount(stmtHandle)
    Debug.Print "Column count: " & colCount
    For i = 0 To colCount - 1
        colName = SQLite3ColumnName(stmtHandle, i)
        colType = SQLite3ColumnType(stmtHandle, i)
        colTypeName = TypeName(colType)
        colValue = ColumnValue(stmtHandle, i, colType)
        Debug.Print "Column " & i & ":", colName, colTypeName, colValue
    Next
End Sub

#If Win64 Then
Sub PrintParameters(ByVal stmtHandle As LongPtr)
#Else
Sub PrintParameters(ByVal stmtHandle As Long)
#End If
    Dim paramCount As Long
    Dim paramName As String
    
    Dim i As Long
    
    paramCount = SQLite3BindParameterCount(stmtHandle)
    Debug.Print "Parameter count: " & paramCount
    For i = 1 To paramCount
        paramName = SQLite3BindParameterName(stmtHandle, i)
        Debug.Print "Parameter " & i & ":", paramName
    Next
End Sub

Function TypeName(ByVal SQLiteType As Long) As String
    Select Case SQLiteType
        Case SQLITE_INTEGER:
            TypeName = "INTEGER"
        Case SQLITE_FLOAT:
            TypeName = "FLOAT"
        Case SQLITE_TEXT:
            TypeName = "TEXT"
        Case SQLITE_BLOB:
            TypeName = "BLOB"
        Case SQLITE_NULL:
            TypeName = "NULL"
    End Select
End Function

#If Win64 Then
Function ColumnValue(ByVal stmtHandle As LongPtr, ByVal ZeroBasedColIndex As Long, ByVal SQLiteType As Long) As Variant
#Else
Function ColumnValue(ByVal stmtHandle As Long, ByVal ZeroBasedColIndex As Long, ByVal SQLiteType As Long) As Variant
#End If
    Select Case SQLiteType
        Case SQLITE_INTEGER:
            ColumnValue = SQLite3ColumnInt32(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_FLOAT:
            ColumnValue = SQLite3ColumnDouble(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_TEXT:
            ColumnValue = SQLite3ColumnText(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_BLOB:
            ColumnValue = SQLite3ColumnText(stmtHandle, ZeroBasedColIndex)
        Case SQLITE_NULL:
            ColumnValue = Null
    End Select
End Function

Public Sub TestBinding()
    Dim testFile As String
    
    #If Win64 Then
    Dim myDbHandle As LongPtr
    Dim myStmtHandle As LongPtr
    #Else
    Dim myDbHandle As Long
    Dim myStmtHandle As Long
    #End If

    Dim RetVal As Long
    Dim stepMsg As String
    Dim i As Long
    
    Dim paramIndexId As Long
    Dim paramIndexDate As Long
    
    Dim startDate As Date
    Dim curDate As Date
    Dim curValue As Double
    Dim offset As Long
    
    Dim testStart As Date
    
    Debug.Print "----- TestBinding Start -----"
    
    ' Open the database - getting a DbHandle back
    testFile = "C:\TestSqlite3ForExcel.db3"
    RetVal = SQLite3Open(testFile, myDbHandle)
    Debug.Print "SQLite3Open returned " & RetVal
    
    '------------------------
    ' Create the table
    ' ================
    ' (O've got no error checking here...)
    SQLite3PrepareV2 myDbHandle, "CREATE TABLE MyBigTable (TheId INTEGER, TheDate REAL, TheText TEXT, TheValue REAL)", myStmtHandle
    SQLite3Step myStmtHandle
    SQLite3Finalize myStmtHandle
    
    '---------------------------
    ' Add an index
    ' ================
    SQLite3PrepareV2 myDbHandle, "CREATE INDEX idx_MyBigTable_Id_Date ON MyBigTable (TheId, TheDate)", myStmtHandle
    SQLite3Step myStmtHandle
    SQLite3Finalize myStmtHandle
    
    ' START Insert Time
    testStart = Now()
    
    '-------------------
    ' Begin transaction
    '==================
    SQLite3PrepareV2 myDbHandle, "BEGIN TRANSACTION", myStmtHandle
    SQLite3Step myStmtHandle
    SQLite3Finalize myStmtHandle

    '-------------------------
    ' Prepare an insert statement with parameters
    ' ===============
    ' Create the sql statement - getting a StmtHandle back
    RetVal = SQLite3PrepareV2(myDbHandle, "INSERT INTO MyBigTable Values (?, ?, ?, ?)", myStmtHandle)
    If RetVal <> SQLITE_OK Then
        Debug.Print "SQLite3PrepareV2 returned " & SQLite3ErrMsg(myDbHandle)
        Beep
    End If
    
    Randomize
    startDate = DateValue("1 Jan 2000")
    
    For i = 1 To 100000
        curDate = startDate + i
        curValue = Rnd() * 1000
        
        RetVal = SQLite3BindInt32(myStmtHandle, 1, 42000 + i)
        If RetVal <> SQLITE_OK Then
            Debug.Print "SQLite3Bind returned " & RetVal, SQLite3ErrMsg(myDbHandle)
            Beep
        End If
        
        RetVal = SQLite3BindDate(myStmtHandle, 2, curDate)
        If RetVal <> SQLITE_OK Then
            Debug.Print "SQLite3Bind returned " & RetVal, SQLite3ErrMsg(myDbHandle)
            Beep
        End If
        
        RetVal = SQLite3BindText(myStmtHandle, 3, "The quick brown fox jumped over the lazy dog.")
        If RetVal <> SQLITE_OK Then
            Debug.Print "SQLite3Bind returned " & RetVal, SQLite3ErrMsg(myDbHandle)
            Beep
        End If
        
        RetVal = SQLite3BindDouble(myStmtHandle, 4, curValue)
        If RetVal <> SQLITE_OK Then
            Debug.Print "SQLite3Bind returned " & RetVal, SQLite3ErrMsg(myDbHandle)
            Beep
        End If
        
        RetVal = SQLite3Step(myStmtHandle)
        If RetVal <> SQLITE_DONE Then
            Debug.Print "SQLite3Step returned " & RetVal, SQLite3ErrMsg(myDbHandle)
            Beep
        End If
    
        RetVal = SQLite3Reset(myStmtHandle)
        If RetVal <> SQLITE_OK Then
            Debug.Print "SQLite3Reset returned " & RetVal, SQLite3ErrMsg(myDbHandle)
            Beep
        End If
    Next
    
    ' Finalize (delete) the statement
    RetVal = SQLite3Finalize(myStmtHandle)
    Debug.Print "SQLite3Finalize returned " & RetVal

    '-------------------
    ' Commit transaction
    '==================
    ' (I'm re-using the same variable myStmtHandle for the new statement)
    SQLite3PrepareV2 myDbHandle, "COMMIT TRANSACTION", myStmtHandle
    SQLite3Step myStmtHandle
    SQLite3Finalize myStmtHandle

    ' STOP Insert Time
    Debug.Print "Insert Elapsed: " & Format(Now() - testStart, "HH:mm:ss")

    ' START Select  Time
    testStart = Now()

    '-------------------------
    ' Select statement
    ' ===============
    ' Create the sql statement - getting a StmtHandle back
    ' Now using named parameters!
    RetVal = SQLite3PrepareV2(myDbHandle, "SELECT TheId, datetime(TheDate), TheText, TheValue FROM MyBigTable WHERE TheId = @FindThisId AND TheDate <= @FindThisDate LIMIT 1", myStmtHandle)
    Debug.Print "SQLite3PrepareV2 returned " & RetVal
    
    paramIndexId = SQLite3BindParameterIndex(myStmtHandle, "@FindThisId")
    If paramIndexId = 0 Then
        Debug.Print "SQLite3BindParameterIndex could not find the Id parameter!"
        Beep
    End If
    
    paramIndexDate = SQLite3BindParameterIndex(myStmtHandle, "@FindThisDate")
    If paramIndexDate = 0 Then
        Debug.Print "SQLite3BindParameterIndex could not find the Date parameter!"
        Beep
    End If
    
    startDate = DateValue("1 Jan 2000")
    
    
    For i = 1 To 100000
        offset = i Mod 10000
        ' Bind the parameters
        RetVal = SQLite3BindInt32(myStmtHandle, paramIndexId, 42000 + 500 + offset)
        If RetVal <> SQLITE_OK Then
            Debug.Print "SQLite3Bind returned " & RetVal, SQLite3ErrMsg(myDbHandle)
            Beep
        End If
    
        RetVal = SQLite3BindDate(myStmtHandle, paramIndexDate, startDate + 500 + offset)
        If RetVal <> SQLITE_OK Then
            Debug.Print "SQLite3Bind returned " & RetVal, SQLite3ErrMsg(myDbHandle)
            Beep
        End If
        
        RetVal = SQLite3Step(myStmtHandle)
        If RetVal = SQLITE_ROW Then
            ' We have access to the result columns here.
            If offset = 1 Then
                Debug.Print "At row " & i
                Debug.Print "------------"
                PrintColumns myStmtHandle
                Debug.Print "============"
            End If
        ElseIf RetVal = SQLITE_DONE Then
            Debug.Print "No row found"
        End If
    
        RetVal = SQLite3Reset(myStmtHandle)
        If RetVal <> SQLITE_OK Then
            Debug.Print "SQLite3Reset returned " & RetVal, SQLite3ErrMsg(myDbHandle)
            Beep
        End If
    Next
        
    ' Finalize (delete) the statement
    RetVal = SQLite3Finalize(myStmtHandle)
    Debug.Print "SQLite3Finalize returned " & RetVal
    
    ' STOP Select time
    Debug.Print "Select Elapsed: " & Format(Now() - testStart, "HH:mm:ss")
    
    ' Close the database
    RetVal = SQLite3Close(myDbHandle)
    Kill testFile

    Debug.Print "----- TestBinding End -----"
End Sub


Public Sub TestBindingMore()
    Dim testFile As String
    
    #If Win64 Then
    Dim myDbHandle As LongPtr
    Dim myStmtHandle As LongPtr
    #Else
    Dim myDbHandle As Long
    Dim myStmtHandle As Long
    #End If

    Dim RetVal As Long
    Dim stepMsg As String
    Dim i As Long
    
    Dim paramIndexId As Long
    Dim paramIndexDate As Long
    
    Dim startDate As Date
    Dim curDate As Date
    Dim curValue As Double
    Dim offset As Long
    
    Dim testStart As Date
    
    Debug.Print "----- TestBinding Start -----"
    
    ' Open the database - getting a DbHandle back
    testFile = "C:\TestSqlite3ForExcel.db3"
    RetVal = SQLite3Open(testFile, myDbHandle)
    Debug.Print "SQLite3Open returned " & RetVal
    
    '------------------------
    ' Create the table
    ' ================
    ' (O've got no error checking here...)
    SQLite3ExecuteNonQuery myDbHandle, "CREATE TABLE MyBigTable (TheId INTEGER, TheDate REAL, TheText TEXT, TheValue REAL)"
    
    '---------------------------
    ' Add an index
    ' ================
    SQLite3ExecuteNonQuery myDbHandle, "CREATE INDEX idx_MyBigTable_Id_Date ON MyBigTable (TheId, TheDate)"
    
    ' START Insert Time
    testStart = Now()
    
    '-------------------
    ' Begin transaction
    '==================
    SQLite3ExecuteNonQuery myDbHandle, "BEGIN TRANSACTION"

    '-------------------------
    ' Prepare an insert statement with parameters
    ' ===============
    ' Create the sql statement - getting a StmtHandle back
    RetVal = SQLite3PrepareV2(myDbHandle, "INSERT INTO MyBigTable Values (?, ?, ?, ?)", myStmtHandle)
    If RetVal <> SQLITE_OK Then
        Debug.Print "SQLite3PrepareV2 returned " & SQLite3ErrMsg(myDbHandle)
        Beep
    End If
    
    PrintParameters myStmtHandle
        
    Randomize
    startDate = DateValue("1 Jan 2000")
    
    For i = 1 To 100000
        curDate = startDate + i
        curValue = Rnd() * 1000
        
        RetVal = SQLite3BindInt32(myStmtHandle, 1, 42000 + i)
        If RetVal <> SQLITE_OK Then
            Debug.Print "SQLite3Bind returned " & RetVal, SQLite3ErrMsg(myDbHandle)
            Beep
        End If
        
        RetVal = SQLite3BindDate(myStmtHandle, 2, curDate)
        If RetVal <> SQLITE_OK Then
            Debug.Print "SQLite3Bind returned " & RetVal, SQLite3ErrMsg(myDbHandle)
            Beep
        End If
        
        RetVal = SQLite3BindText(myStmtHandle, 3, "The quick brown fox jumped over the lazy dog.")
        If RetVal <> SQLITE_OK Then
            Debug.Print "SQLite3Bind returned " & RetVal, SQLite3ErrMsg(myDbHandle)
            Beep
        End If
        
        RetVal = SQLite3BindDouble(myStmtHandle, 4, curValue)
        If RetVal <> SQLITE_OK Then
            Debug.Print "SQLite3Bind returned " & RetVal, SQLite3ErrMsg(myDbHandle)
            Beep
        End If
        
        RetVal = SQLite3Step(myStmtHandle)
        If RetVal <> SQLITE_DONE Then
            Debug.Print "SQLite3Step returned " & RetVal, SQLite3ErrMsg(myDbHandle)
            Beep
        End If
    
        RetVal = SQLite3Reset(myStmtHandle)
        If RetVal <> SQLITE_OK Then
            Debug.Print "SQLite3Reset returned " & RetVal, SQLite3ErrMsg(myDbHandle)
            Beep
        End If
    Next
    
    ' Finalize (delete) the statement
    RetVal = SQLite3Finalize(myStmtHandle)
    Debug.Print "SQLite3Finalize returned " & RetVal

    '-------------------
    ' Commit transaction
    '==================
    SQLite3ExecuteNonQuery myDbHandle, "COMMIT TRANSACTION"

    ' STOP Insert Time
    Debug.Print "Insert Elapsed: " & Format(Now() - testStart, "HH:mm:ss")

    ' START Select  Time
    testStart = Now()

    '-------------------------
    ' Select statement
    ' ===============
    ' Create the sql statement - getting a StmtHandle back
    ' Now using named parameters!
    RetVal = SQLite3PrepareV2(myDbHandle, "SELECT TheId, datetime(TheDate), TheText, TheValue FROM MyBigTable WHERE TheId = @FindThisId AND TheDate <= julianday(@FindThisDate) LIMIT 1", myStmtHandle)
    Debug.Print "SQLite3PrepareV2 returned " & RetVal
    
    PrintParameters myStmtHandle

    paramIndexId = SQLite3BindParameterIndex(myStmtHandle, "@FindThisId")
    If paramIndexId = 0 Then
        Debug.Print "SQLite3BindParameterIndex could not find the Id parameter!"
        Beep
    End If
    
    paramIndexDate = SQLite3BindParameterIndex(myStmtHandle, "@FindThisDate")
    If paramIndexDate = 0 Then
        Debug.Print "SQLite3BindParameterIndex could not find the Date parameter!"
        Beep
    End If
    
    startDate = DateValue("1 Jan 2000")
    
    For i = 1 To 10000
        offset = i Mod 1000
        ' Bind the parameters
        RetVal = SQLite3BindInt32(myStmtHandle, paramIndexId, 4200 + 500 + offset)
        If RetVal <> SQLITE_OK Then
            Debug.Print "SQLite3Bind returned " & RetVal, SQLite3ErrMsg(myDbHandle)
            Beep
        End If
    
        RetVal = SQLite3BindText(myStmtHandle, paramIndexDate, Format(startDate + 500 + offset, "yyyy-MM-dd HH:mm:ss"))
        If RetVal <> SQLITE_OK Then
            Debug.Print "SQLite3Bind returned " & RetVal, SQLite3ErrMsg(myDbHandle)
            Beep
        End If
        
        RetVal = SQLite3Step(myStmtHandle)
        If RetVal = SQLITE_ROW Then
            ' We have access to the result columns here.
            If offset = 1 Then
                Debug.Print "At row " & i
                Debug.Print "------------"
                PrintColumns myStmtHandle
                Debug.Print "============"
            End If
        ElseIf RetVal = SQLITE_DONE Then
            Debug.Print "No row found"
        End If
    
        RetVal = SQLite3Reset(myStmtHandle)
        If RetVal <> SQLITE_OK Then
            Debug.Print "SQLite3Reset returned " & RetVal, SQLite3ErrMsg(myDbHandle)
            Beep
        End If
    Next
        
    ' Finalize (delete) the statement
    RetVal = SQLite3Finalize(myStmtHandle)
    Debug.Print "SQLite3Finalize returned " & RetVal
    
    ' STOP Select time
    Debug.Print "Select Elapsed: " & Format(Now() - testStart, "HH:mm:ss")
    
    ' Close the database
    RetVal = SQLite3Close(myDbHandle)
    Kill testFile

    Debug.Print "----- TestBinding End -----"
End Sub

Public Sub TestDates()
    Dim testFile As String
    
    #If Win64 Then
    Dim myDbHandle As LongPtr
    Dim myStmtHandle As LongPtr
    #Else
    Dim myDbHandle As Long
    Dim myStmtHandle As Long
    #End If

    Dim RetVal As Long
    Dim stepMsg As String
    Dim i As Long
    
    Dim myDate As Date
    Dim myEvent As String
    
    Debug.Print "----- TestDates Start -----"
    
    ' Open the database - getting a DbHandle back
    testFile = "C:\TestSqlite3ForExcel.db3"
    RetVal = SQLite3Open(testFile, myDbHandle)
    Debug.Print "SQLite3Open returned " & RetVal
    
    '------------------------
    ' Create the table
    ' ================
    ' (I've got no error checking here...)
    SQLite3ExecuteNonQuery myDbHandle, "CREATE TABLE MyDateTable (MyDate REAL, MyEvent TEXT)"
    
    '-------------------------
    ' Prepare an insert statement with parameters
    ' ===============
    ' Create the sql statement - getting a StmtHandle back
    RetVal = SQLite3PrepareV2(myDbHandle, "INSERT INTO MyDateTable Values (@SomeDate, @SomeEvent)", myStmtHandle)
    If RetVal <> SQLITE_OK Then
        Debug.Print "SQLite3PrepareV2 returned " & SQLite3ErrMsg(myDbHandle)
        Beep
    End If
    
    RetVal = SQLite3BindDate(myStmtHandle, 1, DateSerial(2010, 6, 19))
    If RetVal <> SQLITE_OK Then
        Debug.Print "SQLite3Bind returned " & RetVal, SQLite3ErrMsg(myDbHandle)
        Beep
    End If
    
    RetVal = SQLite3BindText(myStmtHandle, 2, "Nice trip somewhere")
    If RetVal <> SQLITE_OK Then
        Debug.Print "SQLite3Bind returned " & RetVal, SQLite3ErrMsg(myDbHandle)
        Beep
    End If
    
    RetVal = SQLite3Step(myStmtHandle)
    If RetVal <> SQLITE_DONE Then
        Debug.Print "SQLite3Step returned " & RetVal, SQLite3ErrMsg(myDbHandle)
        Beep
    End If
    
    ' Finalize the statement
    RetVal = SQLite3Finalize(myStmtHandle)
    Debug.Print "SQLite3Finalize returned " & RetVal

    '-------------------------
    ' Select statement
    ' ===============
    ' Create the sql statement - getting a StmtHandle back
    ' Now using named parameters!
    RetVal = SQLite3PrepareV2(myDbHandle, "SELECT * FROM MyDateTable", myStmtHandle)
    Debug.Print "SQLite3PrepareV2 returned " & RetVal
    
    RetVal = SQLite3Step(myStmtHandle)
    If RetVal = SQLITE_ROW Then
        ' We have access to the result columns here.
        myDate = SQLite3ColumnDate(myStmtHandle, 0)
        myEvent = SQLite3ColumnText(myStmtHandle, 1)
        Debug.Print "Event: " & myEvent, "Date: " & myDate
    ElseIf RetVal = SQLITE_DONE Then
        Debug.Print "No row found"
    End If
        
    ' Finalize (delete) the statement
    RetVal = SQLite3Finalize(myStmtHandle)
    Debug.Print "SQLite3Finalize returned " & RetVal
    
    ' Close the database
    RetVal = SQLite3Close(myDbHandle)
    Kill testFile

    Debug.Print "----- TestDates End -----"
End Sub


Public Sub TestStrings()
    Dim testFile As String
    #If Win64 Then
    Dim myDbHandle As LongPtr
    Dim myStmtHandle As LongPtr
    #Else
    Dim myDbHandle As Long
    Dim myStmtHandle As Long
    #End If
    Dim RetVal As Long
    Dim stepMsg As String
    Dim i As Long
    
    Dim myString1 As String
    Dim myString2 As String
    Dim myLongString As String
    Dim myStringResult As String
    
    Debug.Print "----- TestStrings Start -----"
    
    ' Open the database - getting a DbHandle back
    testFile = "C:\TestSqlite3ForExcel.db3"
    RetVal = SQLite3Open(testFile, myDbHandle)
    Debug.Print "SQLite3Open returned " & RetVal
    
    myString2 = ""
    myLongString = String(10000, "A")
    
    '------------------------
    ' Create the table
    ' ================
    ' (I've got no error checking here...)
    SQLite3ExecuteNonQuery myDbHandle, "CREATE TABLE MyStringTable (MyValue TEXT)"
    
    '-------------------------
    ' Prepare an insert statement with parameters
    ' ===============
    ' Create the sql statement - getting a StmtHandle back
    RetVal = SQLite3PrepareV2(myDbHandle, "INSERT INTO MyStringTable Values (@SomeString)", myStmtHandle)
    If RetVal <> SQLITE_OK Then
        Debug.Print "SQLite3PrepareV2 returned " & SQLite3ErrMsg(myDbHandle)
        Beep
    End If
    
    RetVal = SQLite3BindText(myStmtHandle, 1, myString1)
    If RetVal <> SQLITE_OK Then
        Debug.Print "SQLite3Bind returned " & RetVal, SQLite3ErrMsg(myDbHandle)
        Beep
    End If
    
    RetVal = SQLite3Step(myStmtHandle)
    If RetVal <> SQLITE_DONE Then
        Debug.Print "SQLite3Step returned " & RetVal, SQLite3ErrMsg(myDbHandle)
        Beep
    End If
    
    RetVal = SQLite3Reset(myStmtHandle)
    If RetVal <> SQLITE_OK Then
        Debug.Print "SQLite3Reset returned " & RetVal, SQLite3ErrMsg(myDbHandle)
        Beep
    End If
    
    RetVal = SQLite3BindText(myStmtHandle, 1, myString2)
    If RetVal <> SQLITE_OK Then
        Debug.Print "SQLite3Bind returned " & RetVal, SQLite3ErrMsg(myDbHandle)
        Beep
    End If
    
    RetVal = SQLite3Step(myStmtHandle)
    If RetVal <> SQLITE_DONE Then
        Debug.Print "SQLite3Step returned " & RetVal, SQLite3ErrMsg(myDbHandle)
        Beep
    End If
    
    RetVal = SQLite3Reset(myStmtHandle)
    If RetVal <> SQLITE_OK Then
        Debug.Print "SQLite3Reset returned " & RetVal, SQLite3ErrMsg(myDbHandle)
        Beep
    End If
    
    RetVal = SQLite3BindText(myStmtHandle, 1, myLongString)
    If RetVal <> SQLITE_OK Then
        Debug.Print "SQLite3Bind returned " & RetVal, SQLite3ErrMsg(myDbHandle)
        Beep
    End If
    
    RetVal = SQLite3Step(myStmtHandle)
    If RetVal <> SQLITE_DONE Then
        Debug.Print "SQLite3Step returned " & RetVal, SQLite3ErrMsg(myDbHandle)
        Beep
    End If
    ' Finalize the statement
    RetVal = SQLite3Finalize(myStmtHandle)
    Debug.Print "SQLite3Finalize returned " & RetVal

    '-------------------------
    ' Select statement
    ' ===============
    ' Create the sql statement - getting a StmtHandle back
    ' Now using named parameters!
    RetVal = SQLite3PrepareV2(myDbHandle, "SELECT * FROM MyStringTable", myStmtHandle)
    Debug.Print "SQLite3PrepareV2 returned " & RetVal
    
    RetVal = SQLite3Step(myStmtHandle)
    If RetVal = SQLITE_ROW Then
        ' We have access to the result columns here.
        myStringResult = SQLite3ColumnText(myStmtHandle, 0)
        Debug.Print "Result1: " + myStringResult
    ElseIf RetVal = SQLITE_DONE Then
        Debug.Print "No row found"
    End If
        
    RetVal = SQLite3Step(myStmtHandle)
    If RetVal = SQLITE_ROW Then
        ' We have access to the result columns here.
        myStringResult = SQLite3ColumnText(myStmtHandle, 0)
        Debug.Print "Result2: " + myStringResult
    ElseIf RetVal = SQLITE_DONE Then
        Debug.Print "No row found"
    End If
        
    RetVal = SQLite3Step(myStmtHandle)
    If RetVal = SQLITE_ROW Then
        ' We have access to the result columns here.
        myStringResult = SQLite3ColumnText(myStmtHandle, 0)
        
        Debug.Print "Long String is the same: " & (myStringResult = myLongString)
    ElseIf RetVal = SQLITE_DONE Then
        Debug.Print "No row found"
    End If
    
    ' Finalize (delete) the statement
    RetVal = SQLite3Finalize(myStmtHandle)
    Debug.Print "SQLite3Finalize returned " & RetVal
    
    ' Close the database
    RetVal = SQLite3Close(myDbHandle)
    Kill testFile

    Debug.Print "----- TestStrings End -----"
End Sub

Public Sub TestBackup()
    Dim testFile As String
    Dim testFileBackup As String
    
    #If Win64 Then
    Dim myDbHandle As LongPtr
    Dim myDbBackupHandle As LongPtr
    Dim myBackupHandle As LongPtr
    #Else
    Dim myDbHandle As Long
    Dim myDbBackupHandle As Long
    Dim myBackupHandle As Long
    #End If
   
    Dim RetVal As Long
    Dim i As Long
    
    Debug.Print "----- TestBackup Start -----"
    
    ' Open the database - getting a DbHandle back
    testFile = "C:\TestSqlite3ForExcel.db3"
    RetVal = SQLite3Open(testFile, myDbHandle)
    Debug.Print "SQLite3Open returned " & RetVal
    
    SQLite3ExecuteNonQuery myDbHandle, "CREATE TABLE MyTestTable (Key INT PRIMARY KEY, Value TEXT)"
    SQLite3ExecuteNonQuery myDbHandle, "INSERT INTO MyTestTable VALUES (1, 'First')"
    SQLite3ExecuteNonQuery myDbHandle, "INSERT INTO MyTestTable VALUES (2, 'Second')"
    SQLite3ExecuteQuery myDbHandle, "SELECT * FROM MyTestTable"
    
    ' Now do a backup
    testFileBackup = "C:\TestSqlite3ForExcel_Backup.db3"
    RetVal = SQLite3Open(testFileBackup, myDbBackupHandle)
    Debug.Print "SQLite3Open returned " & RetVal
    
    myBackupHandle = SQLite3BackupInit(myDbBackupHandle, "main", myDbHandle, "main")
    If myBackupHandle <> 0 Then
        RetVal = SQLite3BackupStep(myBackupHandle, -1)
        Debug.Print "SQLite3BackupStep returned " & RetVal
        RetVal = SQLite3BackupFinish(myBackupHandle)
        Debug.Print "SQLite3BackupFinish returned " & RetVal
    End If
    RetVal = SQLite3ErrCode(myDbBackupHandle)
    Debug.Print "Backup result " & RetVal
    Debug.Print "Selecting from backup:"
    SQLite3ExecuteQuery myDbBackupHandle, "SELECT * FROM MyTestTable"
    
    RetVal = SQLite3Close(myDbHandle)
    RetVal = SQLite3Close(myDbBackupHandle)
    
    Kill testFile
    Kill testFileBackup
    
    Debug.Print "----- TestBackup End -----"
End Sub


Public Sub TestBlob()
    Dim testFile As String
    
    #If Win64 Then
    Dim myDbHandle As LongPtr
    Dim myStmtHandle As LongPtr
    #Else
    Dim myDbHandle As Long
    Dim myStmtHandle As Long
    #End If
    Dim RetVal As Long
    Dim stepMsg As String
    Dim i As Long
    
    Dim myBlob(2) As Byte
    Dim myBlobResult() As Byte
    
    Debug.Print "----- TestBlob Start -----"
    
    ' Open the database - getting a DbHandle back
    testFile = "C:\TestSqlite3ForExcel.db3"
    RetVal = SQLite3Open(testFile, myDbHandle)
    Debug.Print "SQLite3Open returned " & RetVal
    
    myBlob(0) = 90
    myBlob(1) = 91
    myBlob(2) = 92
    
    '------------------------
    ' Create the table
    ' ================
    ' (I've got no error checking here...)
    SQLite3ExecuteNonQuery myDbHandle, "CREATE TABLE MyBlobTable (MyValue BLOB)"
    
    '-------------------------
    ' Prepare an insert statement with parameters
    ' ===============
    ' Create the sql statement - getting a StmtHandle back
    RetVal = SQLite3PrepareV2(myDbHandle, "INSERT INTO MyBlobTable Values (@SomeString)", myStmtHandle)
    If RetVal <> SQLITE_OK Then
        Debug.Print "SQLite3PrepareV2 returned " & SQLite3ErrMsg(myDbHandle)
        Beep
    End If
    
    RetVal = SQLite3BindBlob(myStmtHandle, 1, myBlob)
    If RetVal <> SQLITE_OK Then
        Debug.Print "SQLite3Bind returned " & RetVal, SQLite3ErrMsg(myDbHandle)
        Beep
    End If
    
    RetVal = SQLite3Step(myStmtHandle)
    If RetVal <> SQLITE_DONE Then
        Debug.Print "SQLite3Step returned " & RetVal, SQLite3ErrMsg(myDbHandle)
        Beep
    End If
    
    ' Finalize the statement
    RetVal = SQLite3Finalize(myStmtHandle)
    Debug.Print "SQLite3Finalize returned " & RetVal

    '-------------------------
    ' Select statement
    ' ===============
    ' Create the sql statement - getting a StmtHandle back
    ' Now using named parameters!
    RetVal = SQLite3PrepareV2(myDbHandle, "SELECT * FROM MyBlobTable", myStmtHandle)
    Debug.Print "SQLite3PrepareV2 returned " & RetVal
    
    RetVal = SQLite3Step(myStmtHandle)
    If RetVal = SQLITE_ROW Then
        ' We have access to the result columns here.
        myBlobResult = SQLite3ColumnBlob(myStmtHandle, 0)
        For i = LBound(myBlobResult) To UBound(myBlobResult)
            Debug.Print "Blob byte " & i & ": " & myBlobResult(i)
        Next
    ElseIf RetVal = SQLITE_DONE Then
        Debug.Print "No row found"
    End If
    
    ' Finalize (delete) the statement
    RetVal = SQLite3Finalize(myStmtHandle)
    Debug.Print "SQLite3Finalize returned " & RetVal
    
    ' Close the database
    RetVal = SQLite3Close(myDbHandle)
    Kill testFile

    Debug.Print "----- TestBlob End -----"
End Sub

Public Sub TestWriteReadOnly()
    Dim testFile As String
    #If Win64 Then
    Dim myDbHandle As LongPtr
    Dim myDbHandleV2 As LongPtr
    Dim myStmtHandle As LongPtr
    #Else
    Dim myDbHandle As Long
    Dim myDbHandleV2 As Long
    Dim myStmtHandle As Long
    #End If
    Dim RetVal As Long
    
    ' Open the database in Read Write Access
    testFile = "C:\TestSqlite3ForExcel.db3"
    RetVal = SQLite3Open(testFile, myDbHandle)
    Debug.Print "SQLite3Open returned " & RetVal
    
    ' Open the database in Read Only Access
    testFile = "C:\TestSqlite3ForExcel.db3"
    RetVal = SQLite3OpenV2(testFile, myDbHandleV2, SQLITE_OPEN_READONLY, Empty)
    Debug.Print "SQLite3OpenV2 returned " & RetVal
    
    ' Create the sql statement - getting a StmtHandle back
    RetVal = SQLite3PrepareV2(myDbHandle, "CREATE TABLE MyFirstTable (TheId INTEGER, TheText TEXT, TheValue REAL)", myStmtHandle)
    Debug.Print "SQLite3PrepareV2 returned " & RetVal
    
    ' Start running the statement
    RetVal = SQLite3Step(myStmtHandle)
    Debug.Print "SQLite3Step returned " & RetVal
    
    ' Finalize (delete) the statement
    RetVal = SQLite3Finalize(myStmtHandle)
    Debug.Print "SQLite3Finalize returned " & RetVal
    
    ' Create the sql statement - getting a StmtHandle back with Read Only
    RetVal = SQLite3PrepareV2(myDbHandleV2, "CREATE TABLE MySecondTable (TheId INTEGER, TheText TEXT, TheValue REAL)", myStmtHandle)
    'RetVal = SQLite3PrepareV2(myDbHandleV2, "SELECT * FROM MyFirstTable", myStmtHandle)
    Debug.Print "SQLite3PrepareV2 returned " & RetVal
    
    ' Start running the statement with Read Only
    RetVal = SQLite3Step(myStmtHandle)
    Debug.Print "SQLite3Step returned " & RetVal
    
    If RetVal = SQLITE_READONLY Then
        Debug.Print "Cannot Write in Read Only database"
    End If
    
    ' Finalize (delete) the statement with Read Only
    RetVal = SQLite3Finalize(myStmtHandle)
    Debug.Print "SQLite3Finalize returned " & RetVal
    
    ' Create the sql statement - getting a StmtHandle back with Read Only
    RetVal = SQLite3PrepareV2(myDbHandleV2, "SELECT * FROM MyFirstTable", myStmtHandle)
    Debug.Print "SQLite3PrepareV2 returned " & RetVal
    
    ' Start running the statement with Read Only
    RetVal = SQLite3Step(myStmtHandle)
    Debug.Print "SQLite3Step returned " & RetVal
        
    If RetVal = SQLITE_DONE Then
        Debug.Print "But Reading is granted on Read Only database"
    End If
    
    ' Finalize (delete) the statement with Read Only
    RetVal = SQLite3Finalize(myStmtHandle)
    Debug.Print "SQLite3Finalize returned " & RetVal
    
    RetVal = SQLite3Close(myDbHandleV2)
    Debug.Print "SQLite3Close V2 returned " & RetVal
    
    RetVal = SQLite3Close(myDbHandle)
    Debug.Print "SQLite3Close returned " & RetVal
    
    Kill testFile

End Sub

' SQLite3 Helper Functions
#If Win64 Then
Public Function SQLite3ExecuteNonQuery(ByVal dbHandle As LongPtr, ByVal SqlCommand As String) As Long
    Dim stmtHandle As LongPtr
#Else
Public Function SQLite3ExecuteNonQuery(ByVal dbHandle As Long, ByVal SqlCommand As String) As Long
    Dim stmtHandle As Long
#End If
    
    SQLite3PrepareV2 dbHandle, SqlCommand, stmtHandle
    SQLite3Step stmtHandle
    SQLite3Finalize stmtHandle
    
    SQLite3ExecuteNonQuery = SQLite3Changes(dbHandle)
End Function

#If Win64 Then
Public Sub SQLite3ExecuteQuery(ByVal dbHandle As LongPtr, ByVal sqlQuery As String)
    Dim stmtHandle As LongPtr
#Else
Public Sub SQLite3ExecuteQuery(ByVal dbHandle As Long, ByVal sqlQuery As String)
    Dim stmtHandle As Long
#End If
    ' Dumps a query to the debug window. No error checking
    
    Dim RetVal As Long

    RetVal = SQLite3PrepareV2(dbHandle, sqlQuery, stmtHandle)
    Debug.Print "SQLite3PrepareV2 returned " & RetVal
    
    ' Start running the statement
    RetVal = SQLite3Step(stmtHandle)
    If RetVal = SQLITE_ROW Then
        Debug.Print "SQLite3Step Row Ready"
        PrintColumns stmtHandle
    Else
        Debug.Print "SQLite3Step returned " & RetVal
    End If
    
    ' Move to next row
    RetVal = SQLite3Step(stmtHandle)
    Do While RetVal = SQLITE_ROW
        Debug.Print "SQLite3Step Row Ready"
        PrintColumns stmtHandle
        RetVal = SQLite3Step(stmtHandle)
    Loop

    If RetVal = SQLITE_DONE Then
        Debug.Print "SQLite3Step Done"
    Else
        Debug.Print "SQLite3Step returned " & RetVal
    End If
    
    ' Finalize (delete) the statement
    RetVal = SQLite3Finalize(stmtHandle)
    Debug.Print "SQLite3Finalize returned " & RetVal
End Sub
