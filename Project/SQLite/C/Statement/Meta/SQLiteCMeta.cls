VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLiteCMeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Retrieves prepared statement metadata."
'@Folder "SQLite.C.Statement.Meta"
'@ModuleDescription "Retrieves prepared statement metadata."
'@PredeclaredId
'@Exposed
Option Explicit

#If VBA7 Then
' 0-based indices
Private Declare PtrSafe Function sqlite3_column_count Lib "SQLite3" (ByVal hStmt As LongPtr) As Long
Private Declare PtrSafe Function sqlite3_column_name Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrString
Private Declare PtrSafe Function sqlite3_column_database_name Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrString
Private Declare PtrSafe Function sqlite3_column_table_name Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrString
Private Declare PtrSafe Function sqlite3_column_origin_name Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrString
Private Declare PtrSafe Function sqlite3_column_decltype Lib "SQLite3" (ByVal hStmt As LongPtr, ByVal iCol As Long) As LongPtr ' PtrString
Private Declare PtrSafe Function sqlite3_table_column_metadata Lib "SQLite3" ( _
    ByVal hDb As LongPtr, ByVal zDbName As LongPtr, ByVal zTableName As LongPtr, ByVal zColumnName As LongPtr, _
    ByRef pzDataType As LongPtr, ByRef pzCollSeq As LongPtr, ByRef pNotNull As Long, _
    ByRef pPrimaryKey As Long, ByRef pAutoinc As Long) As SQLiteResultCodes
#Else
' 0-based indices
Private Declare Function sqlite3_column_count Lib "SQLite3" (ByVal hStmt As Long) As Long
Private Declare Function sqlite3_column_name Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrString
Private Declare Function sqlite3_column_database_name Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrString
Private Declare Function sqlite3_column_table_name Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrString
Private Declare Function sqlite3_column_origin_name Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrString
Private Declare Function sqlite3_column_decltype Lib "SQLite3" (ByVal hStmt As Long, ByVal iCol As Long) As Long ' PtrString
Private Declare Function sqlite3_table_column_metadata Lib "SQLite3" ( _
    ByVal hDb As Long, ByVal zDbName As Long, ByVal zTableName As Long, ByVal zColumnName As Long, _
    ByRef pzDataType As Long, ByRef pzCollSeq As Long, ByRef pNotNull As Long, _
    ByRef pPrimaryKey As Long, ByRef pAutoinc As Long) As SQLiteResultCodes
#End If

Private Type TSQLiteCMeta
    DbStmt As SQLiteCStatement
    TableMeta() As SQLiteCColumnMeta
    ColumnMap As Scripting.Dictionary
    ColumnCount As Long
    AffinityMap() As Variant ' SQLiteType
End Type
Private this As TSQLiteCMeta


'@DefaultMember
Public Function Create(ByVal DbStmt As SQLiteCStatement) As SQLiteCMeta
Attribute Create.VB_UserMemId = 0
    Dim Instance As SQLiteCMeta
    Set Instance = New SQLiteCMeta
    Instance.Init DbStmt
    Set Create = Instance
End Function


Friend Sub Init(ByVal DbStmt As SQLiteCStatement)
    Guard.NullReference DbStmt
    With this
        Set .DbStmt = DbStmt
        Set .ColumnMap = New Scripting.Dictionary
        .ColumnMap.CompareMode = TextCompare
        .AffinityMap = Array(SQLITE_NONE, SQLITE_BLOB, SQLITE_TEXT, _
                             SQLITE_TEXT, SQLITE_INTEGER, SQLITE_FLOAT)
    End With
End Sub


Public Property Get TableMeta() As SQLiteCColumnMeta()
    TableMeta = this.TableMeta
End Property


Friend Property Get AffinityMap() As Variant
    AffinityMap = this.AffinityMap
End Property


Public Function SQLiteTypeName(ByVal DataTypeCode As SQLiteDataType) As String
    SQLiteTypeName = Array("NONE", "INTEGER", "FLOAT", _
                           "TEXT", "BLOB", "NULL")(DataTypeCode)
End Function


Public Function SQLiteTypeAffinityName(ByVal DataTypeAffinityCode As SQLiteTypeAffinity) As String
    SQLiteTypeAffinityName = Array("NONE", "BLOB", "TEXT", "NUMERIC", _
                                   "INTEGER", "REAL")(DataTypeAffinityCode - SQLITE_AFF_NONE)
End Function


'''' Reference: https://www.sqlite.org/c3ref/column_count.html
''''
'@Description "Retrieves number of columns returned by the prepared statement."
Public Function ColumnCountAPI() As Long
Attribute ColumnCountAPI.VB_Description = "Retrieves number of columns returned by the prepared statement."
    FixGuard.StmtNotPrepared this.DbStmt, "SQLiteCMeta/ColumnCountAPI"
    this.ColumnCount = sqlite3_column_count(this.DbStmt.StmtHandle)
    ColumnCountAPI = this.ColumnCount
End Function


'''' Reference: https://www.sqlite.org/datatype3.html
''''            See section 3.1
''''
'@Description "Determines column type affinity."
Friend Function TypeAffinityFromDeclaredType(ByVal DeclaredType As String) As SQLiteTypeAffinity
Attribute TypeAffinityFromDeclaredType.VB_Description = "Determines column type affinity."
    Dim DeclType As String
    DeclType = UCase$(DeclaredType)
    If InStr(DeclType, "INT") > 0 Then
        TypeAffinityFromDeclaredType = SQLITE_AFF_INTEGER
    ElseIf InStr(DeclType, "CHAR") > 0 Or _
           InStr(DeclType, "CLOB") > 0 Or _
           InStr(DeclType, "TEXT") > 0 Then
        TypeAffinityFromDeclaredType = SQLITE_AFF_TEXT
    ElseIf InStr(DeclType, "BLOB") > 0 Or Len(DeclType) = 0 Then
        TypeAffinityFromDeclaredType = SQLITE_AFF_BLOB
    ElseIf InStr(DeclType, "REAL") > 0 Or _
           InStr(DeclType, "FLOA") > 0 Or _
           InStr(DeclType, "DOUB") > 0 Then
        TypeAffinityFromDeclaredType = SQLITE_AFF_REAL
    ElseIf Len(DeclType) = 0 Then
        TypeAffinityFromDeclaredType = SQLITE_AFF_NONE
    Else
        TypeAffinityFromDeclaredType = SQLITE_AFF_NUMERIC
    End If
End Function


'''' Reference: https://www.sqlite.org/c3ref/column_name.html
'''' Reference: https://www.sqlite.org/c3ref/column_database_name.html
'''' Reference: https://www.sqlite.org/c3ref/column_decltype.html
'''' Reference: https://www.sqlite.org/c3ref/column_blob.html
'''' Reference: https://www.sqlite.org/c3ref/c_blob.html
'''' Reference: https://www.sqlite.org/c3ref/table_column_metadata.html
''''
'''' For calculated columns, only the column_name API should work. Other APIs
'''' should return null pointer, without indicating an error otherwise.
''''
'@Description "Retrieves columns metadata."
Friend Function ColumnMetaAPI(ByRef ColumnInfo As SQLiteCColumnMeta _
                             ) As SQLiteResultCodes
Attribute ColumnMetaAPI.VB_Description = "Retrieves columns metadata."
    Dim dbs As SQLiteCStatement
    Set dbs = this.DbStmt
    Dim dbc As SQLiteCConnection
    Set dbc = dbs.DbConnection
    FixGuard.StmtNotPrepared dbs, "SQLiteCMeta/ColumnMetaAPI"
    
    #If VBA7 Then
        Dim hStmt As LongPtr
    #Else
        Dim hStmt As Long
    #End If
    hStmt = dbs.StmtHandle
    
    Dim ResultCode As SQLiteResultCodes
    '@Ignore AssignmentNotUsed: False positive
    ResultCode = SQLITE_OK
    With ColumnInfo
        If .Initialized <> -1 Then Err.Raise ErrNo.InvalidParameterErr, _
            "SQLiteCExecSQL", "Set .ColumnIndex and .Initialized=-1 before the call!"

        Dim ColumnIndex As Long
        ColumnIndex = .ColumnIndex
        
        Dim ColumnNamePtr As Variant
        ColumnNamePtr = sqlite3_column_name(hStmt, ColumnIndex)
        If ColumnNamePtr = 0 Then
            ColumnMetaAPI = dbc.ErrInfoRetrieve
            Exit Function
        End If
        Dim AliasName As String
        AliasName = UTFlib.StrFromUTF8Ptr(ColumnNamePtr)
        .Name = AliasName
        '@Ignore IndexedDefaultMemberAccess
        this.ColumnMap(.Name) = ColumnIndex
        AliasName = Replace(LCase$(AliasName), "_", vbNullString)
        .RowId = (AliasName = "rowid" Or AliasName = "oid")
        
        Dim DatabaseNamePtr As Variant
        DatabaseNamePtr = sqlite3_column_database_name(hStmt, ColumnIndex)
        If DatabaseNamePtr <> 0 Then .DbName = UTFlib.StrFromUTF8Ptr(DatabaseNamePtr)
        
        Dim TableNamePtr As Variant
        TableNamePtr = sqlite3_column_table_name(hStmt, ColumnIndex)
        If TableNamePtr <> 0 Then .TableName = UTFlib.StrFromUTF8Ptr(TableNamePtr)
        
        ColumnNamePtr = sqlite3_column_origin_name(hStmt, ColumnIndex)
        If ColumnNamePtr <> 0 Then .OriginName = UTFlib.StrFromUTF8Ptr(ColumnNamePtr)
        
        Dim DataTypePtr As Variant
        DataTypePtr = sqlite3_column_decltype(hStmt, ColumnIndex)
        If DataTypePtr <> 0 Then
            .DeclaredTypeC = UTFlib.StrFromUTF8Ptr(DataTypePtr)
            .Affinity = TypeAffinityFromDeclaredType(.DeclaredTypeC)
            .AffinityType = this.AffinityMap(.Affinity - SQLITE_AFF_NONE)
            .TableMeta = True
        Else
            '''' SQLITE_RANGE is expected for a calculated column, such as "count(*)"
            '''' Only return an error, if this is not the case; otherwise, continue.
            ResultCode = dbc.ErrInfoRetrieve
            If ResultCode <> SQLITE_OK And ResultCode <> SQLITE_RANGE Then
                ColumnMetaAPI = ResultCode
                Exit Function
            Else
                ResultCode = SQLITE_OK
            End If
            .Affinity = SQLITE_AFF_NONE
            .AffinityType = this.AffinityMap(.Affinity - SQLITE_AFF_NONE)
            .TableMeta = False
        End If

        '''' For queries against PRAGMA functions, table_column_metadata
        '''' should return SQLITE_ERROR (no such table column...).
        ''''
        '''' Make sure that relevant previous APIs did not return null pointers.
        ''''
        If DatabaseNamePtr <> 0 And TableNamePtr <> 0 And _
           ColumnNamePtr <> 0 And DataTypePtr <> 0 Then
            Dim NotNull As Long
            NotNull = False
            Dim PrimaryKey As Long
            PrimaryKey = False
            Dim AutoIncrement As Long
            AutoIncrement = False
            Dim CollationPtr As Variant: CollationPtr = 0 '''' RD workaround
            ResultCode = sqlite3_table_column_metadata(dbc.DbHandle, _
                DatabaseNamePtr, TableNamePtr, ColumnNamePtr, _
                DataTypePtr, CollationPtr, NotNull, PrimaryKey, AutoIncrement)
            If ResultCode = SQLITE_OK And DataTypePtr <> 0 And CollationPtr <> 0 Then
                .DeclaredTypeT = UTFlib.StrFromUTF8Ptr(DataTypePtr)
                .Collation = UTFlib.StrFromUTF8Ptr(CollationPtr)
                .NotNull = CBool(NotNull)
                .PrimaryKey = CBool(PrimaryKey)
                .AutoIncrement = CBool(AutoIncrement)
            Else
                ColumnMetaAPI = dbc.ErrInfoRetrieve
                Exit Function
            End If
        End If
    End With
    
    ColumnMetaAPI = ResultCode
End Function


'@Description "Collects table metadata."
Public Function TableMetaCollect() As SQLiteResultCodes
Attribute TableMetaCollect.VB_Description = "Collects table metadata."
    Dim ColumnCount As Long
    ColumnCount = ColumnCountAPI
    If ColumnCount = 0 Then
        Debug.Print "No data avaialable"
        TableMetaCollect = SQLITE_ERROR
        Exit Function
    End If
    ReDim this.TableMeta(0 To ColumnCount - 1)
    Dim ResultCode As SQLiteResultCodes
    Dim ColumnIndex As Long
    For ColumnIndex = 0 To ColumnCount - 1
        this.TableMeta(ColumnIndex).ColumnIndex = ColumnIndex
        this.TableMeta(ColumnIndex).Initialized = -1
        ResultCode = ColumnMetaAPI(this.TableMeta(ColumnIndex))
        If ResultCode <> SQLITE_OK Then
            TableMetaCollect = ResultCode
            Exit Function
        End If
    Next ColumnIndex
    
    With this.DbStmt.DbExecutor
        .TableMeta = this.TableMeta
        .MetaLoaded = True
    End With
    TableMetaCollect = ResultCode
End Function
