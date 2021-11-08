VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLlib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Provides SQL queries and templates"
'@Folder "SQLite.SQLlib"
'@ModuleDescription "Provides SQL queries and templates"
'@PredeclaredId
'@Exposed
'@IgnoreModule IndexedDefaultMemberAccess
Option Explicit

Public Enum FieldFormat
    CastAsIs
    CastIdAsText
    CastAllAsText
End Enum

Private Type TSQL
    TableName As String
    Limit As Long
End Type
Private this As TSQL


'''' @ClassMethodStrict
'''' This method should only be used on the default instance
''''
'@DefaultMember
'@Description "Default factory"
Public Function Create(ByVal TableName As String) As SQLlib
Attribute Create.VB_Description = "Default factory"
Attribute Create.VB_UserMemId = 0
    Dim Instance As SQLlib
    Set Instance = New SQLlib
    Instance.Init TableName
    Set Create = Instance
End Function


'@Description "Default constructor"
Friend Sub Init(ByVal TableName As String)
Attribute Init.VB_Description = "Default constructor"
    Guard.EmptyString TableName
    this.TableName = TableName
End Sub


Public Property Get TableName() As String
    TableName = this.TableName
End Property


'@Ignore WriteOnlyProperty
Public Property Let Limit(ByVal LimitValue As Long)
    this.Limit = LimitValue
End Property


'''' @ClassMethod
'''' This method can also be used on the default instance
''''
'''' Examples:
''''   >>> ?SQLlib.AsText("id")
''''   "CAST([id] AS TEXT) AS [id]"
''''
'@Description "Casts field as TEXT"
Public Function AsText(ByVal FieldName As String) As String
Attribute AsText.VB_Description = "Casts field as TEXT"
    Guard.EmptyString FieldName
    AsText = "CAST([" & FieldName & "] AS TEXT) AS [" & FieldName & "]"
End Function


'''' @ClassMethod
'''' This method can also be used on the default instance
''''
'''' Examples:
''''   >>> ?SQLlib.AliasedField("id")
''''   "[id] AS [id]"
''''
'@Ignore ProcedureNotUsed
'@Description "Adds field alias"
Public Function AliasedField(ByVal FieldName As String) As String
Attribute AliasedField.VB_Description = "Adds field alias"
    Guard.EmptyString FieldName
    AliasedField = "[" & FieldName & "] AS [" & FieldName & "]"
End Function


'''' Generates a basic SELECT FROM clause
''''
'''' Generates a basic SELECT FROM clause using table name provided to
'''' the factory. If field names are provided, an excplicit SELECT is
'''' generated; otherwise, a wildcard SELECT is generated. If this.LIMIT
'''' attribute is set (>0), the LIMIT clause is also added.
''''
'''' Args:
''''   FieldNames (array, optional, "*"):
''''     1D array of field names. A single field name can be supplied as scalar.
''''     If not provided, a wildcard is used.
''''   EscapeFields (boolean, optional, True):
''''     If True (default), field names are escaped with brackets
''''
'''' Returns:
''''   String, containing SELECT FROM clause
''''
'''' Examples:
''''   >>> ?SQLlib("people").SelectAll
''''   "SELECT * FROM [people]"
''''
''''   >>> ?SQLlib("people").SelectAll(Array("id", "firstName", "age"))
''''   "SELECT [id], [firstName], [age] FROM [people]"
''''
'@Description "Generates a basic SELECT FROM clause"
Public Function SelectAll(Optional ByVal FieldNames As Variant = Empty, _
                          Optional ByVal EscapeFields As Boolean = True) As String
Attribute SelectAll.VB_Description = "Generates a basic SELECT FROM clause"
    Dim EscapeOpen As String
    EscapeOpen = IIf(EscapeFields, "[", vbNullString)
    Dim EscapeClose As String
    EscapeClose = IIf(EscapeFields, "]", vbNullString)
    
    Dim Fields As String
    If (VarType(FieldNames) And vbArray) = vbArray Then
        Fields = EscapeOpen & Join(FieldNames, EscapeClose & ", " & EscapeOpen) & EscapeClose
    ElseIf Len(FieldNames) > 0 Then
        Fields = EscapeOpen & FieldNames & EscapeClose
    Else
        Fields = "*"
    End If
    Dim LimitClause As String
    If this.Limit > 0 Then LimitClause = " LIMIT " & CStr(this.Limit)
    SelectAll = "SELECT " & Fields & " FROM [" & this.TableName & "]" & LimitClause
End Function


'''' Generates a SELECT FROM clause, having id cast as TEXT
''''
'''' Generates a SELECT FROM clause, having id cast as TEXT. All field names
'''' are escaped with brackets.
''''
'''' Args:
''''   FieldNames (array):
''''     1D array of field names.
''''
'''' Returns:
''''   String, containing SELECT FROM clause, having id cast as TEXT
''''
'''' Examples:
''''   >>> ?SQLlib("people").SelectIdAsText(Array("id", "email", "age"))
''''   "SELECT CAST([id] AS TEXT) AS [id], [email], [age] FROM [people]"
''''
'@Description "Generates a SELECT FROM clause, having id cast as TEXT"
Public Function SelectIdAsText(ByVal FieldNames As Variant) As String
Attribute SelectIdAsText.VB_Description = "Generates a SELECT FROM clause, having id cast as TEXT"
    Dim TypeCheck As Boolean
    TypeCheck = ((VarType(FieldNames) And vbArray) = vbArray)
    Guard.Expression TypeCheck, "SQLlib", "FieldNames is not array"
    
    Dim Fields() As String
    ReDim Fields(LBound(FieldNames) To UBound(FieldNames))
    Dim FieldIndex As Long
    FieldIndex = LBound(FieldNames)
    Fields(FieldIndex) = AsText(FieldNames(FieldIndex))
    For FieldIndex = LBound(FieldNames) + 1 To UBound(FieldNames)
        Fields(FieldIndex) = "[" & FieldNames(FieldIndex) & "]"
    Next FieldIndex
    SelectIdAsText = SelectAll(Fields, False)
End Function


'''' Generates a SELECT FROM clause, having all non-TEXT fields cast as TEXT
''''
'''' Generates a SELECT FROM clause, having all non-TEXT fields cast as TEXT.
'''' All field names are escaped with brackets.
''''
'''' Args:
''''   FieldNames (array):
''''     1D array of field names.
''''
''''   FieldTypes (array of ADODB.DataTypeEnum):
''''     Field types (ADODB.DataTypeEnum) information is collected during initial
''''     table introspection. Fields reported as non-TEXT (at present, only adVarWChar
''''     is considered as TEXT) are cast as TEXT.
''''
'''' Returns:
''''   String, containing SELECT FROM clause, having all non-TEXT fields cast as TEXT
''''
'''' Examples:
''''   >>> ?SQLlib("people").SelectAllAsText(Array("id", "email", "age"), Array(adInteger, adVarWChar, adInteger))
''''   "SELECT CAST([id] AS TEXT) AS [id], [email], CAST([age] AS TEXT) AS [age] FROM [people]"
''''
'@Description "Generates a SELECT FROM clause, having all non-TEXT fields cast as TEXT"
Public Function SelectAllAsText(ByVal FieldNames As Variant, ByVal FieldTypes As Variant) As String
Attribute SelectAllAsText.VB_Description = "Generates a SELECT FROM clause, having all non-TEXT fields cast as TEXT"
    '''' Make sure that both FieldNames and FieldTypes are arrays and have the same base and size
    Guard.IncompatibleArrays FieldNames, FieldTypes
    Guard.EmptyString FieldNames(LBound(FieldNames))
    
    Dim Fields() As Variant
    ReDim Fields(LBound(FieldNames) To UBound(FieldNames))
    Dim FieldIndex As Long
    For FieldIndex = LBound(FieldNames) To UBound(FieldNames)
        Select Case FieldTypes(FieldIndex)
            Case ADODB.DataTypeEnum.adVarWChar
                Fields(FieldIndex) = "[" & FieldNames(FieldIndex) & "]"
            Case Else
                Fields(FieldIndex) = AsText(FieldNames(FieldIndex))
        End Select
    Next FieldIndex
    SelectAllAsText = SelectAll(Fields, False)
End Function


'''' Examples:
''''   >>> ?SQLlib.QuoteSQLString("Abc")
''''   "'Abc'"
''''
''''   >>> ?SQLlib.QuoteSQLString("A'b'c")
''''   "'A''b''c'"
''''
'@Description "Quotes SQL string, doubling single quotes."
Public Function QuoteSQLString(ByVal SQLString As String) As String
Attribute QuoteSQLString.VB_Description = "Quotes SQL string, doubling single quotes."
    QuoteSQLString = "'" & Replace(SQLString, "'", "''") & "'"
End Function


'@Description "Wraps a non-CTE SELECT query to yield SQL returning RecordCount for the original query."
Public Function CountSelectNoCTE(ByVal SQLQuery As String) As String
Attribute CountSelectNoCTE.VB_Description = "Wraps a non-CTE SELECT query to yield SQL returning RecordCount for the original query."
    CountSelectNoCTE = _
        "WITH" & vbNewLine & "  subset AS (" & vbNewLine & _
        "    " & Replace(Replace(SQLQuery, ";", vbNullString), vbNewLine, vbNewLine & "    ") & _
        vbNewLine & ")" & vbNewLine & _
        "SELECT count(*) AS counter FROM subset;"
End Function


'@Description "Wraps a SELECT query to yield SQL returning RecordCount for the original query."
Public Function CountSelect(ByVal SQLQuery As String) As String
Attribute CountSelect.VB_Description = "Wraps a SELECT query to yield SQL returning RecordCount for the original query."
    CountSelect = "SELECT count(*) AS counter FROM (" & vbNewLine & _
                  "    " & Replace(Replace(SQLQuery, ";", vbNullString), vbNewLine, vbNewLine & "    ") & _
                  vbNewLine & ");"
End Function


'@Description "Formats a SELECT query as a subquery suitable for the FROM clause."
Public Function SubQuery(ByVal SQLQuery As String) As String
Attribute SubQuery.VB_Description = "Formats a SELECT query as a subquery suitable for the FROM clause."
    SubQuery = " (" & vbNewLine & _
               "    " & Replace(Replace(SQLQuery, ";", vbNullString), vbNewLine, vbNewLine & "    ") & _
               vbNewLine & "    )" & vbNewLine
End Function


'''' TODO: Switch to LiteMetaSQLEngine
'@Description "Generates query returning SQLite version"
Public Property Get SQLiteVersion() As String
Attribute SQLiteVersion.VB_Description = "Generates query returning SQLite version"
    SQLiteVersion = "SELECT sqlite_version() AS version"
End Property


'''' TODO: Switch to LiteMetaSQLEngine
'@Description "Generates query returning available SQLite collations"
Public Property Get SQLiteCollations() As String
Attribute SQLiteCollations.VB_Description = "Generates query returning available SQLite collations"
    SQLiteCollations = "SELECT * FROM pragma_collation_list AS collations ORDER BY name"
End Property


'''' Generates parametrized UPDATE statement for a single record
''''
'''' Generates a fully parametrized UPDATE query for a single record using
'''' table name supplied to the factory.
''''
'''' Args:
''''   FieldNames (array):
''''     1D array of field names. The first field name is assumed to be the
''''     primary key for the WHERE clause.
''''
'''' Returns:
''''   String, containing parametrized UPDATE statement
''''
'''' Examples:
''''   >>> ?SQLlib("people").UpdateSingleRecord(Array("id", "email", "age"))
''''   "UPDATE [people] SET ([email], [age]) = (?, ?) WHERE [id] = ?"
''''
'@Description "Generates parametrized UPDATE statement for a single record"
Public Function UpdateSingleRecord(ByVal FieldNames As Variant) As String
Attribute UpdateSingleRecord.VB_Description = "Generates parametrized UPDATE statement for a single record"
    Dim FieldCount As Long
    Dim PKName As String
    PKName = FieldNames(LBound(FieldNames)) ' Assumes that the primary key is the first field
    Dim SetFieldList As String
    SetFieldList = "[" & Join(FieldNames, "], [") & "]" ' Field list for SET clause
    SetFieldList = Right$(SetFieldList, Len(SetFieldList) - Len(PKName) - 4) ' Remove primary key field, which goes into the WHERE clause
    FieldCount = UBound(FieldNames, 1) - LBound(FieldNames, 1) + 1
    
    ' Since there is no standarad routine for repeating a multi-character patterns,
    ' prepare a sequence of placeholder characters with String and use Replace to
    ' generate the final pattern.
    Dim ValuePlaceholderList As String
    ValuePlaceholderList = Replace(String(FieldCount - 2, "@"), "@", "?, ") & "?"
    
    UpdateSingleRecord = "UPDATE [" & this.TableName & "] " & _
                         "SET (" & SetFieldList & ") = (" & ValuePlaceholderList & ") " & _
                         "WHERE [" & PKName & "] = ?"
End Function


'@Description "SQL for retrieving a single column containing attached databases paths or the path of the alias."
Public Function AttachedDbPaths( _
            Optional ByVal SchemaAlias As String = vbNullString) As String
    Dim WhereClause As String
    If Len(SchemaAlias) > 0 Then
        WhereClause = " WHERE name = " & SQLlib.QuoteSQLString(SchemaAlias)
    Else
        WhereClause = vbNullString
    End If
    AttachedDbPaths = "SELECT file FROM pragma_database_list" & WhereClause
End Function


'''' @ClassMethod
'''' This method can also be used on the default instance
''''
'''' Args:
''''   DatabaseName (string):
''''     name of the SQLite database to be attached. This argument must be
''''     cleared by the caller. It can be ":memory:", a valid file pathname or
''''     an empty string.
''''   DatabaseAlias (string, optional, vbNullString):
''''     database alias; defaults to the database file name without extension.
''''
'''' Returns:
''''   String, containing ATTACH statement
''''
'''' Examples:
''''   >>> ?SQLlib.Attach(":memory:", "memory")
''''   "ATTACH ':memory:' AS [memory]"
''''
'@Description "Generates SQLite ATTACH statement"
Public Function Attach(ByVal Database As String, _
              Optional ByVal DatabaseAlias As String = vbNullString) As String
Attribute Attach.VB_Description = "Generates SQLite ATTACH statement"
    '@Ignore SelfAssignedDeclaration
    Dim fso As New Scripting.FileSystemObject
    Dim Alias As String
    If Len(DatabaseAlias) > 0 Then
        Alias = DatabaseAlias
    ElseIf Database = ":memory:" Then
        Alias = "memory"
    Else
        Alias = fso.GetBaseName(Database)
    End If
    Attach = "ATTACH '" & Replace(Database, "'", "''") & "' AS [" & Alias & "]"
End Function


'''' @ClassMethod
'''' This method can also be used on the default instance
''''
'''' Examples:
''''   >>> ?SQLlib.Detach("SQLiteDB")
''''   "DETACH [SQLiteDB]"
''''
'@Description "Generates SQLite DETACH statement"
Public Function Detach(ByVal DatabaseName As String) As String
Attribute Detach.VB_Description = "Generates SQLite DETACH statement"
    Detach = "DETACH [" & DatabaseName & "]"
End Function


'@Description "Generates a query returning/setting the foreign keys status."
Public Function FKStatus(Optional ByVal Status As Variant) As String
Attribute FKStatus.VB_Description = "Generates a query returning/setting the foreign keys status."
    If IsMissing(Status) Then
        FKStatus = "SELECT * FROM pragma_foreign_keys"
    Else
        FKStatus = "PRAGMA foreign_keys=" & Status
    End If
End Function


'''' Note:
''''   This pragma is automatically switched off at conclusion of transaction
'@Description "Generates a query deferring foreign keys before TX."
Public Property Get DeferForeignKeys() As String
Attribute DeferForeignKeys.VB_Description = "Generates a query deferring foreign keys before TX."
    DeferForeignKeys = "PRAGMA defer_foreign_keys=True"
End Property


'''' @ClassMethod
'''' This method can also be used on the default instance
''''
'''' Examples:
''''   >>> ?SQLlib.Vacuum
''''   "VACUUM"
''''
''''   >>> ?SQLlib.Vacuum("main")
''''   "VACUUM [main]"
''''
''''   >>> ?SQLlib.Vacuum(vbNullString, vbNullString)
''''   "VACUUM"
''''
''''   >>> ?SQLlib.Vacuum(vbNullString, "C:\TEMP\qqq.db")
''''   "VACUUM INTO 'C:\TEMP\qqq.db'"
''''
''''   >>> ?SQLlib.Vacuum("main", "C:\TEMP\qq'q.db")
''''   "VACUUM [main] INTO 'C:\TEMP\qq''q.db'"
''''
'@Description "Generates SQLite VACUUM statement"
Public Function Vacuum(Optional ByVal SrcSchemaAlias As String = vbNullString, _
                       Optional ByVal DstDbPathName As String = vbNullString _
                      ) As String
Attribute Vacuum.VB_Description = "Generates SQLite VACUUM statement"
    Dim Alias As String
    If Len(SrcSchemaAlias) > 0 Then
        Alias = " [" & SrcSchemaAlias & "]"
    End If
    Dim IntoClause As String
    If Len(DstDbPathName) > 0 Then
        IntoClause = " INTO '" & Replace(DstDbPathName, "'", "''") & "'"
    End If
    Vacuum = "VACUUM" & Alias & IntoClause
End Function


'''' @ClassMethod
'''' This method can also be used on the default instance
''''
'''' Generates query cloning data between two tables. Tables may belong to
'''' different databases attached to the same connection.
''''
'''' Args:
''''   SourceDatabase (string):
''''     "Copy from" database alias
''''
''''   SourceTable (string):
''''     "Copy from" table name
''''
''''   TargetDatabase (string, optional, "main"):
''''     "Copy to" database alias
''''
''''   TargetTable (string, optional, <SourceTable>):
''''     "Copy to" table name
''''
'''' Examples:
''''   >>> ?SQLlib.CopyTableData("People", "Contacts")
''''   "INSERT INTO [main].[Contacts] SELECT * FROM [People].[Contacts]"
''''
'@Description "Generates a query cloning table data."
Public Function CopyTableData(ByVal SourceDatabase As String, _
                              ByVal SourceTable As String, _
                     Optional ByVal TargetDatabase As String = "main", _
                     Optional ByVal TargetTable As String = vbNullString) As String
Attribute CopyTableData.VB_Description = "Generates a query cloning table data."
    Dim InsertClause As String
    Dim SelectClause As String
    InsertClause = IIf(Len(TargetTable) > 0, _
        "INSERT INTO [" & TargetDatabase & "].[" & TargetTable & "] ", _
        "INSERT INTO [" & TargetDatabase & "].[" & SourceTable & "] " _
    )
    SelectClause = _
        "SELECT * FROM [" & SourceDatabase & "].[" & SourceTable & "]"
    CopyTableData = InsertClause & SelectClause
End Function
