VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableOFromDbHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "SQLite table helper class. Retrieves table metadata from an SQLite database."
'@Folder "SQLiteDBdev.DB Objects"
'@ModuleDescription "SQLite table helper class. Retrieves table metadata from an SQLite database."
'@PredeclaredId
'@Exposed
'@IgnoreModule IndexedDefaultMemberAccess
Option Explicit


Private Type TTableOFromDbHelper
    TableName As String
    DbManager As LiteDB
    FieldCount As Long
End Type
Private this As TTableOFromDbHelper


'@DefaultMember
'@Description "Default factory"
Public Function Create(ByVal TableName As String, _
                       ByVal DbManager As LiteDB) As TableOFromDbHelper
Attribute Create.VB_Description = "Default factory"
Attribute Create.VB_UserMemId = 0
    Dim Instance As TableOFromDbHelper
    Set Instance = New TableOFromDbHelper
    Instance.Init TableName, DbManager
    Set Create = Instance
End Function


'@Description "Default constructor"
Friend Sub Init(ByVal TableName As String, _
                ByVal DbManager As LiteDB)
Attribute Init.VB_Description = "Default constructor"
    Guard.EmptyString TableName
    Guard.NullReference DbManager
    
    this.TableName = TableName
    Set this.DbManager = DbManager
End Sub


'''' This is a helper routine for the "FromDb" factory. It uses
'''' GetTableColumnsEx to retrieve field metadata via a combination
'''' of a Pragma function and direct table schema parsing.
'''' It should not access private class attributes, and the required
'''' dependencies should be explicitly injected.
'''' It should not be ordinarily accessed from outside, but it can be
'''' tested independently, hence Friend qualification.
''''
'''' Args:
''''   TableName (string):
''''     Table name
''''
''''   DbManager (SQLiteDB)
''''     SQLite database manager
''''
'''' Returns:
''''   Array of FieldO
''''
'@Description "Gets fields metadata from the database."
Friend Function GetDbFields() As Variant
Attribute GetDbFields.VB_Description = "Gets fields metadata from the database."
    Dim TableName As String
    TableName = this.TableName
    Dim DbManager As LiteDB
    Set DbManager = this.DbManager
    
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = DbManager.GetTableColumnsEx(TableName)
    Dim DbFieldsProps As Variant
    DbFieldsProps = ArrayLib.TransposeArray(AdoRecordset.GetRows)
    Dim FieldPropMap As Scripting.Dictionary
    Set FieldPropMap = New Scripting.Dictionary
    FieldPropMap.CompareMode = TextCompare
    Dim FieldPropIndex As Long
    For FieldPropIndex = 0 To AdoRecordset.Fields.Count - 1
        FieldPropMap(AdoRecordset.Fields(FieldPropIndex).Name) = FieldPropIndex
    Next FieldPropIndex
    
    Dim DbFields() As FieldO
    ReDim DbFields(1 To UBound(DbFieldsProps, 1) + 1)
    Dim DbField As FieldO
    Dim DefaultValue As Variant
    Dim UniqueValue As Variant
    Dim FieldIndex As Long
    For FieldIndex = 0 To UBound(DbFields) - 1
        DefaultValue = DbFieldsProps(FieldIndex, FieldPropMap("dflt_value"))
        DefaultValue = IIf(IsNull(DefaultValue), vbNullString, "(" & DefaultValue & ")")
        UniqueValue = DbFieldsProps(FieldIndex, FieldPropMap("unique"))
        UniqueValue = IIf(IsNull(UniqueValue), False, UniqueValue)
        UniqueValue = CBool(UniqueValue)
        Set DbField = FieldO( _
            FieldName:=DbFieldsProps(FieldIndex, FieldPropMap("name")), _
            DataType:=DbFieldsProps(FieldIndex, FieldPropMap("type")), _
            NotNull:=CBool(DbFieldsProps(FieldIndex, FieldPropMap("notnull"))), _
            Default:=DefaultValue, _
            Check:=DbFieldsProps(FieldIndex, FieldPropMap("check")), _
            Unique:=UniqueValue, _
            Collation:=DbFieldsProps(FieldIndex, FieldPropMap("collate")) _
        )
        Set DbFields(FieldIndex + 1) = DbField
    Next FieldIndex
    this.FieldCount = UBound(DbFields)
    GetDbFields = DbFields
End Function


'''' This is a helper routine for the "FromDb" factory.
'''' It should not access private class attributes, and the required
'''' dependencies should be explicitly injected.
'''' It should not be ordinarily accessed from outside, but it can be
'''' tested independently, hence Friend qualification.
''''
'@Description "Parses table schema from the database and returns constraints."
Friend Function GetConstraints() As Scripting.Dictionary
Attribute GetConstraints.VB_Description = "Parses table schema from the database and returns constraints."
    Dim TableName As String
    TableName = this.TableName
    Dim DbManager As LiteDB
    Set DbManager = this.DbManager
    Dim FieldCount As Long
    FieldCount = this.FieldCount
                               
    '''' Constraints nees to be parsed from the schema
    Dim TableSQL As String
    '@Ignore ImplicitDefaultMemberAccess, IndexedRecursiveDefaultMemberAccess
    TableSQL = DbManager.GetAdoRecordset(LiteMetaSQL.TableSQL(TableName))("sql")
    ' Debug.Print TableSQL
    Dim CreateTableLines As Variant
    '''' The last schema line <vbLf & ")"> is chopped off
    CreateTableLines = Left$(TableSQL, Len(TableSQL) - 2)
    '''' Split yields a 0-based array.
    '''' Skip the first "CREATE TABLE" line and the next <FieldCount> lines
    '''' The first constraint line has index <FieldCount> + 1
    CreateTableLines = Split(CreateTableLines, vbLf)

    Dim DbConstraints As Scripting.Dictionary
    Set DbConstraints = InitDbConstraints
    
    Dim ConSQL As String
    Dim ConName As String
    Dim ConType As String
    Dim ConIndex As Long
    For ConIndex = FieldCount + 1 To UBound(CreateTableLines)
        ConSQL = CreateTableLines(ConIndex)
        ConSQL = GetConstraintNameType(ConSQL, ConName, ConType)
        Select Case ConType
            Case "C"
                Dim ConCK As ConstraintCK
                Set ConCK = ParseCheck(ConSQL, ConName)
                If ConName = vbNullString Then
                    ConName = "____" & ConCK.SQL
                End If
                Set DbConstraints("CKs")(ConName) = ConCK
            Case "F"
                ConSQL = Replace(ConSQL, " ", vbNullString)
                Dim ConFK As ConstraintFK
                Set ConFK = ParseForeignKey(ConSQL, ConName)
                If ConName = vbNullString Then
                    ConName = "____" & ConFK.SQL
                End If
                Set DbConstraints("FKs")(ConName) = ConFK
            Case "U"
                ConSQL = Replace(ConSQL, " ", vbNullString)
                Dim ConUQ As ConstraintUQ
                Set ConUQ = ParseUnique(ConSQL, ConName)
                If ConName = vbNullString Then
                    ConName = "____" & ConUQ.SQL
                End If
                Set DbConstraints("UQs")(ConName) = ConUQ
            Case "P"
                ConSQL = Replace(ConSQL, " ", vbNullString)
                Dim ConPK As ConstraintPK
                Set ConPK = ParsePrimaryKey(ConSQL, ConName)
                Set DbConstraints("PK") = ConPK
        End Select
    Next ConIndex

    Set GetConstraints = DbConstraints
End Function


'''' This is a helper routine for the "FromDb" factory.
''''
'''' Returns:
''''   Scripting.Dictionary with container members to be used
''''   as follows (same as the default TableO constuctor):
''''     FKs (dictionary):
''''       key - FOREIGN KEY constraint name (for anon constraint, use constraint SQL prefixed with '____')
''''       value - ConstraintFK
''''     UQs (dictionary):
''''       key - UNIQUE constraint name (for anon constraint, use constraint SQL prefixed with '____')
''''       value - ConstraintUQ:
''''     CKs (dictionary):
''''       key - CHECK constraint name (for anon constraint, use constraint SQL prefixed with '____')
''''       value - ConstraintCK:
''''     PKs (ConstraintPK):
''''
'@Description "Initializes a structure for constraint objects."
Private Function InitDbConstraints() As Scripting.Dictionary
Attribute InitDbConstraints.VB_Description = "Initializes a structure for constraint objects."
    Dim DbConstraints As Scripting.Dictionary
    Set DbConstraints = New Scripting.Dictionary
    DbConstraints.CompareMode = TextCompare
    
    Set DbConstraints("UQs") = New Scripting.Dictionary
    DbConstraints("UQs").CompareMode = TextCompare

    Set DbConstraints("FKs") = New Scripting.Dictionary
    DbConstraints("FKs").CompareMode = TextCompare
    
    Set DbConstraints("CKs") = New Scripting.Dictionary
    DbConstraints("CKs").CompareMode = TextCompare
    
    Set DbConstraints("PK") = Nothing

    Set InitDbConstraints = DbConstraints
End Function


'''' Parses constraint name and type from TABLE CREATE line.
''''
'''' Args:
''''   ConSQL (string):
''''     TABLE CREATE line from table schema, containg constraint SQL definition
''''
''''   ConName (string, out):
''''     Returned constraint name
''''
''''   ConType (string, out):
''''     Returned constraint type (the first letter of the SQL constraint
''''     definition term.
''''
'''' Returns:
''''   String, updated constraint SQL line with name clause removed.
''''
'@Description "Parses constraint name and type from TABLE CREATE line."
Friend Function GetConstraintNameType(ByRef ConSQL As String, _
                                      ByRef ConName As String, _
                                      ByRef ConType As String) As String
Attribute GetConstraintNameType.VB_Description = "Parses constraint name and type from TABLE CREATE line."
    Static reTrim As RegExp
    Set reTrim = New RegExp
    reTrim.Global = False
    reTrim.IgnoreCase = True
    reTrim.Pattern = "^\s*"
    
    ConSQL = reTrim.Replace(ConSQL, vbNullString)
    
    Static reConName As RegExp
    Set reConName = New RegExp
    reConName.Global = False
    reConName.IgnoreCase = True
    reConName.Pattern = "^CONSTRAINT\s*""?(\w+)""?\s*"
    
    ConName = vbNullString
    Dim ReMatches As MatchCollection
    Set ReMatches = reConName.Execute(ConSQL)
    If ReMatches.Count > 0 Then
        Dim ReMatch As Match
        Set ReMatch = ReMatches(0)
        ConSQL = reConName.Replace(ConSQL, vbNullString)
        '@Ignore IndexedUnboundDefaultMemberAccess
        ConName = ReMatch.SubMatches(0)
    End If
    ConType = Left$(ConSQL, 1)
    
    GetConstraintNameType = ConSQL
End Function


Friend Function ParseCheck(ByVal ConSQL As String, ByVal ConName As String) As ConstraintCK
    Dim EndPos As Long
    EndPos = InStrRev(ConSQL, ")")
    Dim ConExpression As String
    ConExpression = Mid$(ConSQL, Len("CHECK(") + 1, EndPos - Len("CHECK(") - 1)
    Set ParseCheck = ConstraintCK(ConExpression, ConName)
End Function


Friend Function ParseForeignKey(ByVal ConSQL As String, ByVal ConName As String) As ConstraintFK
    Dim BegPos As Long
    Dim EndPos As Long
    
    BegPos = InStr(ConSQL, "(""") + Len("(""")
    EndPos = InStrRev(ConSQL, """)REFERENCES""")
    Dim FieldList As String
    FieldList = Mid$(ConSQL, BegPos, EndPos - BegPos)
    Dim FieldNames As Variant
    FieldNames = Split(FieldList, """,""")
    
    BegPos = EndPos + Len(""")REFERENCES""")
    EndPos = InStr(BegPos, ConSQL, """(""")
    Dim ForeignTableName As String
    ForeignTableName = Mid$(ConSQL, BegPos, EndPos - BegPos)
    
    BegPos = EndPos + Len("""(""")
    EndPos = InStr(BegPos, ConSQL, """)")
    FieldList = Mid$(ConSQL, BegPos, EndPos - BegPos)
    Dim ForeignFieldNames As Variant
    ForeignFieldNames = Split(FieldList, """,""")
    
    Dim FKClauseTerms As Scripting.Dictionary
    Set FKClauseTerms = ParseFKClause(ConSQL)
    Dim OnDelete As String
    OnDelete = FKClauseTerms("ONDELETE")
    Dim OnUpdate As String
    OnUpdate = FKClauseTerms("ONUPDATE")
    
    Set ParseForeignKey = ConstraintFK(FieldNames, _
                                       ForeignTableName, _
                                       ForeignFieldNames, _
                                       OnDelete, _
                                       OnUpdate, _
                                       ConName)
End Function


'''' Parsing helper for the "FromDb" factory. It does not access any private attributes.
'''' It should not be ordinarily accessed from outside, but it can be tested
'''' independently, hence Friend qualification.
''''
'''' Examples:
''''   >>> ?ParseFKClause(" on  delete no  action on  update set  null")("ONDELETE")
''''   NO ACTION
''''
''''   >>> ?ParseFKClause(" on  delete no  action on  update set  null")("ONUPDATE")
''''   SET NULL
''''
'@Description "Parses foreign key clause"
Friend Function ParseFKClause(ByVal FKClause As String) As Scripting.Dictionary
Attribute ParseFKClause.VB_Description = "Parses foreign key clause"
    Dim FKActionPrefixMap As Scripting.Dictionary
    Set FKActionPrefixMap = New Scripting.Dictionary
    FKActionPrefixMap.CompareMode = TextCompare
    FKActionPrefixMap("SETN") = "SET NULL"
    FKActionPrefixMap("SETD") = "SET DEFAULT"
    FKActionPrefixMap("CASC") = "CASCADE"
    FKActionPrefixMap("REST") = "RESTRICT"
    FKActionPrefixMap("NOAC") = "NO ACTION"
    
    Dim FKClauseTerms As Scripting.Dictionary
    Set FKClauseTerms = New Scripting.Dictionary
    FKClauseTerms.CompareMode = TextCompare
    
    Dim FKClauseS As String
    FKClauseS = UCase$(Replace(FKClause, " ", vbNullString))
    
    Dim FKAction As String
    Dim FKActionKey As String
    Dim TermPosition As Long
    
    FKActionKey = "ONDELETE"
    TermPosition = InStr(FKClauseS, FKActionKey)
    FKAction = Mid$(FKClauseS, TermPosition + Len(FKActionKey), 4)
    If TermPosition > 0 And FKActionPrefixMap.Exists(FKAction) Then
        FKClauseTerms(FKActionKey) = FKActionPrefixMap(FKAction)
    End If
    
    FKActionKey = "ONUPDATE"
    TermPosition = InStr(FKClauseS, FKActionKey)
    FKAction = Mid$(FKClauseS, TermPosition + Len(FKActionKey), 4)
    If TermPosition > 0 And FKActionPrefixMap.Exists(FKAction) Then
        FKClauseTerms(FKActionKey) = FKActionPrefixMap(FKAction)
    End If
    
    Set ParseFKClause = FKClauseTerms
End Function


Friend Function ParseUnique(ByVal ConSQL As String, ByVal ConName As String) As ConstraintUQ
    Dim BegPos As Long
    BegPos = InStr(ConSQL, "(""") + Len("(""")
    Dim EndPos As Long
    EndPos = InStrRev(ConSQL, """)")
    Dim FieldList As String
    FieldList = Mid$(ConSQL, BegPos, EndPos - BegPos)
    Dim FieldNames As Variant
    FieldNames = Split(FieldList, """,""")
    
    Set ParseUnique = ConstraintUQ(FieldNames, ConName)
End Function


Friend Function ParsePrimaryKey(ByVal ConSQL As String, ByVal ConName As String) As ConstraintPK
    Dim BegPos As Long
    BegPos = InStr(ConSQL, "(""") + Len("(""")
    Dim EndPos As Long
    EndPos = InStrRev(ConSQL, """AUTOINCREMENT)", -1, vbTextCompare)
    Dim AutoIncrement As Boolean
    If EndPos > 0 Then
        AutoIncrement = True
    Else
        AutoIncrement = False
        EndPos = InStrRev(ConSQL, """)")
    End If
    Dim FieldList As String
    FieldList = Mid$(ConSQL, BegPos, EndPos - BegPos)
    Dim FieldNames As Variant
    FieldNames = Split(FieldList, """,""")
    
    Set ParsePrimaryKey = ConstraintPK(FieldNames, ConName, AutoIncrement)
End Function
