VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LiteSQLlib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "SQL generator for LiteADOlib."
'@ModuleDescription "SQL generator for LiteADOlib."
'@PredeclaredId
'@Folder "SQLite.Abstract"
'@IgnoreModule IndexedDefaultMemberAccess

'''' Common Args:
''''   TableName (string):
''''     Name of the destination table, possibly with schema alias.
''''
''''   TableData (array):
''''     2D array with data and field names in the first row.
''''
''''   ColNames (array):
''''     1D array of field names to be included. The order may be
''''     important in terms of performance. When there is a corresponding
''''     usable multi-field index, it is better to supply field names
''''     according to their appearance in the index.
''''
''''   DataRowIndices (array):
''''     1D array of record indices. The 1st data row (2nd row in the
''''     TableData array) has index of LBound(TableData, 2) (one is
''''     added by the code.

Option Explicit

Private Type TObjectState
    TableName As String
End Type
Private this As TObjectState


'''' @ClassMethodStrict
'''' This method should only be used on the default instance
''''
'@DefaultMember
'@Description "Default factory"
Public Function Create(ByVal TableName As String) As LiteSQLlib
Attribute Create.VB_Description = "Default factory"
Attribute Create.VB_UserMemId = 0
    Dim Instance As LiteSQLlib
    Set Instance = New LiteSQLlib
    Instance.Init TableName
    Set Create = Instance
End Function


'@Description "Default constructor"
Friend Sub Init(ByVal TableName As String)
Attribute Init.VB_Description = "Default constructor"
    this.TableName = TableName
End Sub


'@Ignore ProcedureNotUsed
Public Property Get TableName() As String
    TableName = this.TableName
End Property


Public Property Let TableName(ByVal Values As String)
    this.TableName = Values
End Property


'''' Extracts field names from 2D array with header.
''''
'''' Extracts field names from the header row of a 2D array table. If any field
'''' name contains double qoute, an error is raised. If any field name contains
'''' a space, the name is double quoted.
''''
'''' Returns:
''''   1D array of field names.
''''
'''' Examples:
''''   >>> ?Join(LiteSQLlib.FieldNamesFrom2DArrayHeader(FixUtils.People2D), "#")
''''   id#first_name#last_name#age#gender#email#country#domain
''''
''''   >>> ?Join(LiteSQLlib.FieldNamesFrom2DArrayHeader(FixUtils.People2DSpacesInFieldNames), "#")
''''   id#"first name"#"last name"#age#gender#email#country#domain
''''
'@Description "Extracts field names from 2D array with header."
Public Function FieldNamesFrom2DArrayHeader(ByRef TableData As Variant) As String()
Attribute FieldNamesFrom2DArrayHeader.VB_Description = "Extracts field names from 2D array with header."
    Dim FieldNames() As String
    ReDim FieldNames(LBound(TableData, 2) To UBound(TableData, 2))
    Dim HeaderRowIndex As Long
    HeaderRowIndex = LBound(TableData, 1)
    Dim FieldIndex As Long
    For FieldIndex = LBound(FieldNames) To UBound(FieldNames)
        FieldNames(FieldIndex) = SQLQuoter.QuoteSQLName( _
                                 TableData(HeaderRowIndex, FieldIndex))
    Next FieldIndex
    FieldNamesFrom2DArrayHeader = FieldNames
End Function


'@Description "Generates the INSERT OR IGNORE INTO VALUES query for bulk insert."
Public Function DefaultUPSERT(ByRef TableData As Variant) As String
Attribute DefaultUPSERT.VB_Description = "Generates the INSERT OR IGNORE INTO VALUES query for bulk insert."
    Dim FieldNames() As String
    ReDim FieldNames(LBound(TableData, 2) To UBound(TableData, 2))
    Dim FieldTerms() As String
    ReDim FieldTerms(LBound(FieldNames) To UBound(FieldNames))
    Dim HeaderRowIndex As Long
    HeaderRowIndex = LBound(TableData, 1)
    Dim FieldName As String
    Dim FieldIndex As Long
    For FieldIndex = LBound(FieldNames) To UBound(FieldNames)
        FieldName = SQLQuoter.QuoteSQLName(TableData(HeaderRowIndex, FieldIndex))
        FieldTerms(FieldIndex) = (FieldName & "=" & "excluded." & FieldName)
    Next FieldIndex
    DefaultUPSERT = "ON CONFLICT DO UPDATE SET" & vbNewLine & _
                    "    " & Join(FieldTerms, "," & vbNewLine & "    ")
End Function


'@Description "Generates an array of record values for INSERT INTO VALUES"
Public Function RecordValuesFrom2DArray(ByRef TableData As Variant, _
            Optional ByVal DataRowIndices As Variant = Empty) As String()
Attribute RecordValuesFrom2DArray.VB_Description = "Generates an array of record values for INSERT INTO VALUES"
    If Not (IsArray(DataRowIndices) Or IsEmpty(DataRowIndices)) Then
        Err.Raise ErrNo.InvalidParameterErr, "SQLlib/RecordValuesFrom2DArray", _
                  "The second argument should be array or Empty/missing."
    End If
    Dim HeaderOffest As Long
    HeaderOffest = 1
    
    Dim RecordValues() As String
    Dim RecordLB As Long
    Dim RecordUB As Long
    Dim RecordSrcIndex As Long
    Dim RecordDstIndex As Long
    If IsEmpty(DataRowIndices) Then
        RecordLB = LBound(TableData, 1)
        RecordUB = UBound(TableData, 1) - HeaderOffest
        ReDim RecordValues(RecordLB To RecordUB)
        For RecordDstIndex = RecordLB To RecordUB
            RecordSrcIndex = RecordDstIndex + HeaderOffest
            GoSub RECORD_LOOP
        Next RecordDstIndex
    Else
        RecordLB = LBound(DataRowIndices)
        RecordUB = UBound(DataRowIndices)
        ReDim RecordValues(RecordLB To RecordUB)
        For RecordDstIndex = RecordLB To RecordUB
            RecordSrcIndex = DataRowIndices(RecordDstIndex) + HeaderOffest
            GoSub RECORD_LOOP
        Next RecordDstIndex
    End If
    RecordValuesFrom2DArray = RecordValues

    Exit Function
RECORD_LOOP:
    Dim FieldUB As Long
    FieldUB = UBound(TableData, 2)
    Dim FieldLB As Long
    FieldLB = LBound(TableData, 2)
    Dim FieldValues() As String
    ReDim FieldValues(FieldLB To FieldUB)
    Dim FieldValue As Variant
    Dim FieldIndex As Long
    For FieldIndex = FieldLB To FieldUB
        FieldValue = TableData(RecordSrcIndex, FieldIndex)
        FieldValues(FieldIndex) = SQLQuoter.CStrSQL(FieldValue)
    Next FieldIndex
    RecordValues(RecordDstIndex) = "(" & Join(FieldValues, ", ") & ")"
    Return
End Function


'@Description "Generates the INSERT OR IGNORE INTO VALUES query for bulk insert."
Public Function InsertNewRecordsFrom2DArray(ByRef TableData As Variant, _
            Optional ByVal DataRowIndices As Variant = Empty) As String
Attribute InsertNewRecordsFrom2DArray.VB_Description = "Generates the INSERT OR IGNORE INTO VALUES query for bulk insert."
    Dim FieldNames() As String
    FieldNames = FieldNamesFrom2DArrayHeader(TableData)
    Dim RecordValues() As String
    RecordValues = RecordValuesFrom2DArray(TableData, DataRowIndices)
    Dim SQLQuery As String
    SQLQuery = "INSERT OR IGNORE INTO " & this.TableName & "(" & _
               Join(FieldNames, ", ") & ")" & vbNewLine & _
               "VALUES" & vbNewLine & "    " & _
               Join(RecordValues, "," & vbNewLine & "    ")
    InsertNewRecordsFrom2DArray = SQLQuery
End Function


'@Description "Generates the INSERT OR UPDATE INTO VALUES query for bulk insert."
Public Function UpsertFrom2DArray(ByRef TableData As Variant, _
            Optional ByVal DataRowIndices As Variant = Empty) As String
Attribute UpsertFrom2DArray.VB_Description = "Generates the INSERT OR UPDATE INTO VALUES query for bulk insert."
    Dim FieldNames() As String
    FieldNames = FieldNamesFrom2DArrayHeader(TableData)
    Dim RecordValues() As String
    RecordValues = RecordValuesFrom2DArray(TableData, DataRowIndices)
    Dim UPSERTClause As String
    UPSERTClause = DefaultUPSERT(TableData)
    Dim SQLQuery As String
    SQLQuery = "INSERT INTO " & this.TableName & "(" & _
               Join(FieldNames, ", ") & ")" & vbNewLine & _
               "VALUES" & vbNewLine & "    " & _
               Join(RecordValues, "," & vbNewLine & "    ") & vbNewLine & _
               UPSERTClause
    UpsertFrom2DArray = SQLQuery
End Function


'''' Generates set membership test clause.
''''
'''' Examples:
''''   >>> ?LiteSQLlib.TestSetMembership(FixUtils.People2D, Array("country", "age"), Array(2, 3))
''''   (country, age) IN (VALUES
''''       ('Kazakhstan', '62'),
''''       ('Belarus', '26'))
''''
''''   >>> ?LiteSQLlib.TestSetMembership(FixUtils.People2D, Array("country"), Array(2, 3))
''''   country IN ('Kazakhstan', 'Belarus')
''''
'@Description "Generates set (IN) condition."
Public Function TestSetMembership(ByRef TableData As Variant, _
                                  ByRef ColNames As Variant, _
                                  ByRef DataRowIndices As Variant) As String
Attribute TestSetMembership.VB_Description = "Generates set (IN) condition."
    Dim RowIdxLow As Long
    RowIdxLow = LBound(DataRowIndices)
    Dim RowIdxHigh As Long
    RowIdxHigh = UBound(DataRowIndices)
    
    Dim ColIdxLow As Long
    ColIdxLow = LBound(ColNames)
    Dim ColIdxHigh As Long
    ColIdxHigh = UBound(ColNames)
    
    '''' If there is one column only, strip unnecessary syntax
    Dim LeftParen As String
    Dim RightParen As String
    If ColIdxLow = ColIdxHigh Then
        LeftParen = vbNullString
        RightParen = vbNullString
    Else
        LeftParen = "("
        RightParen = ")"
    End If
    
    Dim FieldMap As Scripting.Dictionary
    Set FieldMap = LiteADOlib.MapFields(TableData)
    Dim FieldName As String
    Dim FieldNames As Variant
    ReDim FieldNames(ColIdxLow To ColIdxHigh)
    Dim ColIndices() As Long
    ReDim ColIndices(ColIdxLow To ColIdxHigh)
    Dim ColIdxIndex As Long
    For ColIdxIndex = ColIdxLow To ColIdxHigh
        FieldName = ColNames(ColIdxIndex)
        Guard.ExpressionErr FieldMap.Exists(FieldName), InvalidParameterErr, _
            "LiteSQLlib/VectorInValues", "Invalid field name: " & FieldName
        ColIndices(ColIdxIndex) = FieldMap(FieldName)
        FieldNames(ColIdxIndex) = SQLQuoter.QuoteSQLName(FieldName)
    Next ColIdxIndex
    FieldNames = LeftParen & Join(FieldNames, ", ") & RightParen
    
    Dim RowValues() As String
    ReDim RowValues(RowIdxLow To RowIdxHigh)
    Dim VectorValues() As String
    ReDim VectorValues(ColIdxLow To ColIdxHigh)
    
    Dim RowIdxIndex As Long
    For RowIdxIndex = RowIdxLow To RowIdxHigh
        For ColIdxIndex = ColIdxLow To ColIdxHigh
            VectorValues(ColIdxIndex) = SQLQuoter.CStrSQL( _
                TableData(DataRowIndices(RowIdxIndex) + 1, ColIndices(ColIdxIndex)))
        Next ColIdxIndex
        RowValues(RowIdxIndex) = LeftParen & Join(VectorValues, ", ") & RightParen
    Next RowIdxIndex
        
    If ColIdxLow = ColIdxHigh Then
        TestSetMembership = FieldNames & " IN (" & Join(RowValues, "," & " ") & ")"
    Else
        TestSetMembership = FieldNames & " IN (VALUES" & vbNewLine & _
            "    " & Join(RowValues, "," & vbNewLine & "    ") & ")"
    End If
End Function


'@Description "Generates SELECT with set (IN) condition."
Public Function SelectSetMembers(ByRef TableData As Variant, _
                                 ByRef ColNames As Variant, _
                                 ByRef DataRowIndices As Variant) As String
Attribute SelectSetMembers.VB_Description = "Generates SELECT with set (IN) condition."
    SelectSetMembers = "SELECT * FROM " & this.TableName & " WHERE " & _
        TestSetMembership(TableData, ColNames, DataRowIndices)
End Function


'@Description "Generates DELETE with set (IN) condition."
Public Function DeleteSetMembers(ByRef TableData As Variant, _
                                 ByRef ColNames As Variant, _
                                 ByRef DataRowIndices As Variant) As String
Attribute DeleteSetMembers.VB_Description = "Generates DELETE with set (IN) condition."
    DeleteSetMembers = "DELETE FROM " & this.TableName & " WHERE " & _
        TestSetMembership(TableData, ColNames, DataRowIndices)
End Function


'''' Examples:
''''   >>> ?LiteSQLlib("people").UpdateTemplate(FixUtils.People2D, Array("age", "last_name"))
''''   Update people
''''   SET (id, first_name, gender, email, country, domain) = (?1, ?2, ?5, ?6, ?7, ?8)
''''   WHERE age = ?4, last_name = ?3
''''
''''   >>> LiteSQLlib("people").UpdateTemplate(FixUtils.People2DSpacesInFieldNames, Array("age", "last name"))
''''   Update people
''''   SET (id, "first name", gender, email, country, domain) = (?1, ?2, ?5, ?6, ?7, ?8)
''''   WHERE age = ?4, "last name" = ?3
''''
'@Description "Generates UPDATE template."
Public Function UpdateTemplate(ByRef TableData As Variant, _
                               ByRef ColNames As Variant) As String
Attribute UpdateTemplate.VB_Description = "Generates UPDATE template."
    Dim SQLTerms(0 To 2) As String
    SQLTerms(0) = "UPDATE " & SQLQuoter.QuoteSQLName(this.TableName)
    
    Dim ColNamesLB As Long
    ColNamesLB = LBound(ColNames)
    Dim ColNamesUB As Long
    ColNamesUB = UBound(ColNames)
        
    Dim FieldMap As Scripting.Dictionary
    Set FieldMap = LiteADOlib.MapFields(TableData)
    Dim BaseShift As Long
    BaseShift = -LBound(TableData, 2) + 1
    
    Dim FieldName As String
    Dim FieldNames As Variant
    ReDim FieldNames(ColNamesLB To ColNamesUB)
    Dim ColNameIndex As Long
    For ColNameIndex = ColNamesLB To ColNamesUB
        FieldName = ColNames(ColNameIndex)
        Guard.ExpressionErr FieldMap.Exists(FieldName), InvalidParameterErr, _
            "LiteSQLlib/GetUpdateTemplate", "Invalid field name: " & FieldName
        FieldNames(ColNameIndex) = SQLQuoter.QuoteSQLName(FieldName) & " = ?" _
            & CStr(FieldMap(FieldName) + BaseShift)
        FieldMap.Remove FieldName
    Next ColNameIndex
    SQLTerms(2) = "WHERE " & Join(FieldNames, ", ")

    Dim PlaceHolders As Variant
    PlaceHolders = FieldMap.Items
    FieldNames = FieldMap.Keys
    Dim NameIndex As Long
    For NameIndex = 0 To UBound(FieldNames)
        FieldNames(NameIndex) = SQLQuoter.QuoteSQLName(FieldNames(NameIndex))
        PlaceHolders(NameIndex) = "?" & CStr(PlaceHolders(NameIndex) + BaseShift)
    Next NameIndex
    
    If UBound(FieldNames) = 0 Then
        SQLTerms(1) = "SET " & FieldNames(0) & " = " & PlaceHolders(0)
    Else
        SQLTerms(1) = "SET " & "(" & Join(FieldNames, ", ") & ") = (" & _
                                     Join(PlaceHolders, ", ") & ")"
    End If
    
    UpdateTemplate = Join(SQLTerms, vbNewLine)
End Function
