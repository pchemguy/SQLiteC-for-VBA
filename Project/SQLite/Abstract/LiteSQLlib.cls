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
Option Explicit

Private Type TObjectState
    TableName As String
    FieldNamesQ() As String
    FieldMap As Scripting.Dictionary
    FieldData As Boolean
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
    this.FieldData = False
    Set this.FieldMap = New Scripting.Dictionary
    this.FieldMap.CompareMode = TextCompare
End Sub


'@Ignore ProcedureNotUsed
Public Property Get TableName() As String
    TableName = this.TableName
End Property


Public Property Let TableName(ByVal Values As String)
    this.TableName = Values
    this.FieldData = False
    this.FieldMap.RemoveAll
End Property


'''' Extracts field names from 2D array with header.
''''
'''' Extracts field names from the header row of a 2D array table. If any field
'''' name contains double qoute, an error is raised. If any field name contains
'''' a space, the name id double quoted.
''''
'''' Args:
''''   TableData (array):
''''     2D array with data and field names in the first row.
''''
'''' Returns:
''''   1D array of field names.
''''
'''' Examples:
''''   >>> ?Join(FieldNamesFrom2DArrayHeader(FixPeopleData.UsedRange.Value2), "#")
''''   "id#first_name#last_name#age#gender#email#country#domain"
''''
'@Description "Extracts field names from 2D array with header."
Public Function FieldNamesFrom2DArrayHeader(ByRef TableData As Variant) As String()
Attribute FieldNamesFrom2DArrayHeader.VB_Description = "Extracts field names from 2D array with header."
    If this.FieldData Then
        FieldNamesFrom2DArrayHeader = this.FieldNamesQ
        Exit Function
    End If
    
    Dim FieldNames() As String
    ReDim FieldNames(1 To UBound(TableData, 2) - LBound(TableData, 2) + 1)
    Dim FieldIndexOffset As Long
    FieldIndexOffset = LBound(TableData, 2) - LBound(FieldNames)
    Dim FieldName As String
    Dim FieldIndex As Long
    For FieldIndex = LBound(FieldNames) To UBound(FieldNames)
        FieldName = TableData(1, FieldIndex + FieldIndexOffset)
        FieldNames(FieldIndex) = SQLlib.QuoteSQLName(FieldName)
    Next FieldIndex
    this.FieldNamesQ = FieldNames
    this.FieldData = True
    FieldNamesFrom2DArrayHeader = FieldNames
End Function


'''' Args:
''''   TableData (array):
''''     2D array with data and field names in the first row.
''''
'@Description "Generates the INSERT OR IGNORE INTO VALUES query for bulk insert."
Public Function DefaultUPSERT(ByRef TableData As Variant) As String
Attribute DefaultUPSERT.VB_Description = "Generates the INSERT OR IGNORE INTO VALUES query for bulk insert."
    Dim FieldNames() As String
    If Not this.FieldData Then
        FieldNames = FieldNamesFrom2DArrayHeader(TableData)
    Else
        FieldNames = this.FieldNamesQ
    End If
    
    Dim FieldTerms() As String
    ReDim FieldTerms(LBound(FieldNames) To UBound(FieldNames))
    Dim FieldName As String
    Dim FieldIndex As Long
    For FieldIndex = LBound(FieldNames) To UBound(FieldNames)
        FieldName = FieldNames(FieldIndex)
        FieldTerms(FieldIndex) = FieldName & "=" & "excluded." & FieldName
    Next FieldIndex
    DefaultUPSERT = "ON CONFLICT DO UPDATE SET" & vbNewLine & _
                    "    " & Join(FieldTerms, "," & vbNewLine & "    ")
End Function


'''' Args:
''''   TableData (array):
''''     2D array with data and field names in the first row.
''''
''''   DataRowIndices (array, optional, Empty):
''''     1D array of data row indices to be used. The 1st data row (the 2nd
''''     row in the TableData) has index of 1. If empty, all rows are used.
''''
'@Description "Generates an array of record values for INSERT INTO VALUES"
Public Function RecordValuesFrom2DArray( _
                     ByRef TableData As Variant, _
            Optional ByVal DataRowIndices As Variant = Empty) As String()
Attribute RecordValuesFrom2DArray.VB_Description = "Generates an array of record values for INSERT INTO VALUES"
    If Not (IsArray(DataRowIndices) Or IsEmpty(DataRowIndices)) Then
        Err.Raise ErrNo.InvalidParameterErr, "SQLlib/RecordValuesFrom2DArray", _
                  "The second argument should be array or Empty/missing."
    End If
    Dim FieldUB As Long
    FieldUB = UBound(TableData, 2)
    Dim FieldLB As Long
    FieldLB = LBound(TableData, 2)
    Dim RecordUB As Long
    RecordUB = UBound(TableData, 1)
    Dim RecordLB As Long
    '''' +1 due to the header row
    RecordLB = LBound(TableData, 1)
    
    Dim FieldValues() As String
    ReDim FieldValues(FieldLB To FieldUB)
    Dim RecordValues() As String
    Dim FieldValue As Variant
    
    Dim RecordSrcIndex As Long
    Dim RecordDstIndex As Long
    If IsEmpty(DataRowIndices) Then
        ReDim RecordValues(RecordLB To RecordUB - 1)
        For RecordDstIndex = RecordLB To RecordUB - 1
            RecordSrcIndex = RecordDstIndex + 1
            GoSub RECORD_LOOP
        Next RecordDstIndex
    Else
        ReDim RecordValues(LBound(DataRowIndices) To UBound(DataRowIndices))
        For RecordDstIndex = LBound(DataRowIndices) To UBound(DataRowIndices)
            RecordSrcIndex = RecordLB + DataRowIndices(RecordDstIndex)
            GoSub RECORD_LOOP
        Next RecordDstIndex
    End If
    RecordValuesFrom2DArray = RecordValues

    Exit Function
RECORD_LOOP:
    Dim FieldIndex As Long
    For FieldIndex = FieldLB To FieldUB
        FieldValue = TableData(RecordSrcIndex, FieldIndex)
        FieldValues(FieldIndex) = SQLlib.CStrSQL(FieldValue)
    Next FieldIndex
    RecordValues(RecordDstIndex) = "(" & Join(FieldValues, ", ") & ")"
    Return
End Function


'''' Args:
''''   TableData (array):
''''     2D array with data and field names in the first row.
''''
''''   DataRowIndices (array, optional, Empty):
''''     1D array of data row indices to be used. The 1st data row (the 2nd
''''     row in the TableData) has index of 1. If empty, all rows are used.
''''
'@Description "Generates the INSERT OR IGNORE INTO VALUES query for bulk insert."
Public Function InsertNewRecordsFrom2DArray( _
                     ByRef TableData As Variant, _
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


'''' Args:
''''   TableData (array):
''''     2D array with data and field names in the first row.
''''
''''   DataRowIndices (array, optional, Empty):
''''     1D array of data row indices to be used. The 1st data row (the 2nd
''''     row in the TableData) has index of 1. If empty, all rows are used.
''''
'@Description "Generates the INSERT OR UPDATE INTO VALUES query for bulk insert."
Public Function UpsertFrom2DArray( _
                     ByRef TableData As Variant, _
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


'''' Args:
''''   TableData (array):
''''     2D array with data and field names in the first row.
''''
'@Description "Generates field map from 2D data array with header."
Public Function MapFields(ByRef TableData As Variant) As Scripting.Dictionary
    Dim FieldMap As Scripting.Dictionary
    If Not this.FieldMap Is Nothing Then
        If this.FieldMap.Count > 0 Then
            Set MapFields = this.FieldMap
            Exit Function
        Else
            Set FieldMap = this.FieldMap
        End If
    Else
        Set FieldMap = New Scripting.Dictionary
    End If
    
    Dim HeaderRowIndex As Long
    HeaderRowIndex = LBound(TableData, 1)
    Dim FieldIndex As Long
    For FieldIndex = LBound(TableData, 2) To UBound(TableData, 2)
        FieldMap(TableData(HeaderRowIndex, FieldIndex)) = FieldIndex
    Next FieldIndex
    Set MapFields = FieldMap
End Function


'''' Generates set membership test clause.
''''
'''' Args:
''''   TableData (array):
''''     2D array with data and field names in the first row.
''''
''''   ColNames (array):
''''     1D array of field names to be included. The order may be
''''     important in terms of performance. When there is a corresponding
''''     usable multi-field index, it is better to supply field names
''''     according to their appearance in the index.
''''
''''   RowIndices (array):
''''     1D array of record indices. The 1st data row (2nd row in the
''''     TableData array) has index of LBound(TableData, 2) (one is
''''     added by the code.
''''
'''' Examples:
''''   >>> ?SQLlib.TestSetMembership(SQLlib.People2D, Array("country", "age"), Array(2, 3))
''''   "(country, age) IN (VALUES
''''       ('Kazakhstan', '62'),
''''       ('Belarus', '26'))"
''''
''''   >>> ?SQLlib.TestSetMembership(SQLlib.People2D, Array("country"), Array(2, 3))
''''   "country IN ('Kazakhstan', 'Belarus')"
''''
'@Description "Generates set (IN) condition."
Public Function TestSetMembership(ByRef TableData As Variant, _
                                  ByRef ColNames As Variant, _
                                  ByRef RowIndices As Variant) As String
    Dim RowIdxLow As Long
    RowIdxLow = LBound(RowIndices)
    Dim RowIdxHigh As Long
    RowIdxHigh = UBound(RowIndices)
    
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
    Set FieldMap = MapFields(TableData)
    Dim FieldName As String
    Dim FieldNames As Variant
    ReDim FieldNames(ColIdxLow To ColIdxHigh)
    Dim ColIndices() As Long
    ReDim ColIndices(ColIdxLow To ColIdxHigh)
    Dim ColIdxIndex As Long
    For ColIdxIndex = ColIdxLow To ColIdxHigh
        FieldName = ColNames(ColIdxIndex)
        Guard.ExpressionErr FieldMap.Exists(FieldName), InvalidParameterErr, _
            "SQLlib/VectorInValues", "Invalid field name: " & FieldName
        ColIndices(ColIdxIndex) = FieldMap(FieldName)
        FieldNames(ColIdxIndex) = SQLlib.QuoteSQLName(FieldName)
    Next ColIdxIndex
    FieldNames = LeftParen & Join(FieldNames, ", ") & RightParen
    
    Dim RowValues() As String
    ReDim RowValues(RowIdxLow To RowIdxHigh)
    Dim VectorValues() As String
    ReDim VectorValues(ColIdxLow To ColIdxHigh)
    
    Dim RowIdxIndex As Long
    For RowIdxIndex = RowIdxLow To RowIdxHigh
        For ColIdxIndex = ColIdxLow To ColIdxHigh
            VectorValues(ColIdxIndex) = SQLlib.CStrSQL(TableData(RowIndices(RowIdxIndex) + 1, ColIndices(ColIdxIndex)))
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


'''' Args:
''''   See TestSetMembership
''''
'@Description "Generates SELECT with set (IN) condition."
Public Function SelectSetMembers(ByRef TableData As Variant, _
                                 ByRef ColNames As Variant, _
                                 ByRef RowIndices As Variant) As String
    Guard.DefaultInstance Me
    SelectSetMembers = "SELECT * FROM " & this.TableName & " WHERE " & _
        TestSetMembership(TableData, ColNames, RowIndices)
End Function


'''' Args:
''''   See TestSetMembership
''''
'@Description "Generates DELETE with set (IN) condition."
Public Function DeleteSetMembers(ByRef TableData As Variant, _
                                 ByRef ColNames As Variant, _
                                 ByRef RowIndices As Variant) As String
    Guard.DefaultInstance Me
    DeleteSetMembers = "DELETE FROM " & this.TableName & " WHERE " & _
        TestSetMembership(TableData, ColNames, RowIndices)
End Function
