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
    FieldNamesUsable As Boolean
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
    this.FieldNamesUsable = False
End Sub


'@Ignore ProcedureNotUsed
Public Property Get TableName() As String
    TableName = this.TableName
End Property


Public Property Let TableName(ByVal Values As String)
    this.TableName = Values
    this.FieldNamesUsable = False '''' Quoted if necessary (not very solid)
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
    If this.FieldNamesUsable Then
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
    this.FieldNamesUsable = True
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
    If Not this.FieldNamesUsable Then
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


'''' This routine splits a semicolon separated multi-statement SQL query into
'''' subqueries. Since semicolons may be a part of a string value, direct
'''' splitting cannot be used. The code scans the query and tracks single and
'''' double quotation marks and square brackets. Only one of these can be
'''' "opened" at any given position. Quotation marks may be escaped by doubling
'''' or by a precedening odd number of backslashes. If quotation mark is escaped,
'''' it is treated literally. When a semicolon is encountered, a split point is
'''' added if all three tracked sets are closed. Otherwise, such a semicolon is
'''' skipped.
''''
'''' Args:
''''   SQLQuery (string):
''''     SQL query to be split
''''
'''' Returns:
''''   Array, containing individual subqueries
''''
'''' '''' "SELECT 1; SELECT ';''\''; ;SELECT 2; SELECT ''';[""]'; SELECT "
''''
'@Description "Splits multi-statement SQL query into subqueries."
Public Function SplitMultiStatement(ByRef SQLQuery As String) As Variant
Attribute SplitMultiStatement.VB_Description = "Splits multi-statement SQL query into subqueries."
    Const SingleQuoteASCII As Byte = 39
    Const DoubleQuoteASCII As Byte = 34
    Const LeftBracketASCII As Byte = 91
    Const RightBracketASCII As Byte = 93
    Const BackSlashASCII As Byte = 92
    Const SemicolonASCII As Byte = 59
    
    If SQLQuery = vbNullString Then
        SplitMultiStatement = Array(vbNullString)
        Exit Function
    End If
    
    Dim QueryChars() As Byte
    QueryChars = StrConv(" " & SQLQuery & " ", vbFromUnicode)
    
    Dim SplitPositions As Scripting.Dictionary
    Set SplitPositions = New Scripting.Dictionary
    SplitPositions.CompareMode = TextCompare
    
    Dim SQuoteOpen As Boolean
    SQuoteOpen = False
    Dim DQuoteOpen As Boolean
    DQuoteOpen = False
    Dim BracketOpen As Boolean
    BracketOpen = False
    
    Dim BackSlashed As Boolean
    
    Dim CurChar As Byte
    Dim ChrJ As Long
    Dim CharIndex As Long
    For CharIndex = 1 To Len(SQLQuery)
        CurChar = QueryChars(CharIndex)
        Select Case CurChar
            Case LeftBracketASCII, RightBracketASCII
                If Not (SQuoteOpen Or DQuoteOpen) Then _
                    BracketOpen = Not BracketOpen
            Case SingleQuoteASCII
                '''' Bail immediately if the other quote or bracket is open
                If Not (DQuoteOpen Or BracketOpen) Then
                    '''' Char is backslashed if preceded by an odd number of "\".
                    BackSlashed = False
                    If QueryChars(CharIndex - 1) = BackSlashASCII Then
                        ChrJ = CharIndex
                        Do While QueryChars(ChrJ - 1) = BackSlashASCII
                            ChrJ = ChrJ - 1
                        Loop
                        BackSlashed = CBool((CharIndex - ChrJ) Mod 2)
                    End If
                    
                    '''' If quote is repeated, shift CharIndex to the last one
                    '''' and flip the quote state if the total count is odd
                    If Not BackSlashed Then
                        ChrJ = CharIndex - 1
                        Do While QueryChars(CharIndex + 1) = SingleQuoteASCII
                            CharIndex = CharIndex + 1
                        Loop
                        If CBool((CharIndex - ChrJ) Mod 2) Then _
                            SQuoteOpen = Not SQuoteOpen
                    End If
                End If
            Case DoubleQuoteASCII
                '''' Bail immediately if the other quote or bracket is open
                If Not (SQuoteOpen Or BracketOpen) Then
                    '''' Char is backslashed if preceded by an odd number of "\".
                    BackSlashed = False
                    If QueryChars(CharIndex - 1) = BackSlashASCII Then
                        ChrJ = CharIndex
                        Do While QueryChars(ChrJ - 1) = BackSlashASCII
                            ChrJ = ChrJ - 1
                        Loop
                        BackSlashed = CBool((CharIndex - ChrJ) Mod 2)
                    End If
                    
                    '''' If quote is repeated, shift CharIndex to the last one
                    '''' and flip the quote state if the total count is odd
                    If Not BackSlashed Then
                        ChrJ = CharIndex - 1
                        Do While QueryChars(CharIndex + 1) = DoubleQuoteASCII
                            CharIndex = CharIndex + 1
                        Loop
                        If CBool((CharIndex - ChrJ) Mod 2) Then _
                            DQuoteOpen = Not DQuoteOpen
                    End If
                End If
            Case SemicolonASCII
                If Not (SQuoteOpen Or DQuoteOpen Or BracketOpen) Then
                    SplitPositions(SplitPositions.Count) = CharIndex
                End If
        End Select
    Next CharIndex
    If SplitPositions.Count = 0 Then
        SplitPositions(0) = Len(SQLQuery) + 1
    ElseIf SplitPositions(SplitPositions.Count - 1) < Len(SQLQuery) Then
        SplitPositions(SplitPositions.Count) = Len(SQLQuery) + 1
    End If
    Dim SplitPosArr As Variant
    SplitPosArr = SplitPositions.Items
    
    Dim SQLStmtStore As Scripting.Dictionary
    Set SQLStmtStore = New Scripting.Dictionary
    SQLStmtStore.CompareMode = TextCompare
    
    Dim LeftPos As Long
    LeftPos = 1
    Dim RightPos As Long
    Dim SubStmt As String
    
    Dim SubStmtIndex As Long
    For SubStmtIndex = LBound(SplitPosArr) To UBound(SplitPosArr)
        RightPos = SplitPosArr(SubStmtIndex)
        SubStmt = Mid$(SQLQuery, LeftPos, RightPos - LeftPos)
        SubStmt = Trim$(Replace(SubStmt, ";" & vbNewLine, "; "))
        If Len(SubStmt) > 0 Then
            SQLStmtStore(SQLStmtStore.Count) = SubStmt
        End If
        LeftPos = RightPos + 1
    Next SubStmtIndex
    
    SplitMultiStatement = SQLStmtStore.Items
End Function
