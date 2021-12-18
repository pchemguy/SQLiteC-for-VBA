VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LiteADOlib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Provides common database operations via the ILiteADO interface."
'@Folder "SQLite.Abstract"
'@ModuleDescription "Provides common database operations via the ILiteADO interface."
'@PredeclaredId
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
    dbq As ILiteADO
    SQLTool As LiteSQLlib
End Type
Private this As TObjectState


'@DefaultMember
Public Function Create(ByVal dbq As ILiteADO) As LiteADOlib
Attribute Create.VB_UserMemId = 0
    Dim Instance As LiteADOlib
    Set Instance = New LiteADOlib
    Instance.Init dbq
    Set Create = Instance
End Function


Friend Sub Init(ByVal dbq As ILiteADO)
    Guard.NullReference dbq
    Set this.dbq = dbq
    Set this.SQLTool = LiteSQLlib(vbNullString)
End Sub


'@Description "Inserts data rows skipping existing."
Public Function InsertSkipExistingFrom2D( _
                     ByVal TableName As String, _
                     ByVal TableData As Variant, _
            Optional ByVal DataRowIndices As Variant = Empty) As Long
Attribute InsertSkipExistingFrom2D.VB_Description = "Inserts data rows skipping existing."
    Guard.EmptyString TableName
    Guard.EmptyOrNullVariant TableData
    With this.SQLTool
        .TableName = TableName
        Dim SQLQuery As String
        SQLQuery = .InsertNewRecordsFrom2DArray(TableData, DataRowIndices)
    End With
    InsertSkipExistingFrom2D = this.dbq.ExecuteNonQuery(SQLQuery)
End Function


'@Description "Inserts data rows skipping existing."
Public Function InsertUpdateExistingFrom2D( _
                     ByVal TableName As String, _
                     ByVal TableData As Variant, _
            Optional ByVal DataRowIndices As Variant = Empty) As Long
Attribute InsertUpdateExistingFrom2D.VB_Description = "Inserts data rows skipping existing."
    Guard.EmptyString TableName
    Guard.EmptyOrNullVariant TableData
    With this.SQLTool
        .TableName = TableName
        Dim SQLQuery As String
        SQLQuery = .UpsertFrom2DArray(TableData, DataRowIndices)
    End With
    InsertUpdateExistingFrom2D = this.dbq.ExecuteNonQuery(SQLQuery)
End Function


'@Description "Selects and returns a headerless 2D array containing a subset of data rows."
Public Function SelectSubsetFrom2D( _
                     ByVal TableName As String, _
                     ByVal TableData As Variant, _
                     ByRef ColNames As Variant, _
                     ByRef DataRowIndices As Variant) As Variant
Attribute SelectSubsetFrom2D.VB_Description = "Selects and returns a headerless 2D array containing a subset of data rows."
    Guard.EmptyString TableName
    Guard.EmptyOrNullVariant TableData
    Guard.EmptyOrNullVariant ColNames
    Guard.EmptyOrNullVariant DataRowIndices
    With this.SQLTool
        .TableName = TableName
        Dim SQLQuery As String
        SQLQuery = .SelectSetMembers(TableData, ColNames, DataRowIndices)
    End With
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = this.dbq.GetAdoRecordset(SQLQuery)
    '@Ignore ValueRequired: False positive
    If AdoRecordset.RecordCount > 0 Then
        SelectSubsetFrom2D = RecordsetToHeaded2D(AdoRecordset)
    Else
        SelectSubsetFrom2D = Empty
    End If
End Function


'@Description "Deletes a subset of data rows."
Public Function DeleteSubsetFrom2D( _
                     ByVal TableName As String, _
                     ByVal TableData As Variant, _
                     ByRef ColNames As Variant, _
                     ByRef DataRowIndices As Variant) As Long
Attribute DeleteSubsetFrom2D.VB_Description = "Deletes a subset of data rows."
    Guard.EmptyString TableName
    Guard.EmptyOrNullVariant TableData
    Guard.EmptyOrNullVariant ColNames
    Guard.EmptyOrNullVariant DataRowIndices
    With this.SQLTool
        .TableName = TableName
        Dim SQLQuery As String
        SQLQuery = .DeleteSetMembers(TableData, ColNames, DataRowIndices)
    End With
    DeleteSubsetFrom2D = this.dbq.ExecuteNonQuery(SQLQuery)
End Function


'@Description "Converts an ADODB Recordset to a 2D data array with field names in the first row."
Public Function RecordsetToHeaded2D(ByVal AdoRecordset As ADODB.Recordset) As Variant
    Guard.NullReference AdoRecordset
    
    Dim RecordCount As Long
    '@Ignore ValueRequired: False positive
    RecordCount = AdoRecordset.RecordCount
    If RecordCount = 0 Then
        RecordsetToHeaded2D = Empty
        Exit Function
    End If
        
    Dim FieldCount As Long
    FieldCount = AdoRecordset.Fields.Count
    Dim TableData() As Variant
    '''' Count the header row
    ReDim TableData(1 To RecordCount + 1, 1 To FieldCount)
    AdoRecordset.MoveFirst
    With AdoRecordset.Fields
        Dim FieldIndex As Long
        Dim RecordIndex As Long
        RecordIndex = 1
        For FieldIndex = 1 To FieldCount
            TableData(RecordIndex, FieldIndex) = .Item(FieldIndex - 1).Name
        Next FieldIndex
        For RecordIndex = 2 To RecordCount + 1
            For FieldIndex = 1 To FieldCount
                TableData(RecordIndex, FieldIndex) = .Item(FieldIndex - 1).Value
            Next FieldIndex
            AdoRecordset.MoveNext
        Next RecordIndex
    End With
    AdoRecordset.MoveFirst
    RecordsetToHeaded2D = TableData
End Function


'''' Reduced type set is used for mapping according to type mapping used with SQLite
'@Description "Converts a 2D data array with field names in the first row to a fabricated ADODB Recordset."
Public Function Headed2DToRecordset(ByVal TableData As Variant, _
                           Optional ByVal InferredTypesRecordOffset = 1) As ADODB.Recordset
    Guard.NotArray TableData
    
    Dim RowLB As Long
    RowLB = LBound(TableData, 1)
    Dim RowUB As Long
    RowUB = UBound(TableData, 1)
    Dim ColLB As Long
    ColLB = LBound(TableData, 2)
    Dim ColUB As Long
    ColUB = UBound(TableData, 2)
    
    Dim FabRecordset As ADODB.Recordset
    Set FabRecordset = New ADODB.Recordset
    Dim OriginName As String
    Dim AdoType As ADODB.DataTypeEnum
    Dim AdoAttr As ADODB.FieldAttributeEnum
    Dim AdoSize As Long

    With FabRecordset.Fields
        Dim Value As Variant
        Dim TypesRecordIndex As Long
        TypesRecordIndex = RowLB + InferredTypesRecordOffset
    
        Dim FieldIndex As Long
        Dim RecordIndex As Long
        RecordIndex = RowLB
        For FieldIndex = ColLB To ColUB
            OriginName = TableData(RecordIndex, FieldIndex)
            AdoAttr = adFldUpdatable Or adFldIsNullable Or adFldMayBeNull
            Value = TableData(TypesRecordIndex, FieldIndex)
            Select Case VarType(Value)
                Case vbString, vbBoolean, vbDate
                    AdoType = adVarWChar
                    AdoSize = 2 ^ 13
                Case vbLong, vbInteger, vbByte, vbLongLong
                    AdoType = adInteger
                    AdoSize = 8
                Case vbDouble, vbSingle, vbCurrency
                    AdoType = adDouble
                    AdoSize = 8
                Case Else
                    AdoType = adVarWChar
                    AdoSize = 2 ^ 16 - 1
            End Select
            .Append OriginName, AdoType, AdoSize, AdoAttr
        Next FieldIndex
        
        FabRecordset.Open
        For RecordIndex = RowLB + 1 To RowUB
            FabRecordset.AddNew
            For FieldIndex = ColLB To ColUB
                Value = TableData(RecordIndex, FieldIndex)
                If VarType(Value) = vbDate Then _
                    Value = Format(Value, "YYYY-MM-DDTHH:MM:SS")
                .Item(FieldIndex - ColLB).Value = Value
            Next FieldIndex
            FabRecordset.Update
        Next RecordIndex
    End With
    FabRecordset.UpdateBatch
    If FabRecordset.RecordCount > 0 Then FabRecordset.MoveFirst
    Set Headed2DToRecordset = FabRecordset
End Function
