VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LiteConvUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Provides shaping and conversion utilities for data structures and queries."
'@Folder "SQLite.Abstract"
'@ModuleDescription "Provides shaping and conversion utilities for data structures and queries."
'@PredeclaredId
Option Explicit


'@Description "Converts an ADODB Recordset to a 2D data array with field names in the first row."
Public Function RecordsetToHeaded2D(ByVal AdoRecordset As ADODB.Recordset) As Variant
Attribute RecordsetToHeaded2D.VB_Description = "Converts an ADODB Recordset to a 2D data array with field names in the first row."
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
                           Optional ByVal InferredTypesRecordOffset As Long = 1 _
                           ) As ADODB.Recordset
Attribute Headed2DToRecordset.VB_Description = "Converts a 2D data array with field names in the first row to a fabricated ADODB Recordset."
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
            '@Ignore ArgumentWithIncompatibleObjectType: False positive
            .Append OriginName, AdoType, AdoSize, AdoAttr
        Next FieldIndex
        
        FabRecordset.Open
        For RecordIndex = RowLB + 1 To RowUB
            FabRecordset.AddNew
            For FieldIndex = ColLB To ColUB
                Value = TableData(RecordIndex, FieldIndex)
                If VarType(Value) = vbDate Then _
                    Value = Format$(Value, "YYYY-MM-DDTHH:MM:SS")
                .Item(FieldIndex - ColLB).Value = Value
            Next FieldIndex
            FabRecordset.Update
        Next RecordIndex
    End With
    FabRecordset.UpdateBatch
    '@Ignore ValueRequired: Flase positive
    If FabRecordset.RecordCount > 0 Then FabRecordset.MoveFirst
    Set Headed2DToRecordset = FabRecordset
End Function
