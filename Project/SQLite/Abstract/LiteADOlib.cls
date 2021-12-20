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


'@Description "Generates field map from 2D data array with header."
Public Function MapFields(ByRef TableData As Variant) As Scripting.Dictionary
Attribute MapFields.VB_Description = "Generates field map from 2D data array with header."
    Dim FieldMap As Scripting.Dictionary
    Set FieldMap = New Scripting.Dictionary
    FieldMap.CompareMode = TextCompare
    
    Dim HeaderRowIndex As Long
    HeaderRowIndex = LBound(TableData, 1)
    Dim FieldIndex As Long
    For FieldIndex = LBound(TableData, 2) To UBound(TableData, 2)
        FieldMap(TableData(HeaderRowIndex, FieldIndex)) = FieldIndex
    Next FieldIndex
    Set MapFields = FieldMap
End Function


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
        SelectSubsetFrom2D = LiteConvUtils.RecordsetToHeaded2D(AdoRecordset)
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


'@Description "Updates selected rows."
Public Function UpdateFrom2D( _
                     ByVal TableName As String, _
                     ByVal TableData As Variant, _
            Optional ByVal DataRowIndices As Variant = Empty) As Long
    Guard.EmptyString TableName
    Guard.EmptyOrNullVariant TableData
    With this.SQLTool
        .TableName = TableName
        Dim SQLQuery As String
        SQLQuery = .UpdateTemplate(TableData, DataRowIndices)
    End With
    With this.dbq
        Dim HandleConnection As Boolean
        If .ConnectionClosed Then HandleConnection = True
        If HandleConnection Then .OpenDb
        .Begin
        Dim RecordJndex As Long
        
        .Commit
        If HandleConnection Then .CloseDb
    End With
    '= this.dbq.ExecuteNonQuery(SQLQuery)
End Function
