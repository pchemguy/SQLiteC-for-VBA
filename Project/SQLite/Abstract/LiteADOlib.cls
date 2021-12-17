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


'''' Args:
''''   TableName (string):
''''     Name of the destination table, possibly with schema alias.
''''
''''   TableData (array):
''''     2D array with data and field names in the first row.
''''
''''   DataRowIndices (array, optional, Empty):
''''     1D array of data row indices to be used. The 1st data row (the 2nd
''''     row in the TableData) has index of 1. If empty, all rows are used.
''''
'@Description "Inserts data rows skipping existing."
Public Function InsertSkipExistingFrom2D( _
                     ByVal TableName As String, _
                     ByVal Data2DWithHeaderRow As Variant, _
            Optional ByVal DataRowIndices As Variant = Empty) As Long
Attribute InsertSkipExistingFrom2D.VB_Description = "Inserts data rows skipping existing."
    Guard.EmptyString TableName
    Guard.EmptyOrNullVariant Data2DWithHeaderRow
    
    With this.SQLTool
        .TableName = TableName
        Dim SQLQuery As String
        SQLQuery = .InsertNewRecordsFrom2DArray(Data2DWithHeaderRow, DataRowIndices)
    End With
    InsertSkipExistingFrom2D = this.dbq.ExecuteNonQuery(SQLQuery)
End Function


'''' Args:
''''   TableName (string):
''''     Name of the destination table, possibly with schema alias.
''''
''''   TableData (array):
''''     2D array with data and field names in the first row.
''''
''''   DataRowIndices (array, optional, Empty):
''''     1D array of data row indices to be used. The 1st data row (the 2nd
''''     row in the TableData) has index of 1. If empty, all rows are used.
''''
'@Description "Inserts data rows skipping existing."
Public Function InsertUpdateExistingFrom2D( _
                     ByVal TableName As String, _
                     ByVal Data2DWithHeaderRow As Variant, _
            Optional ByVal DataRowIndices As Variant = Empty) As Long
Attribute InsertUpdateExistingFrom2D.VB_Description = "Inserts data rows skipping existing."
    Guard.EmptyString TableName
    Guard.EmptyOrNullVariant Data2DWithHeaderRow
    
    With this.SQLTool
        .TableName = TableName
        Dim SQLQuery As String
        SQLQuery = .UpsertFrom2DArray(Data2DWithHeaderRow, DataRowIndices)
    End With
    InsertUpdateExistingFrom2D = this.dbq.ExecuteNonQuery(SQLQuery)
End Function
