VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILiteADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder "SQLite.Abstract"
'@Interface
'@Exposed
Option Explicit


'''' Absolute file pathname to the main database.
Public Property Get MainDB() As String
End Property

'''' Connection string used by the implementing package.
Public Property Get ConnectionString() As String
End Property

'''' A variant error value may be returned in case of an issue
Public Function GetScalar(ByVal SQLQuery As String, _
                 Optional ByVal QueryParameters As Variant = Empty) As Variant
End Function

Public Function GetAdoRecordset(ByVal SQLQuery As String, _
                       Optional ByVal QueryParameters As Variant = Empty) As ADODB.Recordset
End Function

'''' Returns RecordsAffected count
Public Function ExecuteNonQuery(ByVal SQLQuery As String, _
                       Optional ByVal QueryParameters As Variant = Empty) As Long
End Function

'@Description "Begins SQLite transaction."
Public Sub Begin(Optional ByVal Mode As SQLiteTxnType = SQLITE_TXN_DEFERRED)
End Sub

'@Description "Commits all SQLite DBO transactions."
Public Sub Commit()
End Sub

'@Description "Begins a named nestable SQLite transaction."
Public Sub SavePoint(ByVal SavePointName As String)
End Sub

'@Description "Removes rollback savepoints from the transaction stack up to the first point with matching name."
Public Sub ReleasePoint(ByVal SavePointName As String)
End Sub

'@Description "Rollbacks savepoints from the transaction stack up to the first point with matching name or the entire transaction."
Public Sub Rollback(Optional ByVal SavePointName As String = vbNullString)
End Sub
