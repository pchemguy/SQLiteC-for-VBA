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


'@Description "Absolute file pathname to the main database."
Public Property Get MainDB() As String
Attribute MainDB.VB_Description = "Absolute file pathname to the main database."
End Property

'@Description "Connection string used by the implementing package."
Public Property Get ConnectionString() As String
Attribute ConnectionString.VB_Description = "Connection string used by the implementing package."
End Property

'''' A variant error value may be returned in case of an issue
Public Function GetScalar(ByVal SQLQuery As String, _
                 Optional ByVal QueryParameters As Variant = Empty) As Variant
End Function

Public Function GetAdoRecordset(ByVal SQLQuery As String, _
                       Optional ByVal QueryParameters As Variant = Empty) As ADODB.Recordset
End Function

'@Description "Returns RecordsAffected count"
Public Function ExecuteNonQuery(ByVal SQLQuery As String, _
                       Optional ByVal QueryParameters As Variant = Empty) As Long
Attribute ExecuteNonQuery.VB_Description = "Returns RecordsAffected count"
End Function

'@Description "Begins SQLite transaction."
Public Sub Begin(Optional ByVal Mode As SQLiteTxnType = SQLITE_TXN_DEFERRED)
Attribute Begin.VB_Description = "Begins SQLite transaction."
End Sub

'@Description "Commits all SQLite DBO transactions."
Public Sub Commit()
Attribute Commit.VB_Description = "Commits all SQLite DBO transactions."
End Sub

'@Description "Begins a named nestable SQLite transaction."
Public Sub SavePoint(ByVal SavePointName As String)
Attribute SavePoint.VB_Description = "Begins a named nestable SQLite transaction."
End Sub

'@Description "Removes rollback savepoints from the transaction stack up to the first point with matching name."
Public Sub ReleasePoint(ByVal SavePointName As String)
Attribute ReleasePoint.VB_Description = "Removes rollback savepoints from the transaction stack up to the first point with matching name."
End Sub

'@Description "Rollbacks savepoints from the transaction stack up to the first point with matching name or the entire transaction."
Public Sub Rollback(Optional ByVal SavePointName As String = vbNullString)
Attribute Rollback.VB_Description = "Rollbacks savepoints from the transaction stack up to the first point with matching name or the entire transaction."
End Sub

'''' Returns 0 on success and non-zero on failure."
'@Description "Opens database connection"
Public Function OpenDb() As Long
End Function

'''' Returns 0 on success and non-zero on failure."
'@Description "Closes database connection"
Public Function CloseDb() As Long
End Function
