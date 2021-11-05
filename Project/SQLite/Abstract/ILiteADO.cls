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

