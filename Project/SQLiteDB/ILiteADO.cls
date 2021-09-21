VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILiteADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder "SQLiteDB"
'@Interface
'@Exposed
Option Explicit


Public Property Get MainDB() As String
End Property

Public Property Get ConnectionString() As String
End Property

Public Property Get AdoConnection() As ADODB.Connection
End Property

Public Property Get AdoCommand() As ADODB.Command
End Property

Public Function GetAdoRecordset(Optional ByVal SQLQuery As String = vbNullString) As ADODB.Recordset
End Function

Public Function GetScalar(Optional ByVal SQLQuery As String = vbNullString) As Variant
End Function

Public Function ExecuteNonQuery(ByVal SQLQuery As String) As Long
End Function

