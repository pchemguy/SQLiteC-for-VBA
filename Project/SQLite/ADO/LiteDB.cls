VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LiteDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "SQLite ADO manager class"
'@Folder "SQLite.ADO"
'@ModuleDescription "SQLite ADO manager class"
'@PredeclaredId
'@IgnoreModule IndexedDefaultMemberAccess
Option Explicit

Private Type TLiteDB
    ExecCI As LiteADO
    Exec As ILiteADO
    Meta As LiteMetaADO
    Util As LiteUtils
End Type
Private this As TLiteDB


'@DefaultMember
Public Function Create(ByVal Database As String, _
              Optional ByVal AllowNonExistent As Variant = False) As LiteDB
Attribute Create.VB_UserMemId = 0
    Dim Instance As LiteDB
    Set Instance = New LiteDB
    Instance.Init Database, AllowNonExistent
    Set Create = Instance
End Function

Friend Sub Init(ByVal Database As String, _
       Optional ByVal AllowNonExistent As Variant = False)
    With this
        Set .Exec = LiteADO(Database, AllowNonExistent)
        Set .Meta = LiteMetaADO(.Exec)
        Set .Util = LiteUtils(.Exec)
        Set .ExecCI = .Exec
    End With
End Sub

Public Property Get ExecCI() As LiteADO
    Set ExecCI = this.ExecCI
End Property

Public Property Get Exec() As ILiteADO
    Set Exec = this.Exec
End Property

Public Property Get Meta() As LiteMetaADO
    Set Meta = this.Meta
End Property

Public Property Get Util() As LiteUtils
    Set Util = this.Util
End Property
