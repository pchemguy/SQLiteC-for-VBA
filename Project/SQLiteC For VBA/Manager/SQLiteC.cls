VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLiteC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Provides common workflows for SQLite db interactions"
'@Folder "SQLiteC For VBA.Manager"
'@ModuleDescription "Provides common workflows for SQLite db interactions"
'@PredeclaredId
Option Explicit

Private Type TSQLiteC
    DllMan As DllManager
    Connection As SQLiteCConnection
End Type
Private this As TSQLiteC


Friend Property Get DllMan() As DllManager
    Set DllMan = this.DllMan
End Property


Friend Property Get Connection() As SQLiteCConnection
    Set Connection = this.Connection
End Property


'@DefaultMember
Public Function Create(ByVal DllPath As String, _
              Optional ByVal DllNames As Variant = Empty) As SQLiteC
Attribute Create.VB_UserMemId = 0
    Dim Instance As SQLiteC
    Set Instance = New SQLiteC
    Instance.Init DllPath, DllNames
    Set Create = Instance
End Function


Friend Sub Init(ByVal DllPath As String, _
       Optional ByVal DllNames As Variant = Empty)
    Dim FileNames As Variant
    If Not IsEmpty(DllNames) Then
        FileNames = DllNames
    Else
        #If WIN64 Then
            FileNames = "sqlite3.dll"
        #Else
            FileNames = Array("icudt68.dll", "icuuc68.dll", "icuin68.dll", _
                              "icuio68.dll", "icutu68.dll", "sqlite3.dll")
        #End If
    End If
    '''' DllManager is responsible for DllPath validation/resolution
    Set this.DllMan = DllManager(DllPath, FileNames)
End Sub


Public Function Connect(ByVal DbPathName As String) As SQLiteCConnection
    '''' SQLiteCConnection is responsible for DbPathName validation/resolution
    Set this.Connection = SQLiteCConnection(DbPathName)
    Set Connect = this.Connection
End Function
