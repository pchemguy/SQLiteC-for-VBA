VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLiteCx64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder "SQLiteCforVBA"
'@ModuleDescription "VBA7 implementation of platform-dependent SQLite wrappers."
'@PredeclaredId
'@Exposed
Option Explicit


'@DefaultMember
Public Function Create() As SQLiteCx64
Attribute Create.VB_UserMemId = 0
    Dim Instance As SQLiteCx64
    Set Instance = New SQLiteCx64
    Set Create = Instance
End Function
