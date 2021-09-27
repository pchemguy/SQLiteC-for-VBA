VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLiteCx32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "x32 specific implementation of platform dependent SQLite wrappers."
'@Folder "SQLiteCforVBA"
'@ModuleDescription "VBA6 implementation of platform-dependent SQLite wrappers."
'@PredeclaredId
'@Exposed
Option Explicit


'@DefaultMember
Public Function Create() As SQLiteCx32
Attribute Create.VB_UserMemId = 0
    Dim Instance As SQLiteCx32
    Set Instance = New SQLiteCx32
    Set Create = Instance
End Function
