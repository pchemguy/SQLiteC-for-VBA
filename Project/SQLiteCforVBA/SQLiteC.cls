VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLiteC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Primary public module implementing platform independent SQLite wrappers and encapsulating a platform dependent module"
'@ModuleDescription "Primary public module implementing platform-independent SQLite wrappers and encapsulating a platform dependent module"
'@PredeclaredId
'@Exposed
'@Folder "SQLiteCforVBA"
Option Explicit


Private Type TSQLiteC
    #If VBA7 Then
        APIxx As SQLiteCx64
    #Else
        APIxx As SQLiteCx32
    #End If
End Type
Private this As TSQLiteC


'@DefaultMember
Public Function Create() As SQLiteC
Attribute Create.VB_UserMemId = 0
    Dim Instance As SQLiteC
    Set Instance = New SQLiteC
    Instance.Init
    Set Create = Instance
End Function


Friend Sub Init()
    #If VBA7 Then
        Set this.APIxx = SQLiteCx64()
    #Else
        Set this.APIxx = SQLiteCx32()
    #End If
End Sub
