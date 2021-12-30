VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DllManagerDemoSQLiteC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "DllTools.Manager.Demo"
'@PredeclaredId
'@IgnoreModule
''''
'''' Run in the immediate pane:
'''' >>> ?DllManagerDemoSQLiteC("").Version
''''
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function sqlite3_libversion_number Lib "SQLite3" () As Long
#Else
Private Declare Function sqlite3_libversion_number Lib "SQLite3" () As Long
#End If

Private Type TObjectState
    DllMan As DllManager
    Connections As Scripting.Dictionary
End Type
Private this As TObjectState


'@DefaultMember
Public Function Create(ByVal DllPath As String, _
              Optional ByVal DllNames As Variant = Empty) As DllManagerDemoSQLiteC
Attribute Create.VB_UserMemId = 0
    Dim Instance As DllManagerDemoSQLiteC
    Set Instance = New DllManagerDemoSQLiteC
    Instance.Init DllPath, DllNames
    Set Create = Instance
End Function

Friend Sub Init(ByVal DllPath As String, _
       Optional ByVal DllNames As Variant = Empty)
    Dim FileNames As Variant
    If Not IsEmpty(DllNames) Then
        FileNames = DllNames
    Else
        #If Win64 Then
            FileNames = "sqlite3.dll"
        #Else
            FileNames = Array("icudt68.dll", "icuuc68.dll", "icuin68.dll", _
                              "icuio68.dll", "icutu68.dll", "sqlite3.dll")
        #End If
    End If
    Set this.DllMan = DllManager.Create(DllPath, FileNames)
    Set this.Connections = New Scripting.Dictionary
    this.Connections.CompareMode = TextCompare
End Sub

Public Function Version() As Long
    Version = sqlite3_libversion_number()
End Function
