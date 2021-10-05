Attribute VB_Name = "AddLibDemo"
'@Folder "DllManager.Demo"
Option Explicit
Option Private Module

#If VBA7 Then
Private Declare PtrSafe Function Add Lib "AddLib" (ByVal ValueA As Long, ByVal ValueB As Long) As Long
#Else
Private Declare Function Add Lib "AddLib" (ByVal ValueA As Long, ByVal ValueB As Long) As Long
#End If


Private Type TDllManagerDemo
    DllMan As DllManager
End Type
Private this As TDllManagerDemo


Private Sub GetSum()
    '''' Absolute or relative to ThisWorkbook.Path
    Dim DllPath As String
    DllPath = "Library\SQLiteCforVBA\Demo - DLL - STDCALL and Adapter\AddLib"
    LoadDlls DllPath
    
    '''' Should print -1
    Debug.Print Add(&HFFFFFFFE, 1)
    Set this.DllMan = Nothing
End Sub


Private Sub LoadDlls(ByVal DllPath As String)
    Dim DllMan As DllManager
    Set DllMan = DllManager(DllPath)
    Set this.DllMan = DllMan
    Dim DllNames As Variant
    DllNames = Array( _
        "AddLib.dll" _
    )
    DllMan.LoadMultiple DllNames
End Sub
