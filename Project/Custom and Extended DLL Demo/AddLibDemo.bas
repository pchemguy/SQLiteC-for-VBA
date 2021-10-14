Attribute VB_Name = "AddLibDemo"
'@Folder "Custom and Extended DLL Demo"
Option Explicit
Option Private Module

#If VBA7 Then
Private Declare PtrSafe Function Add Lib "AddLib" (ByVal ValueA As Long, ByVal ValueB As Long) As Long
#Else
Private Declare Function Add Lib "AddLib" (ByVal ValueA As Long, ByVal ValueB As Long) As Long
#End If

Private Type TAddLibDemo
    DllMan As DllManager
End Type
Private this As TAddLibDemo


'@EntryPoint
Private Sub GetSum()
    '''' Absolute or relative to ThisWorkbook.Path
    Dim DllPath As String
    #If VBA7 Then
        '''' TODO
        '''' DllPath = "Library\SQLiteCforVBA\Demo - DLL - STDCALL and Adapter\AddLib\x64"
    #Else
        DllPath = "Library\SQLiteCforVBA\Demo - DLL - STDCALL and Adapter\AddLib\x32"
    #End If
    LoadDlls DllPath
    
    '''' Should print -1
    Debug.Print Add(&HFFFFFFFE, 1)
    Set this.DllMan = Nothing
End Sub


Private Sub LoadDlls(ByVal DllPath As String)
    Dim DllMan As DllManager
    '@Ignore IndexedDefaultMemberAccess
    Set DllMan = DllManager.Create(DllPath)
    Set this.DllMan = DllMan
    Dim DllNames As Variant
    DllNames = Array( _
        "AddLib.dll" _
    )
    '@Ignore FunctionReturnValueDiscarded
    DllMan.LoadMultiple DllNames
End Sub

