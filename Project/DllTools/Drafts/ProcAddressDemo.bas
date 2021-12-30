Attribute VB_Name = "ProcAddressDemo"
'@Folder "DllTools.Drafts"
Option Explicit


#If VBA7 Then
Public Function ProcAddress(ByVal ProcPtr As LongPtr) As LongPtr
#Else
Public Function ProcAddress(ByVal ProcPtr As Long) As Long
#End If
    ProcAddress = ProcPtr
End Function


Private Sub DemoProcAddress()
    Debug.Print ProcAddress(AddressOf ProcAddress)
End Sub
