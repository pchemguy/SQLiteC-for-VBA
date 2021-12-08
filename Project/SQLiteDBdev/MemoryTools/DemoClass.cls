VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DemoClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "SQLiteDBdev.MemoryTools"
'@PredeclaredId
'@IgnoreModule: This is part of the https://github.com/cristianbuse/VBA-MemoryTools
Option Explicit

Private m_id As Long

Public Function Factory(ByVal newID As Long) As DemoClass
    Dim C As New DemoClass
    Init C, newID
    Set Factory = C
End Function

Private Function Init(ByVal C As DemoClass, ByVal newID As Long) As Boolean
    RedirectInstance VarPtr(Init), Me, C
    m_id = newID
End Function

Public Function Factory2(ByVal newID As Long) As DemoClass
    Dim C As New DemoClass
    '
    C.Init2 newID
    Set Factory2 = C
End Function
Public Function Init2(ByVal newID As Long) As DemoClass
    m_id = newID
End Function

Public Property Get ID() As Long
    ID = m_id
End Property
