VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FixMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Meta package encapsulating individual components"
'@Folder "SQLite.Fixtures"
'@ModuleDescription "Meta package encapsulating individual components"
'@PredeclaredId
Option Explicit

Private Type TFixMain
    SQL As FixSQLMain
    ObjC As FixObjC
    Utils As FixUtils
End Type
Private this As TFixMain


Private Sub Class_Initialize()
    With this
        Set .SQL = FixSQLMain
        Set .ObjC = FixObjC
        Set .Utils = FixUtils
    End With
End Sub

Public Property Get SQL() As FixSQLMain
    Set SQL = this.SQL
End Property

Public Property Get ObjC() As FixObjC
    Set ObjC = this.ObjC
End Property

Public Property Get Utils() As FixUtils
    Set Utils = this.Utils
End Property


