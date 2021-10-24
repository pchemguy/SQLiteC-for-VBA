VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FixSQLMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Meta package encapsulating individual SQL modules."
'@Folder "SQLite.Fixtures"
'@ModuleDescription "Meta package encapsulating individual SQL modules."
'@PredeclaredId
Option Explicit

Private Type TFixSQLMain
    ITRB As FixSQLITRB
    Misc As FixSQLGeneral
    Func As FixSQLFunctions
End Type
Private this As TFixSQLMain


Private Sub Class_Initialize()
    With this
        Set .ITRB = FixSQLITRB
        Set .Misc = FixSQLGeneral
        Set .Func = FixSQLFunctions
    End With
End Sub

Public Property Get ITRB() As FixSQLITRB
    Set ITRB = this.ITRB
End Property

Public Property Get Misc() As FixSQLGeneral
    Set Misc = this.Misc
End Property

Public Property Get Func() As FixSQLFunctions
    Set Func = this.Func
End Property

