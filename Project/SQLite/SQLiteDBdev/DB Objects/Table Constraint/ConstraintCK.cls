VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConstraintCK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "SQLite Check Constraint class."
'@Folder "SQLite.SQLiteDBdev.DB Objects.Table Constraint"
'@ModuleDescription "SQLite Check Constraint class."
'@PredeclaredId
'@Exposed
'@IgnoreModule ProcedureNotUsed
Option Explicit


Private Type TConstraintCK
    Name As String
    Expression As String
    TableName As String
End Type
Private this As TConstraintCK


'''' @ClassMethodStrict
'''' This method should only be used on the default instance
''''
'@DefaultMember
'@Description "Default factory"
Public Function Create(ByVal Expression As String, _
              Optional ByVal Name As String, _
              Optional ByVal TableName As String = vbNullString) As ConstraintCK
Attribute Create.VB_Description = "Default factory"
Attribute Create.VB_UserMemId = 0
    Dim Instance As ConstraintCK
    Set Instance = New ConstraintCK
    Instance.Init Expression, Name, TableName
    Set Create = Instance
End Function


'@Description "Default constructor"
Public Sub Init(ByVal Expression As String, _
       Optional ByVal Name As String, _
       Optional ByVal TableName As String = vbNullString)
Attribute Init.VB_Description = "Default constructor"
    Guard.EmptyString Expression
    this.Expression = Expression
    If Len(Name) > 0 Then
        this.Name = CheckName(Name)
    Else
        this.Name = vbNullString
    End If
    If Len(TableName) > 0 Then
        this.TableName = CheckName(TableName)
    Else
        this.TableName = vbNullString
    End If
End Sub


'@Description "Checks name value"
Public Function CheckName(ByVal Name As String) As String
Attribute CheckName.VB_Description = "Checks name value"
    Guard.EmptyString Name

    Dim PassPattern As String
    PassPattern = "^\w*$"
    Dim re As RegExp
    Set re = New RegExp
    re.Pattern = PassPattern
    Guard.InvalidCharacter re.Test(Name), "SQLlib/ConstraintCK", "Only alpha-numeric characters and the underscore are allowed."
    
    CheckName = Name
End Function


Public Property Get SQL() As String
    SQL = "    " _
          & IIf(Len(this.Name) > 0, "CONSTRAINT """ & this.Name & """ ", vbNullString) _
          & "CHECK(" & this.Expression & ")"
End Property


Public Property Let Name(ByVal NameVal As String)
    this.Name = NameVal
End Property

Public Property Get Name() As String
    Name = this.Name
End Property

Public Property Let Expression(ByVal ExpressionVal As String)
    this.Expression = ExpressionVal
End Property

Public Property Get Expression() As String
    Expression = this.Expression
End Property

Public Property Let TableName(ByVal TableNameVal As String)
    this.TableName = TableNameVal
End Property

Public Property Get TableName() As String
    TableName = this.TableName
End Property
