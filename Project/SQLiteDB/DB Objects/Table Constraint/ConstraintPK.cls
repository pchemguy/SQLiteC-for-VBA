VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConstraintPK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "SQLite Primary Key Constraint class."
'@Folder "SQLiteDB.DB Objects.Table Constraint"
'@ModuleDescription "SQLite Primary Key Constraint class."
'@PredeclaredId
'@Exposed
Option Explicit


Private Type TConstraintPK
    Name As String
    FieldNames As Variant
    AutoIncrement As Boolean
    TableName As String
End Type
Private this As TConstraintPK


'''' @ClassMethodStrict
'''' This method should only be used on the default instance
''''
'@DefaultMember
'@Description "Default factory"
Public Function Create(ByVal FieldNames As Variant, _
              Optional ByVal Name As String = vbNullString, _
              Optional ByVal AutoIncrement As Boolean = False, _
              Optional ByVal TableName As String = vbNullString) As ConstraintPK
Attribute Create.VB_Description = "Default factory"
Attribute Create.VB_UserMemId = 0
    Dim Instance As ConstraintPK
    Set Instance = New ConstraintPK
    Instance.Init FieldNames, Name, AutoIncrement, TableName
    Set Create = Instance
End Function


'@Description "Default constructor"
Public Sub Init(ByVal FieldNames As Variant, _
       Optional ByVal Name As String = vbNullString, _
       Optional ByVal AutoIncrement As Boolean = False, _
       Optional ByVal TableName As String = vbNullString)
Attribute Init.VB_Description = "Default constructor"
    this.FieldNames = CheckFieldNames(FieldNames)
    If Len(Name) > 0 Then
        this.Name = CheckName(Name)
    Else
        this.Name = vbNullString
    End If
    this.AutoIncrement = AutoIncrement
    If Len(TableName) > 0 Then
        this.TableName = CheckName(TableName)
    Else
        this.TableName = vbNullString
    End If
End Sub


'@Ignore ProcedureNotUsed
Public Property Let TableName(ByVal TableNameVal As String)
    this.TableName = TableNameVal
End Property


'@Ignore ProcedureNotUsed
Public Property Get TableName() As String
    TableName = this.TableName
End Property


'@Ignore ProcedureNotUsed
Public Property Get FieldNames() As Variant
    FieldNames = this.FieldNames
End Property


'@Ignore ProcedureNotUsed
Public Property Get AutoIncrement() As Boolean
    AutoIncrement = this.AutoIncrement
End Property


'@Description "Checks name value"
Public Function CheckName(ByVal Name As String) As String
Attribute CheckName.VB_Description = "Checks name value"
    Guard.EmptyString Name

    Dim PassPattern As String
    PassPattern = "^\w*$"
    Dim re As RegExp
    Set re = New RegExp
    re.Pattern = PassPattern
    Guard.InvalidCharacter re.Test(Name), "SQLlib/ConstraintPK", "Only alpha-numeric characters and the underscore are allowed."
    
    CheckName = Name
End Function


Public Function CheckFieldNames(ByVal FieldNames As Variant) As Variant
    Dim Fields As Variant
    Fields = FieldNames
    '''' If a single field name supplied convert to array
    If VarType(Fields) = vbString Then
        Fields = Array(Fields)
    End If
    
    '''' If not (String or Array), raise error
    Guard.Expression ((VarType(Fields) And vbArray) = vbArray), "SQLlib/ConstraintPK", "Expected array or string"
    
    Dim FieldIndex As Long
    For FieldIndex = LBound(Fields) To UBound(Fields)
        Guard.TypeMismatch VarType(Fields(FieldIndex)) = vbString, "SQLlib/ConstraintPK", "Expected array of strings"
        Fields(FieldIndex) = CheckName(Fields(FieldIndex))
    Next FieldIndex

    CheckFieldNames = Fields
End Function


Public Property Get SQL() As String
    SQL = "    " _
          & IIf(Len(this.Name) > 0, "CONSTRAINT """ & this.Name & """ ", vbNullString) _
          & "PRIMARY KEY(""" & Join(this.FieldNames, """,""") & """" _
          & IIf(this.AutoIncrement, " AUTOINCREMENT", vbNullString) & ")"
End Property
