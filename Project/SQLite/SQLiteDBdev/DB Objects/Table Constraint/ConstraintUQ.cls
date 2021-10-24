VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConstraintUQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "SQLite Unique Constraint class."
'@Folder "SQLite.SQLiteDBdev.DB Objects.Table Constraint"
'@ModuleDescription "SQLite Unique Constraint class."
'@PredeclaredId
'@Exposed
Option Explicit


Private Type TConstraintUQ
    Name As String
    FieldNames As Variant
    TableName As String
End Type
Private this As TConstraintUQ


'''' @ClassMethodStrict
'''' This method should only be used on the default instance
''''
'@DefaultMember
'@Description "Default factory"
Public Function Create(ByVal FieldNames As Variant, _
              Optional ByVal Name As String, _
              Optional ByVal TableName As String = vbNullString) As ConstraintUQ
Attribute Create.VB_Description = "Default factory"
Attribute Create.VB_UserMemId = 0
    Dim Instance As ConstraintUQ
    Set Instance = New ConstraintUQ
    Instance.Init FieldNames, Name, TableName
    Set Create = Instance
End Function


'@Description "Default constructor"
Public Sub Init(ByVal FieldNames As Variant, _
       Optional ByVal Name As String, _
       Optional ByVal TableName As String = vbNullString)
Attribute Init.VB_Description = "Default constructor"
    this.FieldNames = CheckFieldNames(FieldNames)
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


'@Description "Checks name value"
Public Function CheckName(ByVal Name As String) As String
Attribute CheckName.VB_Description = "Checks name value"
    Guard.EmptyString Name

    Dim PassPattern As String
    PassPattern = "^\w*$"
    Dim re As RegExp
    Set re = New RegExp
    re.Pattern = PassPattern
    Guard.InvalidCharacter re.Test(Name), "SQLlib/ConstraintUQ", "Only alpha-numeric characters and the underscore are allowed."
    
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
    Guard.Expression ((VarType(Fields) And vbArray) = vbArray), "SQLlib/ConstraintUQ", "Expected array or string"
    
    Dim FieldIndex As Long
    For FieldIndex = LBound(Fields) To UBound(Fields)
        Guard.TypeMismatch VarType(Fields(FieldIndex)) = vbString, "SQLlib/ConstraintUQ", "Expected array of strings"
        Fields(FieldIndex) = CheckName(Fields(FieldIndex))
    Next FieldIndex

    CheckFieldNames = Fields
End Function


Public Property Get SQL() As String
    SQL = "    " _
          & IIf(Len(this.Name) > 0, "CONSTRAINT """ & this.Name & """ ", vbNullString) _
          & "UNIQUE(""" & Join(this.FieldNames, """,""") & """)"
End Property
