VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConstraintFK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "SQLite Foreign Key Constraint class."
'@Folder "SQLite.SQLiteDBdev.DB Objects.Table Constraint"
'@ModuleDescription "SQLite Foreign Key Constraint class."
'@PredeclaredId
'@Exposed
Option Explicit


Private Type TConstraintFK
    Name As String
    FieldNames As Variant
    ForeignTableName As String
    ForeignFieldNames As Variant
    OnDelete As String
    OnUpdate As String
    TableName As String
End Type
Private this As TConstraintFK


'''' @ClassMethodStrict
'''' This method should only be used on the default instance
''''
'@DefaultMember
'@Description "Default factory"
Public Function Create(ByVal FieldNames As Variant, _
                       ByVal ForeignTableName As Variant, _
                       ByVal ForeignFieldNames As Variant, _
              Optional ByVal OnDelete As String = vbNullString, _
              Optional ByVal OnUpdate As String = vbNullString, _
              Optional ByVal Name As String = vbNullString, _
              Optional ByVal TableName As String = vbNullString) As ConstraintFK
Attribute Create.VB_Description = "Default factory"
Attribute Create.VB_UserMemId = 0
    Dim Instance As ConstraintFK
    Set Instance = New ConstraintFK
    Instance.Init FieldNames, ForeignTableName, ForeignFieldNames, OnDelete, OnUpdate, Name, TableName
    Set Create = Instance
End Function


'@Description "Default constructor"
Public Sub Init(ByVal FieldNames As Variant, _
                ByVal ForeignTableName As Variant, _
                ByVal ForeignFieldNames As Variant, _
       Optional ByVal OnDelete As String = vbNullString, _
       Optional ByVal OnUpdate As String = vbNullString, _
       Optional ByVal Name As String = vbNullString, _
       Optional ByVal TableName As String = vbNullString)
Attribute Init.VB_Description = "Default constructor"
    this.FieldNames = CheckFieldNames(FieldNames)
    this.ForeignTableName = CheckName(ForeignTableName)
    this.ForeignFieldNames = CheckFieldNames(ForeignFieldNames)
    this.OnDelete = CheckFKAction(OnDelete)
    this.OnUpdate = CheckFKAction(OnUpdate)
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
Public Function CheckFKAction(ByVal FKAction As String) As String
Attribute CheckFKAction.VB_Description = "Checks name value"
    Dim AllowedActions As String
    Dim Action As String
    Action = UCase$(FKAction)
    AllowedActions = "__ _CASCADE_ _RESTRICT_ _NO ACTION_ _SET NULL_ _SET DEFAULT_"
    Guard.ActionNotSupported InStr(AllowedActions, "_" & Action & "_"), _
                             "SQLlib/ConstraintFK", _
                             "Not supported foreign key action."
    CheckFKAction = Action
End Function


'@Description "Checks name value"
Public Function CheckName(ByVal Name As String) As String
Attribute CheckName.VB_Description = "Checks name value"
    Guard.EmptyString Name

    Dim PassPattern As String
    PassPattern = "^\w*$"
    Dim re As RegExp
    Set re = New RegExp
    re.Pattern = PassPattern
    Guard.InvalidCharacter re.Test(Name), "SQLlib/ConstraintFK", "Only alpha-numeric characters and the underscore are allowed."
    
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
    Guard.Expression ((VarType(Fields) And vbArray) = vbArray), "SQLlib/ConstraintFK", "Expected array or string"
    
    Dim FieldIndex As Long
    For FieldIndex = LBound(Fields) To UBound(Fields)
        Guard.TypeMismatch VarType(Fields(FieldIndex)) = vbString, "SQLlib/ConstraintFK", "Expected array of strings"
        Fields(FieldIndex) = CheckName(Fields(FieldIndex))
    Next FieldIndex

    CheckFieldNames = Fields
End Function


Public Property Get SQL() As String
    SQL = "    " _
          & IIf(Len(this.Name) > 0, "CONSTRAINT """ & this.Name & """ ", vbNullString) _
          & "FOREIGN KEY(""" & Join(this.FieldNames, """,""") & """) " _
          & "REFERENCES """ & this.ForeignTableName & """" _
          & "(""" & Join(this.ForeignFieldNames, """,""") & """)" _
          & IIf(Len(this.OnDelete) > 0, " ON DELETE " & this.OnDelete, vbNullString) _
          & IIf(Len(this.OnUpdate) > 0, " ON UPDATE " & this.OnUpdate, vbNullString)
End Property
