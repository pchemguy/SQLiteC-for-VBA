VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IndexO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "SQLite Index class."
'@Folder "SQLiteDBdev.DB Objects"
'@ModuleDescription "SQLite Index class."
'@PredeclaredId
'@Exposed
Option Explicit


Private Type TIndexO
    Name As String
    TableName As String
    FieldNames As Variant
    Unique As Boolean
    Where As String
End Type
Private this As TIndexO


'''' @ClassMethodStrict
'''' This method should only be used on the default instance
''''
'''' Args:
''''   FieldNames:
''''     2D 0-based Variant array. Each row is a triplet
''''     (FieldName/Expression, SortOrder, Collation). If
''''     FieldName/Expression only contains \w+, it is assumed
''''     to be a column name.
''''     Array of Array format is also accepted.
''''
'@DefaultMember
'@Description "Default factory"
Public Function Create(ByVal Name As String, _
                       ByVal TableName As String, _
                       ByVal FieldNames As Variant, _
              Optional ByVal Unique As Boolean = False, _
              Optional ByVal Where As String = vbNullString) As IndexO
Attribute Create.VB_Description = "Default factory"
Attribute Create.VB_UserMemId = 0
    Dim Instance As IndexO
    Set Instance = New IndexO
    Instance.Init Name, TableName, FieldNames, Unique, Where
    Set Create = Instance
End Function


'@Description "Default constructor"
Public Sub Init(ByVal Name As String, _
                ByVal TableName As String, _
                ByVal FieldNames As Variant, _
       Optional ByVal Unique As Boolean = False, _
       Optional ByVal Where As String = vbNullString)
Attribute Init.VB_Description = "Default constructor"
    this.Name = CheckName(Name)
    this.TableName = CheckName(TableName)
    this.FieldNames = CheckFieldNames(FieldNames)
    this.Unique = Unique
    this.Where = Where
End Sub


'@Ignore ProcedureNotUsed
Public Property Let TableName(ByVal TableNameVal As String)
    this.TableName = TableNameVal
End Property


'@Ignore ProcedureNotUsed
Public Property Get TableName() As String
    TableName = this.TableName
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
    Guard.InvalidCharacter re.Test(Name), "SQLlib/IndexO", "Only alpha-numeric characters and the underscore are allowed."
    
    CheckName = Name
End Function


Public Function CheckFieldNames(ByVal FieldNames As Variant) As Variant
    Dim NDim As Long
    NDim = ArrayLib.NumberOfArrayDimensions(FieldNames)
    Dim Fields As Variant: Fields = Empty
    Dim Field(0 To 0, 0 To 2) As String
    Select Case NDim
        '''' If FieldNames is a String, it represents
        '''' either a single field or an expression
        Case 0
            If VarType(FieldNames) = vbString Then
                Field(0, 0) = FieldNames
                CheckFieldNames = Field
            Else
                Guard.Expression False, "SQLlib/IndexO", "Wrong shape of FieldNames"
            End If
        Case 1
            If (VarType(FieldNames(0)) And vbArray) <> vbArray Then
                '''' Assume 1D array in 1D array
                If LBound(FieldNames, 1) = 0 And UBound(FieldNames, 1) = 2 Then
                    Field(0, 0) = FieldNames(0)
                    Field(0, 1) = FieldNames(1)
                    Field(0, 2) = FieldNames(2)
                    CheckFieldNames = Field
                Else
                    Guard.Expression False, "SQLlib/IndexO", "Wrong shape of FieldNames"
                End If
            ElseIf ArrayLib.NumberOfArrayDimensions(FieldNames(0)) = 1 _
                   And LBound(FieldNames(0), 1) = 0 _
                   And UBound(FieldNames(0), 1) <= 2 Then
                ReDim Fields(0 To UBound(FieldNames, 1), 0 To 2)
                Dim FieldIndex As Long
                '''' Convert 1D of 1D to 2D
                For FieldIndex = 0 To UBound(Fields, 1)
                    Fields(FieldIndex, 0) = FieldNames(FieldIndex)(0)
                    If UBound(FieldNames(0), 1) >= 1 Then
                        Fields(FieldIndex, 1) = FieldNames(FieldIndex)(1)
                    Else
                        Fields(FieldIndex, 1) = Empty
                    End If
                    If UBound(FieldNames(0), 1) = 2 Then
                        Fields(FieldIndex, 2) = FieldNames(FieldIndex)(2)
                    Else
                        Fields(FieldIndex, 2) = Empty
                    End If
                Next FieldIndex
                CheckFieldNames = Fields
            Else
                Guard.Expression False, "SQLlib/IndexO", "Wrong shape of FieldNames"
            End If
        Case 2
            If LBound(FieldNames, 2) = 0 And UBound(FieldNames, 2) = 2 Then
                CheckFieldNames = FieldNames
            Else
                Guard.Expression False, "SQLlib/IndexO", "Wrong shape of FieldNames"
            End If
        Case Else
            Guard.Expression False, "SQLlib/IndexO", "Wrong shape of FieldNames"
    End Select
End Function


Public Property Get SQL() As String
    Dim Fields() As String
    ReDim Fields(0 To UBound(this.FieldNames, 1))
    
    Dim PassPattern As String
    PassPattern = "^\w*$"
    Dim re As RegExp
    Set re = New RegExp
    re.Pattern = PassPattern
    
    Dim FieldIndex As Long
    Dim Expr As String
    Dim Collate As String
    Dim Order As String
    For FieldIndex = 0 To UBound(Fields)
        Expr = this.FieldNames(FieldIndex, 0)
        Expr = IIf(re.Test(Expr), """" & Expr & """", Expr)
        Collate = this.FieldNames(FieldIndex, 2)
        Collate = IIf(Len(Collate) > 0, " COLLATE " & Collate, vbNullString)
        Order = UCase$(this.FieldNames(FieldIndex, 1))
        Order = IIf(Order = "ASC" Or Order = "DESC", " " & Order, vbNullString)
        Fields(FieldIndex) = Expr & Collate & Order
    Next FieldIndex
    SQL = "CREATE " _
        & IIf(this.Unique, "UNIQUE ", vbNullString) _
        & "INDEX """ & this.Name & """ " _
        & "ON """ & this.TableName & """(" _
        & Join(Fields, ", ") & ")"
End Property
