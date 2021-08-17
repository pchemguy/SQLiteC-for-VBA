VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FieldO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "SQLite field class. The suffix is used to avoid possible collision with ADOX"
'@Folder "SQLiteDB.DB Objects"
'@ModuleDescription "SQLite field class. The suffix is used to avoid possible collision with ADOX"
'@PredeclaredId
'@Exposed
'@IgnoreModule IndexedDefaultMemberAccess
'@IgnoreModule ProcedureNotUsed
Option Explicit

Private Type TFieldO
    FieldName As String
    DataType As String
    NotNull As Boolean
    Default As String
    Check As String
    Unique As Boolean
    Collation As String
    TableName As String
End Type
Private this As TFieldO


'''' @ClassMethodStrict
'''' This method should only be used on the default instance
''''
'@DefaultMember
'@Description "Default factory"
Public Function Create(ByVal FieldName As String, _
              Optional ByVal DataType As String = vbNullString, _
              Optional ByVal NotNull As Boolean = False, _
              Optional ByVal Default As Variant = Empty, _
              Optional ByVal Check As String = vbNullString, _
              Optional ByVal Unique As Boolean = False, _
              Optional ByVal Collation As String = vbNullString, _
              Optional ByVal TableName As String = vbNullString) As FieldO
Attribute Create.VB_Description = "Default factory"
Attribute Create.VB_UserMemId = 0
    Dim Field As FieldO
    Set Field = New FieldO
    Field.Init FieldName, DataType, NotNull, Default, Check, Unique, Collation, TableName
    Set Create = Field
End Function


'@Description "Default constructor"
Public Sub Init(ByVal FieldName As String, _
       Optional ByVal DataType As String = vbNullString, _
       Optional ByVal NotNull As Boolean = False, _
       Optional ByVal Default As Variant = Empty, _
       Optional ByVal Check As String = vbNullString, _
       Optional ByVal Unique As Boolean = False, _
       Optional ByVal Collation As String = vbNullString, _
       Optional ByVal TableName As String = vbNullString)
Attribute Init.VB_Description = "Default constructor"
    this.FieldName = CheckName(FieldName)
    this.DataType = Trim$(DataType)
    this.NotNull = NotNull
    this.Default = CheckDefault(Default)
    this.Check = Check
    this.Unique = Unique
    this.Collation = Collation
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
    PassPattern = "^\w+$"
    Dim re As RegExp
    Set re = New RegExp
    re.Pattern = PassPattern
    Guard.InvalidCharacter re.Test(Name), "SQLlib/FieldO", "Only alpha-numeric characters and the underscore are allowed."
    
    CheckName = Name
End Function


'@Description "Checks 'Default' values and quotes as necessary"
Public Function CheckDefault(ByVal Default As Variant) As Variant
Attribute CheckDefault.VB_Description = "Checks 'Default' values and quotes as necessary"
    If Len(CStr(Default)) = 0 Then
        CheckDefault = vbNullString
        Exit Function
    End If
    
    If Not VarType(Default) = vbString Then
        CheckDefault = "(" & CStr(Default) & ")"
        Exit Function
    End If
    
    Dim DefaultVal As String
    DefaultVal = Trim$(Default)
    If (Left$(DefaultVal, 1) = "(") And (Right$(DefaultVal, 1) = ")") Then
        CheckDefault = DefaultVal
    Else
        Guard.InvalidCharacter Not CBool(InStr(DefaultVal, "'")), "SQLlib/FieldO", "Single quotes are not allowed."
        CheckDefault = "('" & CStr(DefaultVal) & "')"
    End If
End Function


'''' Examples:
''''   >>> ?FieldO("id", "INTEGER", True, 1, "id > 0", True, "nocase").SQL
''''   "    "id" INTEGER NOT NULL DEFAULT 1 CHECK(id > 0) UNIQUE COLLATE NOCASE"
''''
''''   >>> ?FieldO("id", "INTEGER", True, "A", "id > 0", True, "nocase").SQL
''''   "    "id" INTEGER NOT NULL DEFAULT 'A' CHECK(id > 0) UNIQUE COLLATE NOCASE"
''''
''''   >>> ?FieldO("id", "INTEGER", True, "(A)", "id > 0", True, "nocase").SQL
''''   "    "id" INTEGER NOT NULL DEFAULT (A) CHECK(id > 0) UNIQUE COLLATE NOCASE"
''''
'@Description "Returns SQL code to be used in TABLE CREATE"
Public Property Get SQL() As String
Attribute SQL.VB_Description = "Returns SQL code to be used in TABLE CREATE"
    Dim FieldSpec As String
    FieldSpec = IIf(Len(this.DataType) > 0, " " & this.DataType, vbNullString) _
              & IIf(this.NotNull, " NOT NULL", vbNullString) _
              & IIf(Len(this.Default) > 0, " DEFAULT " & this.Default, vbNullString) _
              & IIf(Len(this.Check) > 0, " CHECK(" & this.Check & ")", vbNullString) _
              & IIf(this.Unique, " UNIQUE", vbNullString) _
              & IIf(Len(this.Collation) > 0, " COLLATE " & UCase$(this.Collation), vbNullString)
    SQL = "    " & """" & this.FieldName & """" _
              & IIf(Len(FieldSpec) > 0, "   " & FieldSpec, vbNullString)
End Property


Public Property Let FieldName(ByVal FieldNameVal As String)
    this.FieldName = FieldNameVal
End Property

Public Property Get FieldName() As String
    FieldName = this.FieldName
End Property

Public Property Let DataType(ByVal DataTypeVal As String)
    this.DataType = DataTypeVal
End Property

Public Property Get DataType() As String
    DataType = this.DataType
End Property

Public Property Let NotNull(ByVal NotNullVal As Boolean)
    this.NotNull = NotNullVal
End Property

Public Property Get NotNull() As Boolean
    NotNull = this.NotNull
End Property

Public Property Let Default(ByVal DefaultVal As String)
    this.Default = DefaultVal
End Property

Public Property Get Default() As String
    Default = this.Default
End Property

Public Property Let Check(ByVal CheckVal As String)
    this.Check = CheckVal
End Property

Public Property Get Check() As String
    Check = this.Check
End Property

Public Property Let Unique(ByVal UniqueVal As Boolean)
    this.Unique = UniqueVal
End Property

Public Property Get Unique() As Boolean
    Unique = this.Unique
End Property

Public Property Let Collation(ByVal CollationVal As String)
    this.Collation = CollationVal
End Property

Public Property Get Collation() As String
    Collation = this.Collation
End Property

Public Property Let TableName(ByVal TableNameVal As String)
    this.TableName = TableNameVal
End Property

Public Property Get TableName() As String
    TableName = this.TableName
End Property

Public Property Get Properties() As Scripting.Dictionary
    Dim Props As Scripting.Dictionary
    Set Props = New Scripting.Dictionary
    Props.CompareMode = TextCompare
    Props("FieldName") = this.FieldName
    Props("DataType") = this.DataType
    Props("NotNull") = this.NotNull
    Props("Default") = this.Default
    Props("Check") = this.Check
    Props("Unique") = this.Unique
    Props("Collation") = this.Collation
    Props("TableName") = this.TableName
    Set Properties = Props
End Property
