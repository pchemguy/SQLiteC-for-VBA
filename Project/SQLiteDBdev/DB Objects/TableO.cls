VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "SQLite table class. The suffix is used to avoid possible collision with ADOX"
'@Folder "SQLiteDBdev.DB Objects"
'@ModuleDescription "SQLite table class. The suffix is used to avoid possible collision with ADOX"
'@PredeclaredId
'@Exposed
'@IgnoreModule IndexedDefaultMemberAccess, ProcedureNotUsed
Option Explicit


Private Type TTableO
    TableName As String
    FieldNames As Variant
    Fields As Scripting.Dictionary
    UQs As Scripting.Dictionary
    CKs As Scripting.Dictionary
    FKs As Scripting.Dictionary
    PK As ConstraintPK
    DbManager As ILiteADO
    FromDbHelper As TableOFromDbHelper
    '''' Interpreted values:
    ''''   vbNullString - No Action
    ''''   DROP - Add drop statement
    ''''   SKIP - Add "IF NOT EXISTS" clause
    CreateActionIfExists As String
End Type
Private this As TTableO


'''' @ClassMethodStrict
'''' This method should only be used on the default instance
''''
'''' Args:
''''   TableName (string):
''''     Table name
''''   FieldNames (array):
''''     1D 1-based array of string field names. Establishes the order of the fields.
''''   Fields (dictionary):
''''     key - field name
''''     value - FieldO
''''   FKs (dictionary):
''''     key - FOREIGN KEY constraint name (for anon constraint, use constraint SQL prefixed with '____')
''''     value - ConstraintFK
''''   UQs (dictionary):
''''     key - UNIQUE constraint name (for anon constraint, use constraint SQL prefixed with '____')
''''     value - ConstraintUQ:
''''   CKs (dictionary):
''''     key - CHECK constraint name (for anon constraint, use constraint SQL prefixed with '____')
''''     value - ConstraintCK:
''''   PKs (ConstraintPK):
''''
'@DefaultMember
'@Description "Default factory"
Public Function Create(ByVal TableName As String, _
                       ByVal FieldNames As Variant, _
                       ByVal Fields As Scripting.Dictionary, _
              Optional ByVal UQs As Scripting.Dictionary = Nothing, _
              Optional ByVal CKs As Scripting.Dictionary = Nothing, _
              Optional ByVal FKs As Scripting.Dictionary = Nothing, _
              Optional ByVal PK As ConstraintPK = Nothing) As TableO
Attribute Create.VB_Description = "Default factory"
Attribute Create.VB_UserMemId = 0
    Dim Instance As TableO
    Set Instance = New TableO
    Instance.Init TableName, FieldNames, Fields, UQs, CKs, FKs, PK
    Set Create = Instance
End Function


'@Description "Factory, pulling metadata from database"
Public Function FromDb(ByVal TableName As String, _
                       ByVal DbManager As ILiteADO) As TableO
Attribute FromDb.VB_Description = "Factory, pulling metadata from database"
    Dim Instance As TableO
    Set Instance = New TableO
    
    Set Instance.FromDbHelper = TableOFromDbHelper(TableName, DbManager)
    
    Dim DbFields() As FieldO
    DbFields = Instance.FromDbHelper.GetDbFields()
    
    Dim Fields As Scripting.Dictionary
    Set Fields = New Scripting.Dictionary
    Fields.CompareMode = TextCompare
    
    Dim FieldNames() As String
    ReDim FieldNames(1 To UBound(DbFields))
    Dim FieldIndex As Long
    For FieldIndex = 1 To UBound(DbFields)
        FieldNames(FieldIndex) = DbFields(FieldIndex).FieldName
        Set Fields(FieldNames(FieldIndex)) = DbFields(FieldIndex)
    Next FieldIndex
    
    Dim DbCons As Scripting.Dictionary
    Set DbCons = Instance.FromDbHelper.GetConstraints()
        
    Instance.Init TableName, FieldNames, Fields, _
                  DbCons("UQs"), DbCons("CKs"), DbCons("FKs"), DbCons("PK")
    Set Instance.DbManager = DbManager
    
    Set FromDb = Instance
End Function


'@Description "Default constructor"
Friend Sub Init(ByVal TableName As String, _
                ByVal FieldNames As Variant, _
                ByVal Fields As Scripting.Dictionary, _
       Optional ByVal UQs As Scripting.Dictionary = Nothing, _
       Optional ByVal CKs As Scripting.Dictionary = Nothing, _
       Optional ByVal FKs As Scripting.Dictionary = Nothing, _
       Optional ByVal PK As ConstraintPK = Nothing)
Attribute Init.VB_Description = "Default constructor"
    Dim Result As Boolean
    Dim FieldIndex As Long
    Dim FieldName As Variant
    
    this.TableName = CheckName(TableName)
    
    Result = (VarType(FieldNames) = (vbArray Or vbString))
    Guard.TypeMismatch Result, "TableO/Constructor", "FieldNames type mismatch"
    this.FieldNames = FieldNames
    
    Result = (UBound(this.FieldNames) = Fields.Count)
    Guard.Expression Result, "TableO/Constructor", "FieldNames/Fields count mismatch"
    
    Result = True
    For FieldIndex = LBound(FieldNames) To UBound(FieldNames)
        Result = Result And Fields.Exists(FieldNames(FieldIndex))
    Next FieldIndex
    Guard.Expression Result, "TableO/Constructor", "FieldNames/Fields mismatch"
    Set this.Fields = Fields
    
    Dim ConName As Variant
    
    If Not UQs Is Nothing Then
        '''' Verifies that all field names specified in
        '''' UQs are available in Fields
        Result = True
        Dim ConUQ As ConstraintUQ
        For Each ConName In UQs.Keys
            Set ConUQ = UQs(ConName)
            For Each FieldName In ConUQ.FieldNames
                Result = Result And Fields.Exists(FieldName)
            Next FieldName
        Next ConName
        Guard.Expression Result, "TableO/Constructor", "Invalid fields in UQs"
    End If
    Set this.UQs = UQs
    
    If Not FKs Is Nothing Then
        '''' Verifies that all source field names specified in
        '''' FKs are available in Fields
        Result = True
        Dim ConFK As ConstraintFK
        For Each ConName In FKs.Keys
            Set ConFK = FKs(ConName)
            For Each FieldName In ConFK.FieldNames
                Result = Result And Fields.Exists(FieldName)
            Next FieldName
        Next ConName
        Guard.Expression Result, "TableO/Constructor", "Invalid fields in FKs"
    End If
    Set this.FKs = FKs

    If Not PK Is Nothing Then
        '''' Verifies that all field names specified in
        '''' the PK are available in Fields
        Result = True
        For Each FieldName In PK.FieldNames
            Result = Result And Fields.Exists(FieldName)
        Next FieldName
        Guard.Expression Result, "TableO/Constructor", "Invalid fields in PK"
    
        If PK.AutoIncrement = True Then
            '''' For AutoIncrement, only single field is allowed
            Result = Not (UBound(PK.FieldNames) > 0)
            Guard.Expression Result, "TableO/Constructor", "PK /w AI should have one field"
            
            '''' For AutoIncrement, only integer field types are allowed
            Dim Field As FieldO
            Set Field = Fields(PK.FieldNames(0))
            Result = (Left$(UCase$(Field.Properties("DataType")), 3) = "INT")
            Guard.Expression Result, "TableO/Constructor", "PK /w AI should be and INT"
        End If
    End If
    Set this.PK = PK

    Set this.CKs = CKs
End Sub


'@Description "Checks name value"
Friend Function CheckName(ByVal Name As String) As String
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


Public Property Get SQL() As String
    Dim LineCount As Long
    LineCount = this.Fields.Count _
              + this.CKs.Count _
              + this.UQs.Count _
              + this.FKs.Count _
              + IIf(Not this.PK Is Nothing, 1, 0)
    Dim CreateTable() As String
    ReDim CreateTable(0 To LineCount - 1)
    
    Dim Keys As Variant
    Dim LineIndex As Long
    LineIndex = 0
    Dim KeyIndex As Long
    
    Keys = this.FieldNames
    For KeyIndex = LBound(Keys) To UBound(Keys)
        CreateTable(LineIndex) = this.Fields(Keys(KeyIndex)).SQL
        LineIndex = LineIndex + 1
    Next KeyIndex
    
    Dim Cons As Scripting.Dictionary
    
    Set Cons = this.CKs
    If Cons.Count > 0 Then
        Keys = Cons.Keys
        For KeyIndex = LBound(Keys) To UBound(Keys)
            CreateTable(LineIndex) = Cons(Keys(KeyIndex)).SQL
            LineIndex = LineIndex + 1
        Next KeyIndex
    End If

    Set Cons = this.UQs
    If Cons.Count > 0 Then
        Keys = Cons.Keys
        For KeyIndex = LBound(Keys) To UBound(Keys)
            CreateTable(LineIndex) = Cons(Keys(KeyIndex)).SQL
            LineIndex = LineIndex + 1
        Next KeyIndex
    End If
    
    Set Cons = this.FKs
    If Cons.Count > 0 Then
        Keys = Cons.Keys
        For KeyIndex = LBound(Keys) To UBound(Keys)
            CreateTable(LineIndex) = Cons(Keys(KeyIndex)).SQL
            LineIndex = LineIndex + 1
        Next KeyIndex
    End If
    
    If Not this.PK Is Nothing Then
        CreateTable(LineIndex) = this.PK.SQL
    End If
    
    Dim DropCmd As String
    If this.CreateActionIfExists = "DROP" Then
        DropCmd = "DROP TABLE IF EXISTS """ & this.TableName & """;" & vbLf
    Else
        DropCmd = vbNullString
    End If
    SQL = DropCmd & "CREATE TABLE " & _
          IIf(this.CreateActionIfExists = "SKIP", "IF NOT EXISTS ", vbNullString) & _
          """" & this.TableName & """ (" & vbLf & _
          Join(CreateTable, "," & vbLf) & vbLf & ")"
End Property


Public Sub CreateTable(Optional ByVal OnConflictAction As Variant = Empty)
    If Not IsEmpty(OnConflictAction) Then
        this.CreateActionIfExists = OnConflictAction
    End If
    '@Ignore FunctionReturnValueDiscarded
    this.DbManager.ExecuteNonQuery Me.SQL
End Sub


'=============================================================================='
'---------------------------------- ACCESSORS ---------------------------------'
'=============================================================================='

Public Property Get TableName() As String
    TableName = this.TableName
End Property

Public Property Let TableName(ByVal TableNameVal As String)
    this.TableName = TableNameVal
End Property

Public Property Get CreateActionIfExists() As String
    CreateActionIfExists = this.CreateActionIfExists
End Property

Public Property Let CreateActionIfExists(ByVal CreateActionIfExistsVal As String)
    this.CreateActionIfExists = UCase$(CreateActionIfExistsVal)
End Property

Public Property Get FieldNames() As Variant
    FieldNames = this.FieldNames
End Property

Public Property Let FieldNames(ByVal FieldNamesVal As Variant)
    this.FieldNames = FieldNamesVal
End Property

Public Property Get Fields() As Scripting.Dictionary
    Set Fields = this.Fields
End Property

Public Property Set Fields(ByVal FieldsInstance As Scripting.Dictionary)
    Set this.Fields = FieldsInstance
End Property

Public Property Get UQs() As Scripting.Dictionary
    Set UQs = this.UQs
End Property

Public Property Set UQs(ByVal UQsInstance As Scripting.Dictionary)
    Set this.UQs = UQsInstance
End Property

Public Property Get CKs() As Scripting.Dictionary
    Set CKs = this.CKs
End Property

Public Property Set CKs(ByVal CKsInstance As Scripting.Dictionary)
    Set this.CKs = CKsInstance
End Property

Public Property Get FKs() As Scripting.Dictionary
    Set FKs = this.FKs
End Property

Public Property Set FKs(ByVal FKsInstance As Scripting.Dictionary)
    Set this.FKs = FKsInstance
End Property

Public Property Get PK() As ConstraintPK
    Set PK = this.PK
End Property

Public Property Set PK(ByVal PKInstance As ConstraintPK)
    Set this.PK = PKInstance
End Property

Public Property Get DbManager() As ILiteADO
    Set DbManager = this.DbManager
End Property

Public Property Set DbManager(ByVal DbManagerInstance As ILiteADO)
    Set this.DbManager = DbManagerInstance
End Property

Public Property Get FromDbHelper() As TableOFromDbHelper
    Set FromDbHelper = this.FromDbHelper
End Property

Public Property Set FromDbHelper(ByVal FromDbHelperInstance As TableOFromDbHelper)
    Set this.FromDbHelper = FromDbHelperInstance
End Property
