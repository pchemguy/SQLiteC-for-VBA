VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ADOlib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Provides ADO related helpers"
'@Folder "SQLite.ADO.ADOlib"
'@ModuleDescription "Provides ADO related helpers"
'@PredeclaredId
'@IgnoreModule IndexedDefaultMemberAccess, ProcedureNotUsed
Option Explicit

Private Type TADOlib
    ConnectionString As String
End Type
Private this As TADOlib


Public Function Create() As ADOlib
    Dim Instance As ADOlib
    Set Instance = New ADOlib
    Set Create = Instance
End Function


Public Sub SetConnectionString(ByVal ConnectionString As String)
    this.ConnectionString = ConnectionString
End Sub


Public Function GetSQLiteConnectionString(Optional ByVal DbPath As String = vbNullString, _
                                          Optional ByVal Features As String = vbNullString) As Scripting.Dictionary
    Dim Driver As String
    Driver = "SQLite3 ODBC Driver"
    
    Dim Database As String
    Database = Right$(DbPath, Abs(Len(DbPath) - 7))
    Database = VerifyOrGetDefaultPath(Database, Array("db", "sqlite"))
    
    Dim Options As String
    Options = IIf(Len(Features) > 0, Features, "SyncPragma=NORMAL;FKSupport=True;")
    
    Dim ConnectionString As Scripting.Dictionary
    Set ConnectionString = New Scripting.Dictionary
    ConnectionString.CompareMode = TextCompare
    
    ConnectionString("ADO") = "Driver=" & Driver & ";" & "Database=" & Database & ";" & Options
    ConnectionString("QT") = "OLEDB;" + ConnectionString("ADO")
    
    Set GetSQLiteConnectionString = ConnectionString
End Function

 
'@Ignore ParameterCanBeByVal
Public Sub GetTableMeta(ByVal TableName As String, ByRef FieldNames As Variant, ByRef FieldTypes As Variant, ByVal FieldMap As Scripting.Dictionary)
    '@Ignore SelfAssignedDeclaration
    Dim Catalog As New ADOX.Catalog
    Catalog.ActiveConnection = this.ConnectionString
    Dim Table As ADOX.Table
    Set Table = Catalog.Tables(TableName)
    Dim FieldCount As Long
    FieldCount = Table.Columns.Count
    
    ReDim FieldNames(1 To FieldCount)
    ReDim FieldTypes(1 To FieldCount)
    
    Dim Column As ADOX.Column
    Dim FieldIndex As Long
    For FieldIndex = 1 To FieldCount
        Set Column = Table.Columns(FieldIndex - 1)
        FieldNames(FieldIndex) = Column.Name
        FieldTypes(FieldIndex) = Column.Type
        FieldMap(Column.Name) = FieldIndex
    Next FieldIndex
End Sub


' Takes an array of field names and types (ADODB.DataTypeEnum) collected by
' GetTableMeta, empties AdoCommand.Parameters and repopulates it.
'
' LIMITATIONS:
' The first field in FieldNames/FieldTypes is assumed to be primary key, which is added last,
' as it is expected to be used in the WHERE clause.
'
'@Description "Prepares ADODB.Parameters for record update."
Public Sub MakeAdoParamsForRecordUpdate(ByVal FieldNames As Variant, _
                                        ByVal FieldTypes As Variant, _
                                        ByVal AdoCommand As ADODB.Command, _
                               Optional ByVal TypeCast As FieldFormat = FieldFormat.CastAsIs)
Attribute MakeAdoParamsForRecordUpdate.VB_Description = "Prepares ADODB.Parameters for record update."
    '''' Make sure that both FieldNames and FieldTypes are arrays and have the same base and size
    Guard.IncompatibleArrays FieldNames, FieldTypes
    Guard.EmptyString FieldNames(LBound(FieldNames))
    Guard.NullReference AdoCommand
    
    Dim AdjustedTypes() As ADODB.DataTypeEnum
    ReDim AdjustedTypes(LBound(FieldTypes) To UBound(FieldTypes))
    Dim FieldIndex As Long
    
    ' Skip primary key types
    For FieldIndex = LBound(FieldNames) + 1 To UBound(FieldNames)
        AdjustedTypes(FieldIndex) = IIf(TypeCast <> CastAllAsText, FieldTypes(FieldIndex), adVarWChar)
    Next FieldIndex
    
    ' Primary key type
    FieldIndex = LBound(FieldNames)
    AdjustedTypes(FieldIndex) = IIf(TypeCast = CastAsIs, FieldTypes(FieldIndex), adVarWChar)
    
    ' Discard any existing parameters
    Dim AdoParams As ADODB.Parameters
    Set AdoParams = AdoCommand.Parameters
    Dim ParamIndex As Long
    For ParamIndex = AdoParams.Count - 1 To 0 Step -1
        AdoParams.Delete ParamIndex
    Next ParamIndex
    
    Const DUMMY_PARAM_SIZE As Long = 1
    Const DUMMY_PARAM_VALUE As Long = 0
    Dim AdoParam As ADODB.Parameter
    For FieldIndex = LBound(FieldNames) + 1 To UBound(FieldNames)
        Set AdoParam = AdoCommand.CreateParameter( _
            FieldNames(FieldIndex), AdjustedTypes(FieldIndex), , DUMMY_PARAM_SIZE, DUMMY_PARAM_VALUE)
        AdoParams.Append AdoParam
    Next FieldIndex
    FieldIndex = LBound(FieldNames)
    Set AdoParam = AdoCommand.CreateParameter( _
        FieldNames(FieldIndex), AdjustedTypes(FieldIndex), , DUMMY_PARAM_SIZE, DUMMY_PARAM_VALUE)
    AdoParams.Append AdoParam
End Sub


Public Sub RecordToAdoParams(ByVal Record As Scripting.Dictionary, ByVal AdoCommand As ADODB.Command)
    Guard.NullReference Record
    Guard.NullReference AdoCommand
    
    Dim AdoParam As Variant
    Dim FieldName As String
    Dim FieldValue As Variant
    For Each AdoParam In AdoCommand.Parameters
        FieldName = AdoParam.Name
        If Record.Exists(FieldName) Then
            FieldValue = Record(FieldName)
            AdoParam.Size = Len(FieldValue)
            AdoParam.Value = FieldValue
        End If
    Next AdoParam
End Sub


'@Description "Outputs Recordset to Excel Worksheet via QueryTable"
Public Sub RecordsetToQT(ByVal AdoRecordset As ADODB.Recordset, ByVal OutputRange As Excel.Range)
Attribute RecordsetToQT.VB_Description = "Outputs Recordset to Excel Worksheet via QueryTable"
    Guard.NullReference AdoRecordset
    Guard.NullReference OutputRange
    
    Dim QTs As Excel.QueryTables
    Set QTs = OutputRange.Worksheet.QueryTables
    
    '''' Cleans up target area before binding the data.
    '''' Provided range reference used to indicate the left column and
    '''' Recordset.Fields.Count determines the width.
    '''' If EntireColumn.Delete method is used, Range object becomes invalid, so
    '''' a textual address must be saved to reset the Range reference.
    '''' However, when multiple QTs are bound to the same worksheet,
    '''' EntireColumn.Delete shifts columns to the left, so the target range
    '''' may not be clear. EntireColumn.Clear clears the contents.
    Dim FieldsCount As Long
    FieldsCount = AdoRecordset.Fields.Count
    Dim QTRangeAddress As String
    QTRangeAddress = OutputRange.Address(External:=True)
    Dim QTRange As Excel.Range
    '@Ignore ImplicitActiveSheetReference
    Set QTRange = Range(QTRangeAddress)
    QTRange.Resize(1, FieldsCount).EntireColumn.Clear
    '@Ignore ImplicitActiveSheetReference
    Set QTRange = Range(QTRangeAddress)
    
    Dim WSQueryTable As Excel.QueryTable
    For Each WSQueryTable In QTs
        WSQueryTable.Delete
    Next WSQueryTable
    
    Dim NamedRange As Excel.Name
    For Each NamedRange In QTRange.Worksheet.Names
        NamedRange.Delete
    Next NamedRange
    
    Set WSQueryTable = QTs.Add(Connection:=AdoRecordset, Destination:=QTRange.Range("A1"))
    With WSQueryTable
        .FieldNames = True
        .RowNumbers = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SaveData = False
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .EnableEditing = True
    End With
    WSQueryTable.Refresh
    QTRange.Worksheet.UsedRange.Rows(1).HorizontalAlignment = xlCenter
End Sub
