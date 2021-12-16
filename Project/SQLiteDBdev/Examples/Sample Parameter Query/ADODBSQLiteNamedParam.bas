Attribute VB_Name = "ADODBSQLiteNamedParam"
'@Folder "SQLiteDBdev.Examples.Sample Parameter Query"
'@IgnoreModule IndexedDefaultMemberAccess
Option Explicit


Private Type TAdoParam
    Name As String
    Value As Variant
    Size As Long
    DataType As ADODB.DataTypeEnum
End Type


'@Description "Determines ADODB Parameter Data Type for a VBA variable"
Private Function GetAdoParamType(ByVal TypeValue As Variant) As ADODB.DataTypeEnum
Attribute GetAdoParamType.VB_Description = "Determines ADODB Parameter Data Type for a VBA variable"
    Select Case VarType(TypeValue)
        Case vbString
            GetAdoParamType = adVarWChar
        Case vbInteger, vbLong
            GetAdoParamType = adInteger
        Case vbSingle, vbDouble
            GetAdoParamType = adDouble
        Case Else
            GetAdoParamType = adVarWChar
    End Select
End Function


'@Description "Selects appropriate ADODB Parameter size for a VBA variable value"
Private Function GetAdoParamSize(ByVal Param As Variant) As Long
Attribute GetAdoParamSize.VB_Description = "Selects appropriate ADODB Parameter size for a VBA variable value"
    Select Case VarType(Param)
        Case vbString
            GetAdoParamSize = Len(Param)
        Case vbInteger, vbLong
            GetAdoParamSize = 8
        Case vbSingle, vbDouble
            GetAdoParamSize = 8
        Case Else
            GetAdoParamSize = 255
    End Select
End Function


'@Description "Constructs ADODB connection string"
Private Function GetConnectionString() As String
Attribute GetConnectionString.VB_Description = "Constructs ADODB connection string"
    Dim Database As String
    Database = ThisWorkbook.Path + "\" + "SQLiteDB.db"
    Dim Driver As String
    Driver = "SQLite3 ODBC Driver"
    Dim Options As String
    Options = "SyncPragma=NORMAL;FKSupport=True;"
    Dim AdoConnStr As String
    AdoConnStr = "Driver=" + Driver + ";" + _
                 "Database=" + Database + ";" + _
                 Options
    GetConnectionString = AdoConnStr
End Function


'''' This routine is defined to sketch logical structure and separate concerns.
'''' In practice, ADODB.Command may be instantiated differently. In the SQLiteDB
'''' class, ADODB.Command is a class attribute instantiated by constructor.
''''
'@Description "Stub/demo routine instantiating ADODB.Command"
Private Function GetBaseAdoCommand(ByVal ConnectionString As String) As ADODB.Command
Attribute GetBaseAdoCommand.VB_Description = "Stub/demo routine instantiating ADODB.Command"
    Dim AdoCommand As ADODB.Command
    Set AdoCommand = New ADODB.Command
    
    With AdoCommand
        .CommandType = adCmdText
        .Prepared = True
        .ActiveConnection = ConnectionString
        .ActiveConnection.CursorLocation = adUseClient
    End With
    
    Set GetBaseAdoCommand = AdoCommand
End Function


'''' Emulates named parameter feature not supported by ODBC.
''''
'''' Docs:
''''   https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/namedparameters-property-ado
''''   https://docs.microsoft.com/en-us/sql/odbc/reference/develop-app/binding-parameters-by-name-named-parameters
''''
'''' Routine emulates prepared SQL statements with the named parameters,
'''' by replacing the names with positional index (?#). Suppport of numbered
'''' positional parameters varies. This module targets the SQLite engine, but
'''' may work with other RDBMS as well.
''''
'''' Args:
''''   SQLQuery (string):
''''     SQL query with name parameters (<NamePrefix><ParameterName>)
''''   FieldNames (array):
''''     1D array of field names given in the same order as the order of
''''     ADODB.Parameter objects in the ADODB.Command.Parameters collection.
''''     In practice, the same FieldNames variable should be used for name
''''     substituion and population of ADODB.Command.Parameters.
''''     In SQLite, the first positional parameter is ?1. This routine accepts
''''     0- or 1-based FieldNames array, and ReBase variable is used to adjust
''''     parameter index if necessary.
''''   NamePrefix (string, optional, "?"):
''''     Prefix marking parameter names. Typical prefixes are [@:$]. Here "?"
''''     is used by default.
''''
'''' Returns:
''''   String, containing query with parameter names replaced with positinal
''''   numbered notation.
''''
'''' Examples:
''''   >>> ?NamedToNumberedParams("WHERE [id] <= ?id AND [Age] < ?Age AND [Gender] = ?Gender AND [Email] LIKE ?Email", Array("Age", "id", "Email", "Gender"))
''''   "WHERE [id] <= ?2 AND [Age] < ?1 AND [Gender] = ?4 AND [Email] LIKE ?3"
''''
'@Description "Converts named parameters in an SQL statement to numbered parameters"
Private Function NamedToNumberedParams(ByVal SQLQuery As String, _
                                       ByVal FieldNames As Variant, _
                              Optional ByVal NamePrefix As String = "?") As String
Attribute NamedToNumberedParams.VB_Description = "Converts named parameters in an SQL statement to numbered parameters"
    Debug.Assert IsArray(FieldNames) = True
    Debug.Assert LBound(FieldNames, 1) <= 1
    
    Dim Query As String
    Query = SQLQuery
    Dim ReBase As Long
    ReBase = 1 - LBound(FieldNames, 1)
    Dim FieldIndex As Long
    For FieldIndex = LBound(FieldNames, 1) To UBound(FieldNames, 1)
        Query = Replace(Query, _
                        NamePrefix & FieldNames(FieldIndex), _
                        "?" & CStr(FieldIndex + ReBase))
    Next FieldIndex
    NamedToNumberedParams = Query
End Function


'''' Prepares ADODB.Command.Parameters collection for a parametrized query.
''''
'''' ADODB.Command.Parameters is emptied, if necessary, and populated with
'''' ADODB.Parameter objects using arrays of names and dummy values.
''''
'''' Args:
''''   AdoCommand (ADODB.Command):
''''     Target ADODB.Command.
''''   FieldNames (array):
''''     1D array of field names given in the same order as the order of
''''     dummy values in FieldTypeValues. These names are used to instantiate
''''     named parameters. Later, the actual values should be set by
''''     dereferencing ADODB.Command.Parameters members using their names.
''''   FieldTypeValues (array):
''''     Provides dummy values representing actual parameters. This values
''''     should have the same type as the corresponding parameters are used
''''     to deduce and set the ADODB.Parameter.DataType attribute.
''''
'@Description "Prepares ADODB.Command.Parameters collection"
Private Sub PrepareADODBParameters(ByVal AdoCommand As ADODB.Command, _
                                   ByVal FieldNames As Variant, _
                                   ByVal FieldTypeValues As Variant)
Attribute PrepareADODBParameters.VB_Description = "Prepares ADODB.Command.Parameters collection"
    Debug.Assert IsArray(FieldNames) = True
    Debug.Assert IsArray(FieldTypeValues) = True
    Debug.Assert LBound(FieldNames, 1) = LBound(FieldTypeValues, 1)
    Debug.Assert UBound(FieldNames, 1) = UBound(FieldTypeValues, 1)
    
    ' Discard any existing parameters
    Dim AdoParams As ADODB.Parameters
    Set AdoParams = AdoCommand.Parameters
    Dim ParamIndex As Long
    For ParamIndex = AdoParams.Count - 1 To 0 Step -1
        AdoParams.Delete ParamIndex
    Next ParamIndex
    
    Dim AdoParamProps As TAdoParam
    Dim AdoParam As ADODB.Parameter
    Dim FieldIndex As Long
    For FieldIndex = LBound(FieldNames, 1) To UBound(FieldNames, 1)
        With AdoParamProps
            .Name = FieldNames(FieldIndex)
            .Value = FieldTypeValues(FieldIndex)
            .Size = GetAdoParamSize(.Value)
            .DataType = GetAdoParamType(.Value)
            '@Ignore ArgumentWithIncompatibleObjectType: False positive
            Set AdoParam = AdoCommand.CreateParameter(.Name, .DataType, adParamInput, .Size, .Value)
            AdoCommand.Parameters.Append AdoParam
        End With
    Next FieldIndex
End Sub
    

'''' Sets actual values for AdoCommand.Parameters members.
''''
'''' First, FieldMap dictionary is populated from FieldNames. Then,
'''' AdoCommand.Parameters is enumerated and each member is initialized
'''' using FieldValues and FieldMap to map Parameter's name to the
'''' corresponding FieldValue.
''''
'''' Args:
''''   AdoCommand (ADODB.Command):
''''     Target ADODB.Command.
''''   FieldNames (array):
''''     1D array of field names matching the names of AdoCommand.Parameters
''''     members.
''''   FieldValues (array):
''''     Desired paramter values given in the same order as the order of
''''     FieldNames items. FieldNames maps the members of the
''''     AdoCommand.Parameters to the positions of associated FieldValues items.
''''
'''' Raises
''''   ErrNo.CustomErr:
''''     If AdoCommand.Parameters collection has a member with
''''     name not in FieldNames.
''''
'@Description "Sets actual values for AdoCommand.Parameters members"
Private Sub SetADODBParameters(ByVal AdoCommand As ADODB.Command, _
                               ByVal FieldNames As Variant, _
                               ByVal FieldValues As Variant)
Attribute SetADODBParameters.VB_Description = "Sets actual values for AdoCommand.Parameters members"
    Debug.Assert IsArray(FieldNames) = True
    Debug.Assert IsArray(FieldValues) = True
    Debug.Assert LBound(FieldNames, 1) = LBound(FieldValues, 1)
    Debug.Assert UBound(FieldNames, 1) = UBound(FieldValues, 1)
                               
    Dim FieldMap As Scripting.Dictionary
    Set FieldMap = New Scripting.Dictionary
    FieldMap.CompareMode = TextCompare
    Dim FieldIndex As Variant
    For FieldIndex = LBound(FieldNames, 1) To UBound(FieldNames, 1)
        FieldMap(FieldNames(FieldIndex)) = FieldIndex
    Next FieldIndex
        
    Dim FieldName As String
    Dim FieldValue As Variant
    Dim AdoParam As ADODB.Parameter
    For Each AdoParam In AdoCommand.Parameters
        FieldName = AdoParam.Name
        FieldIndex = FieldMap(FieldName)
        Guard.Expression Not IsEmpty(FieldIndex), "SQLiteDB", "<" & FieldName & "> value not provided."
        FieldValue = FieldValues(FieldMap(FieldName))
        AdoCommand.Parameters(FieldName).Size = GetAdoParamSize(FieldValue)
        AdoCommand.Parameters(FieldName).Value = FieldValue
    Next AdoParam
End Sub


'''' https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/namedparameters-property-ado
'''' https://docs.microsoft.com/en-us/sql/odbc/reference/develop-app/binding-parameters-by-name-named-parameters
''''
'@EntryPoint
'@Description "Emulates named parameters"
Private Sub DemoSQLite3WithNamedParams()
Attribute DemoSQLite3WithNamedParams.VB_Description = "Emulates named parameters"
    Dim AdoCommand As ADODB.Command
    Set AdoCommand = GetBaseAdoCommand(GetConnectionString)
    Dim TableName As String
    TableName = "contacts"
    Dim SQLQuery As String
    SQLQuery = "SELECT * FROM " & TableName & _
              " WHERE [id] <= ?id AND [Age] < ?Age AND [Gender] = ?Gender AND [Email] LIKE ?Email"
              
    Dim FieldNames As Variant
    FieldNames = Array("Age", "id", "Email", "Gender")
    AdoCommand.CommandText = NamedToNumberedParams(SQLQuery, FieldNames)
    
    Dim FieldTypeValues As Variant
    FieldTypeValues = Array(0, 0, " ", " ")
    PrepareADODBParameters AdoCommand, FieldNames, FieldTypeValues

    Dim FieldValues As Variant
    FieldValues = Array(50, 500, "%.net", "male")
    SetADODBParameters AdoCommand, FieldNames, FieldValues
    
    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = New ADODB.Recordset
    With AdoRecordset
        Set .Source = AdoCommand
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockReadOnly
        .Open Options:=adAsyncFetch
        Set .ActiveConnection = Nothing
    End With
    AdoCommand.ActiveConnection.Close

    Dim WSQueryTable As Excel.QueryTable
    For Each WSQueryTable In Buffer.QueryTables
        WSQueryTable.Delete
    Next WSQueryTable
    Buffer.UsedRange.EntireColumn.Delete
    
    Dim NamedRange As Excel.Name
    For Each NamedRange In Buffer.Names
        NamedRange.Delete
    Next NamedRange
        
    Set WSQueryTable = Buffer.QueryTables.Add(Connection:=AdoRecordset, Destination:=Buffer.Range("A1"))
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
    Buffer.UsedRange.Rows(1).HorizontalAlignment = xlCenter
End Sub


'''' Executes the same query as DemoSQLite3WithNamedParams with values
'''' substituted directly into the query.
''''
'@EntryPoint
'@Description "Reference implementation for DemoSQLite3WithNamedParams"
Private Sub DemoSQLite3Ref()
Attribute DemoSQLite3Ref.VB_Description = "Reference implementation for DemoSQLite3WithNamedParams"
    Dim TableName As String
    TableName = "contacts"
    Dim SQLQuery As String
    SQLQuery = "SELECT * FROM " & TableName & _
              " WHERE [id] <= 500 AND [Age] < 50 AND [Gender] = 'male' AND [Email] LIKE '%.net'"

    Dim AdoRecordset As ADODB.Recordset
    Set AdoRecordset = New ADODB.Recordset
    Dim AdoCommand As ADODB.Command
    Set AdoCommand = New ADODB.Command

    With AdoCommand
        .CommandType = adCmdText
        .CommandText = SQLQuery
        .ActiveConnection = GetConnectionString
        .ActiveConnection.CursorLocation = adUseClient
    End With

    With AdoRecordset
        Set .Source = AdoCommand
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockReadOnly
        .Open Options:=adAsyncFetch
        Set .ActiveConnection = Nothing
    End With
    AdoCommand.ActiveConnection.Close

    Dim WSQueryTable As Excel.QueryTable
    For Each WSQueryTable In Buffer.QueryTables
        WSQueryTable.Delete
    Next WSQueryTable
    Buffer.UsedRange.EntireColumn.Delete
    
    Dim NamedRange As Excel.Name
    For Each NamedRange In Buffer.Names
        NamedRange.Delete
    Next NamedRange
        
    Set WSQueryTable = Buffer.QueryTables.Add(Connection:=AdoRecordset, Destination:=Buffer.Range("A1"))
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
    Buffer.UsedRange.Rows(1).HorizontalAlignment = xlCenter
End Sub
