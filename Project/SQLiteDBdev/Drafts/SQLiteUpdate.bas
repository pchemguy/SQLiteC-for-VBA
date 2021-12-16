Attribute VB_Name = "SQLiteUpdate"
'@Folder "SQLiteDBdev.Drafts"
'@IgnoreModule VariableNotUsed, ProcedureNotUsed, IndexedDefaultMemberAccess, SelfAssignedDeclaration
Option Explicit


Private Type TAdoParam
    Name As String
    Value As Variant
    Size As Long
    DataType As ADODB.DataTypeEnum
End Type


Private Function GetConnectionString() As Scripting.Dictionary
    Dim Driver As String: Driver = "SQLite3 ODBC Driver"

    Dim Database As String
    Database = ThisWorkbook.Path & Application.PathSeparator & "ADODBTemplates.db"

    Dim Options As String
    Options = "SyncPragma=NORMAL;FKSupport=True;"
    
    Dim ConnectionString As Scripting.Dictionary
    Set ConnectionString = New Scripting.Dictionary
    ConnectionString.CompareMode = TextCompare
    
    ConnectionString("ADO") = "Driver=" & Driver & ";" & "Database=" & Database & ";" & Options
    ConnectionString("OLEDB") = "OLEDB;" + ConnectionString("ADO")

    Set GetConnectionString = ConnectionString
End Function


Private Function GetAdoParamType(ByVal TypeValue As Variant) As ADODB.DataTypeEnum
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


Private Function GetAdoParamSize(ByVal Param As Variant) As Long
    Select Case VarType(Param)
        Case vbString
            GetAdoParamSize = Len(Param)
        Case vbInteger, vbLong
            GetAdoParamSize = 0
        Case vbSingle, vbDouble
            GetAdoParamSize = 0
        Case Else
            GetAdoParamSize = 0
    End Select
End Function


Private Function GetSQLUpdate(ByVal TableName As String, ByVal FieldNames As Variant) As Scripting.Dictionary
    Dim SQL As New Scripting.Dictionary: SQL.CompareMode = TextCompare
    Dim fso As New Scripting.FileSystemObject
    
    SQL("TableName") = TableName
    Dim FieldCount As Long
    Dim PKName As String: PKName = FieldNames(LBound(FieldNames, 1))
    Dim SetFields As String
    SetFields = Join(FieldNames, ", ")
    SetFields = Right$(SetFields, Len(SetFields) - Len(PKName) - 2)
    FieldCount = UBound(FieldNames, 1) - LBound(FieldNames, 1) + 1
    
    SQL("UpdateQuery") = "UPDATE " & _
                         TableName & _
                         " SET " & _
                         "(" & SetFields & ")" & _
                         " = " & _
                         "(" & Replace(String(FieldCount - 2, "@"), "@", "?, ") & "?" & ")" & _
                         " WHERE " & _
                         PKName & _
                         " = ?"
    Set GetSQLUpdate = SQL
End Function


Private Function GetAdoCommand(ByVal TableName As String, _
                               ByVal FieldNames As Variant, _
                               ByVal FieldTypeValues As Variant) As ADODB.Command
    Debug.Assert LBound(FieldNames, 1) = LBound(FieldTypeValues, 1)
    Debug.Assert UBound(FieldNames, 1) = UBound(FieldTypeValues, 1)
    
    Dim ConnectionString As Scripting.Dictionary
    Set ConnectionString = GetConnectionString

    Dim SQL As New Scripting.Dictionary: SQL.CompareMode = TextCompare
    Set SQL = GetSQLUpdate(TableName, FieldNames)

    Dim AdoCommand As ADODB.Command
    Set AdoCommand = New ADODB.Command
    With AdoCommand
        .CommandType = adCmdText
        .CommandText = SQL("UpdateQuery")
        .ActiveConnection = ConnectionString("ADO")
        .Prepared = True
    End With
    
    Dim AdoParamProps As TAdoParam
    Dim AdoParam As ADODB.Parameter
    Dim FieldIndex As Long
    
    For FieldIndex = LBound(FieldNames, 1) + 1 To UBound(FieldNames, 1)
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
    
    Dim PKName As String: PKName = FieldNames(LBound(FieldNames, 1))
    Dim PKType As ADODB.DataTypeEnum: PKType = GetAdoParamType(FieldTypeValues(LBound(FieldTypeValues, 1)))
    Set AdoParam = AdoCommand.CreateParameter(PKName, PKType, adParamInput)
    AdoCommand.Parameters.Append AdoParam
    
    Set GetAdoCommand = AdoCommand
End Function


Private Sub PrepareAdoCommand()
    Dim TableName As String: TableName = "people"
    Dim FieldNames As Variant
    Dim FieldTypeValues As Variant
    Dim Records As Variant
    FieldNames = Array("id", "FirstName", "LastName", "Age", "Gender", "Email", "Country", "Domain")
    FieldTypeValues = Array(0, " ", " ", 0, " ", " ", " ", " ")
    Records = Array(Array(1, "Teresa", "Glover", 52, "male", "Teresa.Glover@bol.com.br_", "Macedonia_", "hiphopmyway.com_"), _
                    Array(5, "Kimble", "Graeme", 57, "male", "Kimble.Graeme@yahoo.co.id_", "Suriname_", "nvhrlq.cc_"))

    Dim AdoCommand As ADODB.Command
    Set AdoCommand = GetAdoCommand(TableName, FieldNames, FieldTypeValues)

    Debug.Assert LBound(FieldNames, 1) = LBound(Records(0), 1)
    Debug.Assert UBound(FieldNames, 1) = UBound(Records(0), 1)
    
    Dim RecordsAffected As Long: RecordsAffected = 0
    Dim FieldIndex As Long
    Dim RecordIndex As Long
    AdoCommand.ActiveConnection.BeginTrans
    For RecordIndex = LBound(Records, 1) To UBound(Records, 1)
        For FieldIndex = LBound(FieldNames, 1) To UBound(FieldNames, 1)
            AdoCommand.Parameters(FieldNames(FieldIndex)).Size = GetAdoParamSize(Records(RecordIndex)(FieldIndex))
            AdoCommand.Parameters(FieldNames(FieldIndex)).Value = Records(RecordIndex)(FieldIndex)
        Next FieldIndex
        AdoCommand.Execute RecordsAffected
    Next RecordIndex
    AdoCommand.ActiveConnection.CommitTrans
End Sub


