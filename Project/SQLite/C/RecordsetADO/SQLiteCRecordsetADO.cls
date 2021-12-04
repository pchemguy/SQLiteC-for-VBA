VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLiteCRecordsetADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Constructs ADO Recordset"
'@Folder "SQLite.C.RecordsetADO"
'@ModuleDescription "Constructs ADO Recordset using table metadata."
'@Exposed
'@PredeclaredId
Option Explicit

Private Type TSQLiteCRecordsetADO
    DbExec As SQLiteCExecSQL
    SQLQuery As String
    SQLQueryExpanded As String
    SQLQueryParameters As Variant
    FabRecordset As ADODB.Recordset
End Type
Private this As TSQLiteCRecordsetADO


'@DefaultMember
Public Function Create(ByVal DbStmt As SQLiteCStatement) As SQLiteCRecordsetADO
Attribute Create.VB_UserMemId = 0
    Dim Instance As SQLiteCRecordsetADO
    Set Instance = New SQLiteCRecordsetADO
    Instance.Init DbStmt
    Set Create = Instance
End Function

Friend Sub Init(ByVal DbStmt As SQLiteCStatement)
    Guard.NullReference DbStmt
    With this
        Set .DbExec = DbStmt.DbExecutor
        Set .FabRecordset = New ADODB.Recordset
        AddMeta
        .SQLQuery = DbStmt.SQLQueryOriginal
        .SQLQueryExpanded = DbStmt.SQLQueryExpanded
        .FabRecordset.Source = DbStmt.SQLQueryExpanded
    End With
End Sub

Public Property Get AdoRecordset() As ADODB.Recordset
    Set AdoRecordset = this.FabRecordset
End Property

Friend Sub AddMeta()
    If Not this.DbExec.MetaLoaded Then
        Err.Raise ErrNo.IncompatibleStatusErr, "SQLiteCRecordsetADO", _
                  "Metadata has not been loaded."
    End If
    With this.FabRecordset
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
    End With
    Dim TableMeta() As SQLiteCColumnMeta
    TableMeta = this.DbExec.TableMeta
    Dim ColumnIndex As Long
    For ColumnIndex = 0 To UBound(TableMeta)
        With TableMeta(ColumnIndex)
            If .TableMeta Then
                Select Case .AffinityType
                    Case SQLITE_INTEGER
                        .AdoType = adInteger
                    Case SQLITE_FLOAT
                        .AdoType = adDouble
                    Case SQLITE_TEXT
                        .AdoType = adVarWChar
                        .AdoSize = 2 ^ 13
                    Case SQLITE_BLOB
                        .AdoType = adLongVarBinary
                        .AdoAttr = .AdoAttr Or adFldLong
                        .AdoSize = 2 ^ 16 - 1
                    Case Else
                        .AdoType = adVarWChar
                        .AdoSize = 2 ^ 16 - 1
                 End Select
                .AdoAttr = .AdoAttr Or _
                           adFldUpdatable Or _
                           ((Not .NotNull) And (adFldIsNullable Or adFldMayBeNull)) Or _
                           (.PrimaryKey And adFldKeyColumn) Or _
                           ((.AutoIncrement Or LCase$(.Name) = "rowid") And adFldRowID)
            Else
                .AdoType = adVarWChar
                .AdoSize = 2 ^ 16 - 1
                .AdoAttr = adFldUnspecified
                .OriginName = .Name
            End If
            '@Ignore ArgumentWithIncompatibleObjectType: False positive
            this.FabRecordset.Fields.Append .OriginName, .AdoType, .AdoSize, .AdoAttr
        End With
    Next ColumnIndex
    this.DbExec.TableMeta = TableMeta
End Sub

'@Ignore ProcedureNotUsed
Public Sub SQLQuerySet(Optional ByVal SQLQuery As String = vbNullString, _
                       Optional ByVal SQLQueryExpanded As String = vbNullString)
    If Len(SQLQuery) > 0 Then this.SQLQuery = SQLQuery
    If Len(SQLQueryExpanded) > 0 Then this.SQLQueryExpanded = SQLQueryExpanded
End Sub

Public Sub SQLQueryParamsSet(ByVal SQLQueryParameters As Variant)
    Select Case VarType(SQLQueryParameters)
        Case vbObject
            If Not SQLQueryParameters Is Nothing Then _
                    Set this.SQLQueryParameters = SQLQueryParameters
        Case Is >= vbArray
            this.SQLQueryParameters = SQLQueryParameters
    End Select
End Sub
