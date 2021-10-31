VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FixGuard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "SQLiteCADO guards"
'@Folder "SQLite.Fixtures"
'@ModuleDescription "SQLiteCADO guards"
'@PredeclaredId
Option Explicit

Private Type TFixGuard
    ModuleName As String
End Type
Private this As TFixGuard


'@DefaultMember
Public Function Create(ByVal ModuleName As String) As FixGuard
Attribute Create.VB_UserMemId = 0
    Dim Instance As FixGuard
    Set Instance = New FixGuard
    Instance.Init ModuleName
    Set Create = Instance
End Function


Friend Sub Init(ByVal ModuleName As String)
    Guard.EmptyString ModuleName
    this.ModuleName = ModuleName
End Sub


'@Description("Raises a run-time error if the specified connection object is not opened.")
Public Sub DbNotOpened(ByVal dbc As SQLiteCConnection, _
              Optional ByVal ErrorSource As String = vbNullString)
Attribute DbNotOpened.VB_Description = "Raises a run-time error if the specified connection object is not opened."
    Dim Source As String
    Source = this.ModuleName & IIf(Len(this.ModuleName) > 0 And Len(ErrorSource) > 0, _
        " / ", vbNullString) & ErrorSource
    
    Dim errorDetails As TError
    With errorDetails
        .Trapped = dbc Is Nothing
        If .Trapped Then
            .Name = "ObjectNotSetErr"
            .Number = ErrNo.ObjectNotSetErr
            .Source = Source
            .Message = "SQLiteCConnection object is not set."
            .Description = .Message & " Source: " & .Source
            Guard.RaiseError errorDetails
        End If
        
        .Trapped = (dbc.DbHandle = 0)
        If .Trapped Then
            .Name = "ConnectionNotOpenedErr"
            .Number = ErrNo.ConnectionNotOpenedErr
            .Source = Source
            .Message = "SQLiteCConnection object is not opened."
            .Description = .Message & " Source: " & .Source
            Guard.RaiseError errorDetails
        End If
    End With
End Sub


'@Description("Raises a run-time error if the specified statement object is not prepared.")
Public Sub StmtNotPrepared(ByVal dbs As SQLiteCStatement, _
              Optional ByVal ErrorSource As String = vbNullString)
Attribute StmtNotPrepared.VB_Description = "Raises a run-time error if the specified statement object is not prepared."
    Dim Source As String
    Source = this.ModuleName & IIf(Len(this.ModuleName) > 0 And Len(ErrorSource) > 0, _
        " / ", vbNullString) & ErrorSource
    
    Dim errorDetails As TError
    With errorDetails
        .Trapped = dbs Is Nothing
        If .Trapped Then
            .Name = "ObjectNotSetErr"
            .Number = ErrNo.ObjectNotSetErr
            .Source = Source
            .Message = "SQLiteCStatement object is not set."
            .Description = .Message & " Source: " & .Source
            Guard.RaiseError errorDetails
        End If
        
        .Trapped = (dbs.StmtHandle = 0)
        If .Trapped Then
            .Name = "StatementNotPreparedErr"
            .Number = ErrNo.StatementNotPreparedErr
            .Source = Source
            .Message = "SQLiteCStatement object is not prepared."
            .Description = .Message & " Source: " & .Source
            Guard.RaiseError errorDetails
        End If
    End With
End Sub
