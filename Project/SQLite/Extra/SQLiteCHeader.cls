VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLiteCHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Parses database file header."
'@Folder "SQLite.Extra"
'@ModuleDescription "Parses database file header."
'@PredeclaredId
Option Explicit

Private Type TSQLiteCHeader
    DbPathName As String
    Header As SQLiteCHeaderData
End Type
Private this As TSQLiteCHeader


Public Function Create(Optional ByVal DbPathName As String = vbNullString _
                      ) As SQLiteCHeader
    Dim Instance As SQLiteCHeader
    Set Instance = New SQLiteCHeader
    Instance.Init DbPathName
    Set Create = Instance
End Function


Friend Sub Init(Optional ByVal DbPathName As String = vbNullString)
    '@Ignore SelfAssignedDeclaration
    Dim fso As New Scripting.FileSystemObject
    If Len(DbPathName) > 0 And Not fso.FileExists(DbPathName) Then
        Err.Raise ErrNo.FileNotFoundErr, "SQLiteCHeader/Init", "Db file not found."
    End If
    this.DbPathName = DbPathName
End Sub


Public Property Get Header() As SQLiteCHeaderData
    Header = this.Header
End Property


Public Property Let Header(ByRef Value As SQLiteCHeaderData)
    this.Header = Value
End Property


'''' Reference: https://sqlite.org/fileformat.html
''''
'''' SQLite database header follows big-endian convention. This routine
'''' loads the header into a temporary packed strucuture with all multi-byte
'''' variables defined as byte arrays. Then, it calls unpacker routine for
'''' further processing.
''''
'@Description "Loads and parses SQLite database file header."
Public Sub LoadHeader()
Attribute LoadHeader.VB_Description = "Loads and parses SQLite database file header."
    Const BlockNumber As Long = 1
    Dim Buffer As SQLiteCHeaderPacked
    '@Ignore IntegerDataType
    Dim FileHandle As Integer
    FileHandle = FreeFile()
    Open this.DbPathName For Random As FileHandle Len = Len(Buffer)
    Get FileHandle, BlockNumber, Buffer
    Close FileHandle
    UnpackHeader Buffer
End Sub


'''' Reference: https://sqlite.org/fileformat.html
''''
'''' SQLite database header follows big-endian convention. This routine unpacks
'''' the header into the structure with native variable types.
''''
'@Description "Unpacks SQLite database file header."
Public Sub UnpackHeader(ByRef PackedHeader As SQLiteCHeaderPacked)
Attribute UnpackHeader.VB_Description = "Unpacks SQLite database file header."
    With this.Header
        .MagicHeaderString = StrConv(PackedHeader.MagicHeaderString, vbUnicode)
        .PageSizeInBytes = FixUtils.SInt16bFromBytesBE(PackedHeader.PageSizeInBytes)
        .FileFormatWrite = PackedHeader.FileFormatWrite
        .FileFormatRead = PackedHeader.FileFormatRead
        .ReservedSpace = PackedHeader.ReservedSpace
        .MaxPayload = PackedHeader.MaxPayload
        .MinPayload = PackedHeader.MinPayload
        .LeafPayload = PackedHeader.LeafPayload
        .ChangeCounter = FixUtils.SLong32bFromBytesBE(PackedHeader.ChangeCounter)
        .DbFilePageCount = FixUtils.SLong32bFromBytesBE(PackedHeader.DbFilePageCount)
        .FirstFreeListPage = FixUtils.SLong32bFromBytesBE(PackedHeader.FirstFreeListPage)
        .FreeListPageCount = FixUtils.SLong32bFromBytesBE(PackedHeader.FreeListPageCount)
        .SchemaCookie = FixUtils.SLong32bFromBytesBE(PackedHeader.SchemaCookie)
        .SchemaFormat = FixUtils.SLong32bFromBytesBE(PackedHeader.SchemaFormat)
        .DefaultPageCacheSize = FixUtils.SLong32bFromBytesBE(PackedHeader.DefaultPageCacheSize)
        .LagestBTreeRootPage = FixUtils.SLong32bFromBytesBE(PackedHeader.LagestBTreeRootPage)
        .DbTextEncoding = FixUtils.SLong32bFromBytesBE(PackedHeader.DbTextEncoding)
        .UserVersion = FixUtils.SLong32bFromBytesBE(PackedHeader.UserVersion)
        .VacuumMode = FixUtils.SLong32bFromBytesBE(PackedHeader.VacuumMode)
        .AppId = FixUtils.SLong32bFromBytesBE(PackedHeader.AppId)
        .Reserved = PackedHeader.Reserved
        .VersionValidFor = FixUtils.SLong32bFromBytesBE(PackedHeader.VersionValidFor)
        .SqliteVersion = FixUtils.SLong32bFromBytesBE(PackedHeader.SqliteVersion)
    End With
End Sub
