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
'@Folder "SQLite.SQLiteC.Connection"
'@ModuleDescription "Parses database file header."
'@PredeclaredId
Option Explicit

Private Type TSQLiteCHeader
    dbc As SQLiteCConnection
    Header As SQLiteCHeaderData
End Type
Private this As TSQLiteCHeader


'@DefaultMember
Public Function Create(ByVal dbc As SQLiteCConnection) As SQLiteCHeader
Attribute Create.VB_UserMemId = 0
    Dim Instance As SQLiteCHeader
    Set Instance = New SQLiteCHeader
    Instance.Init dbc
    Set Create = Instance
End Function


Friend Sub Init(ByVal dbc As SQLiteCConnection)
    Guard.NullReference dbc, "SQLiteCHeader/Init"
    Set this.dbc = dbc
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
'''' variables defined as byte arrays. Then, it converts and copies the data
'''' into the header structure with native variable types.
''''
'@Description "Loads and parses SQLite database file header."
Public Sub LoadHeader()
Attribute LoadHeader.VB_Description = "Loads and parses SQLite database file header."
    Const HeaderSize As Long = 100
    Const BlockNumber As Long = 1
    Dim Buffer As SQLiteCHeaderPacked
    '@Ignore IntegerDataType
    Dim FileHandle As Integer
    FileHandle = FreeFile()
    Open this.dbc.DbPathName For Random As FileHandle Len = Len(Buffer)
    Get FileHandle, BlockNumber, Buffer
    Close FileHandle
    With this.Header
        .MagicHeaderString = StrConv(Buffer.MagicHeaderString, vbUnicode)
        .PageSizeInBytes = FixUtils.SInt16bFromBytesBE(Buffer.PageSizeInBytes)
        .FileFormatWrite = Buffer.FileFormatWrite
        .FileFormatRead = Buffer.FileFormatRead
        .ReservedSpace = Buffer.ReservedSpace
        .MaxPayload = Buffer.MaxPayload
        .MinPayload = Buffer.MinPayload
        .LeafPayload = Buffer.LeafPayload
        .ChangeCounter = FixUtils.SLong32bFromBytesBE(Buffer.ChangeCounter)
        .DbFilePageCount = FixUtils.SLong32bFromBytesBE(Buffer.DbFilePageCount)
        .FirstFreeListPage = FixUtils.SLong32bFromBytesBE(Buffer.FirstFreeListPage)
        .FreeListPageCount = FixUtils.SLong32bFromBytesBE(Buffer.FreeListPageCount)
        .SchemaCookie = FixUtils.SLong32bFromBytesBE(Buffer.SchemaCookie)
        .SchemaFormat = FixUtils.SLong32bFromBytesBE(Buffer.SchemaFormat)
        .DefaultPageCacheSize = FixUtils.SLong32bFromBytesBE(Buffer.DefaultPageCacheSize)
        .LagestBTreeRootPage = FixUtils.SLong32bFromBytesBE(Buffer.LagestBTreeRootPage)
        .DbTextEncoding = FixUtils.SLong32bFromBytesBE(Buffer.DbTextEncoding)
        .UserVersion = FixUtils.SLong32bFromBytesBE(Buffer.UserVersion)
        .VacuumMode = FixUtils.SLong32bFromBytesBE(Buffer.VacuumMode)
        .AppId = FixUtils.SLong32bFromBytesBE(Buffer.AppId)
        .Reserved = Buffer.Reserved
        .VersionValidFor = FixUtils.SLong32bFromBytesBE(Buffer.VersionValidFor)
        .SqliteVersion = FixUtils.SLong32bFromBytesBE(Buffer.SqliteVersion)
    End With
End Sub
