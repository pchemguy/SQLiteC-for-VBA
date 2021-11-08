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
'@Folder "SQLiteDBdev.Extra"
'@ModuleDescription "Parses database file header."
'@PredeclaredId
'@IgnoreModule IntegerDataType
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByRef pDest As Any, ByRef pSource As Any, ByVal Length As Long)
#Else
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef pDest As Any, ByRef pSource As Any, ByVal Length As Long)
#End If

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


'@Ignore ProcedureNotUsed
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
        .PageSizeInBytes = SInt16bFromBytesBE(PackedHeader.PageSizeInBytes)
        .FileFormatWrite = PackedHeader.FileFormatWrite
        .FileFormatRead = PackedHeader.FileFormatRead
        .ReservedSpace = PackedHeader.ReservedSpace
        .MaxPayload = PackedHeader.MaxPayload
        .MinPayload = PackedHeader.MinPayload
        .LeafPayload = PackedHeader.LeafPayload
        .ChangeCounter = SLong32bFromBytesBE(PackedHeader.ChangeCounter)
        .DbFilePageCount = SLong32bFromBytesBE(PackedHeader.DbFilePageCount)
        .FirstFreeListPage = SLong32bFromBytesBE(PackedHeader.FirstFreeListPage)
        .FreeListPageCount = SLong32bFromBytesBE(PackedHeader.FreeListPageCount)
        .SchemaCookie = SLong32bFromBytesBE(PackedHeader.SchemaCookie)
        .SchemaFormat = SLong32bFromBytesBE(PackedHeader.SchemaFormat)
        .DefaultPageCacheSize = SLong32bFromBytesBE(PackedHeader.DefaultPageCacheSize)
        .LagestBTreeRootPage = SLong32bFromBytesBE(PackedHeader.LagestBTreeRootPage)
        .DbTextEncoding = SLong32bFromBytesBE(PackedHeader.DbTextEncoding)
        .UserVersion = SLong32bFromBytesBE(PackedHeader.UserVersion)
        .VacuumMode = SLong32bFromBytesBE(PackedHeader.VacuumMode)
        .AppId = SLong32bFromBytesBE(PackedHeader.AppId)
        .Reserved = PackedHeader.Reserved
        .VersionValidFor = SLong32bFromBytesBE(PackedHeader.VersionValidFor)
        .SQLiteVersion = SLong32bFromBytesBE(PackedHeader.SQLiteVersion)
    End With
End Sub


Friend Function PackedHeaderFromBytes(ByRef HeaderBuffer() As Byte _
                                     ) As SQLiteCHeaderPacked
    Dim PackedHeader As SQLiteCHeaderPacked
    If LBound(HeaderBuffer) <> 0 Or _
       UBound(HeaderBuffer) <> Len(PackedHeader) - 1 Then
        Err.Raise 9, "SQLiteCHeader/PackedHeaderFromBytes", "Expected Array(0 To 99)"
    End If
    RtlMoveMemory PackedHeader, HeaderBuffer(0), Len(PackedHeader)
    PackedHeaderFromBytes = PackedHeader
End Function


'@Description "Converts 2-byte big-endian array to signed 16-bit integer (Integer)"
Friend Function SInt16bFromBytesBE(ByRef BytesBE() As Byte) As Integer
Attribute SInt16bFromBytesBE.VB_Description = "Converts 2-byte big-endian array to signed 16-bit integer (Integer)"
    Dim SignBit As Integer
    SignBit = -CInt((BytesBE(0) And &H80) / &H80) And &H8000
    SInt16bFromBytesBE = CInt(BytesBE(0) And &H7F) * &H100 + CInt(BytesBE(1))
    SInt16bFromBytesBE = SInt16bFromBytesBE Or SignBit
End Function


'@Description "Converts 4-byte big-endian array to signed 32-bit integer (Long)"
Friend Function SLong32bFromBytesBE(ByRef BytesBE() As Byte) As Long
Attribute SLong32bFromBytesBE.VB_Description = "Converts 4-byte big-endian array to signed 32-bit integer (Long)"
    Dim SignBit As Long
    SignBit = -CLng((BytesBE(0) And &H80) / &H80) And &H80000000
    SLong32bFromBytesBE = ((CLng(BytesBE(0) And &H7F) _
                 * &H100& + CLng(BytesBE(1))) _
                 * &H100& + CLng(BytesBE(2))) _
                 * &H100& + CLng(BytesBE(3))
    SLong32bFromBytesBE = SLong32bFromBytesBE Or SignBit
End Function


'@Description "Flips 32-bit Long byte order"
Friend Function SLong32bFlipBytes(ByVal Source As Long) As Long
Attribute SLong32bFlipBytes.VB_Description = "Flips 32-bit Long byte order"
    Dim BytesLE(0 To 3) As Byte
    BytesLE(0) = Source And &HFF&
    BytesLE(1) = (Source And &HFF00&) / 256&
    BytesLE(2) = (Source And &HFF0000) / (256& * 256&)
    BytesLE(3) = (Source And &H7F000000) / (256& * 256& * 256&) + _
                 128& * Abs(CBool(Source And &H80000000))
    SLong32bFlipBytes = SLong32bFromBytesBE(BytesLE)
End Function


Public Function GenBlankDb(Optional ByVal UserVersion As Long = 0, _
                           Optional ByVal ApplicationId As Long = 0) As Long()
    '''' Allocate buffer equal to one default SQLite page and set it to 0
    Const SQLITE_PAGE_IN_BYTES As Long = 4096&
    Dim DbPage() As Long
    ReDim DbPage(0 To SQLITE_PAGE_IN_BYTES / Len(DbPage(0)) - 1)
    Dim Blanks() As Byte
    Blanks = StrConv(String(SQLITE_PAGE_IN_BYTES, vbNullChar), vbFromUnicode)
    RtlMoveMemory DbPage(0), Blanks(0), SQLITE_PAGE_IN_BYTES
    
    Dim Buffer() As Byte
    Dim ByteCount As Long
    Dim HexString As String
    Dim DestOffsetIndex As Long
    
    '''' ===== BEGIN HEADER =====
    
    '''' DbPage(0) = &H694C5153
    '''' DbPage(1) = &H66206574
    '''' DbPage(2) = &H616D726F
    '''' DbPage(3) = &H332074
    Buffer = StrConv("SQLite format 3" & vbNullChar, vbFromUnicode)
    DestOffsetIndex = 0
    ByteCount = Len(Buffer(0)) * (UBound(Buffer) - LBound(Buffer) + 1)
    RtlMoveMemory DbPage(DestOffsetIndex), Buffer(0), ByteCount
    
    '''' DbPage(4) = &H1010010
    '''' DbPage(5) = &H20204000
    '''' DbPage(6) = &H1000000
    '''' DbPage(7) = &H1000000
    HexString = "0x|10 00 01 01|00 40 20 20|00 00 00 01|00 00 00 01|"
    Buffer = FixUtils.BytesFromHexString(HexString)
    DestOffsetIndex = 4
    ByteCount = Len(Buffer(0)) * (UBound(Buffer) - LBound(Buffer) + 1)
    RtlMoveMemory DbPage(DestOffsetIndex), Buffer(0), ByteCount
    
    DbPage(10) = &H1000000
    DbPage(15) = SLong32bFlipBytes(UserVersion)
    DbPage(17) = SLong32bFlipBytes(ApplicationId)
    
    DbPage(23) = &H1000000
    DbPage(24) = &H48572E00
    
    '''' ===== END OF HEADER =====
    
    DbPage(25) = &HD&
    DbPage(26) = &H1000&
    
    GenBlankDb = DbPage
End Function


Public Function GenBlankDbHeader(Optional ByVal UserVersion As Long = 0, _
                                 Optional ByVal ApplicationId As Long = 0) As Byte()
    Dim PackedHeader As SQLiteCHeaderPacked
    Dim HeaderSize As Long
    HeaderSize = Len(PackedHeader)
    Dim HeaderBuffer() As Byte
    ReDim HeaderBuffer(0 To HeaderSize - 1)
    HeaderBuffer = String(50, vbNullChar)
    Dim HeaderPage() As Long
    HeaderPage = GenBlankDb(UserVersion, ApplicationId)
    RtlMoveMemory HeaderBuffer(0), HeaderPage(0), HeaderSize
    GenBlankDbHeader = HeaderBuffer
End Function
