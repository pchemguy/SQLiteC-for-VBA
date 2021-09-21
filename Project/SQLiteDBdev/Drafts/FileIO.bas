Attribute VB_Name = "FileIO"
'@Folder "SQLiteDBdev.Drafts"
Option Explicit
Option Private Module
Option Compare Text

Private Const LIB_NAME As String = "SQLiteDBVBA"
Private Const PATH_SEP As String = "\"
Private Const REL_PREFIX As String = "Library" & PATH_SEP & LIB_NAME & PATH_SEP




Private Function GetFileBytes(ByVal FilePathName As String) As Byte()
    Dim FileHandle As Long
    Dim ReadBuffer(0 To 15) As Byte
    FileHandle = FreeFile
    Open FilePathName For Binary Access Read Write As FileHandle
    Get FileHandle, , ReadBuffer
    Close FileHandle
    GetFileBytes = ReadBuffer
End Function


Public Sub TestGetFileBytes()
    Dim FilePathName As String
    FilePathName = ThisWorkbook.Path & "\Library\SQLiteDBVBA\TestC.db"
    Dim ByteBuffer() As Byte
    ByteBuffer = GetFileBytes(FilePathName)
    Dim TextBuffer As String
    TextBuffer = StrConv(ByteBuffer, vbUnicode)
    Dim SQLiteDbSignature As String
    SQLiteDbSignature = "SQLite format 3" & vbNullChar
    Debug.Print SQLiteDbSignature = TextBuffer
End Sub
    
'   Set DbTextStream = fso.GetFile(FilePathName & "-shm").OpenAsTextStream(ForReading, TRISTATE_OPEN_AS_ASCII)
'   Dim Buffer As String
'   Buffer = DbTextStream.Read(4)
'
'    Debug.Print fso.FolderExists("F:\Archive\Business\FID\PolMaFID\Drafts\Knowledge Management System\VBA\SQLiteDB VBA Library\Library\SQLiteDBVBA\")
'
'    Debug.Print FileLen("F:\Archive\Business\FID\PolMaFID\Drafts\Knowledge Management System\VBA\SQLiteDB VBA Library\Library\SQLiteDBVBA\QQ")
    
    '''' Split FilePathName
'    Dim SplitPosition As Long
'    SplitPosition = InStrRev(FilePathName, Application.PathSeparator)
'    Dim FileName As String
'    FileName = Mid$(FilePathName, SplitPosition + 1)
'
'    Const MagicHeaderString As String = "SQLite format 3" & vbNullChar
'    Dim FileHandle As Long
'    Dim ReadBuffer(0 To Len(MagicHeaderString) - 1) As Byte
'
'    FileHandle = FreeFile
'    Open FilePathName For Binary Access Read Write As FileHandle
'    Get FileHandle, , ReadBuffer
'    Close FileHandle
'    VerifyMagicHeader = (MagicHeaderString = StrConv(ReadBuffer, vbUnicode))
