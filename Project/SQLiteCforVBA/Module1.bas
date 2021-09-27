Attribute VB_Name = "Module1"
'@Folder "SQLiteCforVBA"
Option Explicit



Public Sub TestTest()
    Dim StrBuffer As String
    Dim VarBuffer As Variant
    Dim ByteBuffer() As Byte
    StrBuffer = "ABCÀÁÂ" & vbNullChar & vbNullChar
    VarBuffer = StringToUtf8Bytes(StrBuffer)
    ByteBuffer = StrConv(StrBuffer, vbFromUnicode)
End Sub
