Attribute VB_Name = "AlphaNumLib"
'@Folder "SQLiteDBdev.Drafts"
'@IgnoreModule
Option Explicit


Private Sub Templates()
    Dim AlphaChar As String * 52
    AlphaChar = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    Dim NumChar As String
    NumChar = "0123456789"
    Dim AlphaExChar As String
    AlphaExChar = "_"
    Dim NonAlphaNumPrintChar As String
    NonAlphaNumPrintChar = " !""#$%&'()*+,-./:;<=>?@[\]^_`{|}~"
    Dim CharMap(0 To 255) As Byte
End Sub
