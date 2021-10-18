Attribute VB_Name = "Module1"
'@Folder("SQLiteC For VBA.RecordsetADO")
'@IgnoreModule
Option Explicit

Public Sub FabricateRecordset()
    Dim objRs As New ADODB.Recordset
    With objRs.Fields
        .Append Name:="StudentID", Type:=adVarWChar, DefinedSize:=2 ^ 16 - 1, Attrib:=adFldUpdatable
        .Append Name:="FullName", Type:=adVarWChar, DefinedSize:=2 ^ 16 - 1, Attrib:=adFldUpdatable
        .Append Name:="PhoneNmbr", Type:=adVarWChar, DefinedSize:=2 ^ 16 - 1, Attrib:=adFldUpdatable
    End With
    With objRs
        .Open
        .AddNew
        .Fields(0) = "123-45-6789"
        .Fields(1) = "John Doe"
        .Fields(2) = "(425) 555-5555"
        .AddNew
        .Fields(0) = "123-45-6780"
        .Fields(1) = "Jane Doe"
        .Fields(2) = "(615) 555-1212"
        .UpdateBatch
        
        .Close
    End With
End Sub

