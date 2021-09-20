VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "SQLiteDB.Drafts.Helper - Working"
'@PredeclaredId
Option Explicit


'''' Formats a quoted field list for an SQL query
''''
'''' Generates a string containing comma-separated list of quoted field names
'''' (no table qualification) suitable for an SQL query.
''''
'''' Args:
''''   FieldNameList:
''''     Accepts either a variable-length argument list of scalar string values
''''     or the same list as a 0-based 1D array argument. UnfoldParamArray is
''''     used to reduce both variants to a 1D array.
''''
'''' Returns:
''''   String, containing comma-separated list of quoted field names
''''
'''' Examples:
''''   >>> ?SQLUtils.FieldsQ(Array("id", "fisrtName", "age"))
''''   "[id], [fisrtName], [age]"
''''
''''   >>> ?SQLUtils.FieldsQ("id", "fisrtName", "age")
''''   "[id], [fisrtName], [age]"
''''
'''' @ClassMethod
'''' This method can also be used on the default instance
''''
'@Description "Fields quoted: [<FieldName>],..."
Public Function FieldsQ(ParamArray FieldNameList() As Variant) As Variant
Attribute FieldsQ.VB_Description = "Fields quoted: [<FieldName>],..."
    Guard.Expression UBound(FieldNameList) >= 0, "SQLiteDB", "At least one field name is required"
    
    Dim FieldNames As Variant
    FieldNames = UnfoldParamArray(FieldNameList)
    Dim QuotedFields() As String
    
    ReDim QuotedFields(LBound(FieldNames) To UBound(FieldNames))
    Dim FieldIndex As Long
    For FieldIndex = LBound(FieldNames) To UBound(FieldNames)
        QuotedFields(FieldIndex) = "[" & FieldNames(FieldIndex) & "]"
    Next FieldIndex
    FieldsQ = Join(QuotedFields, ", ")
End Function


'''' Formats a quoted aliased field list for an SQL query
''''
'''' Generates a string containing comma-separated list of quoted aliased field
'''' names (no table qualification) suitable for an SQL query.
''''
'''' Args:
''''   FieldNameList:
''''     Accepts either a variable-length argument list of scalar string values
''''     or the same list as a 0-based 1D array argument. UnfoldParamArray is
''''     used to reduce both variants to a 1D array.
''''
'''' Returns:
''''   String, containing comma-separated list of quoted aliased field names
''''
'''' Examples:
''''   >>> ?SQLUtils.FieldsQA(Array("id", "fisrtName", "age"))
''''   "[id] AS [id], [fisrtName] AS [fisrtName], [age] AS [age]"
''''
''''   >>> ?SQLUtils.FieldsQA("id", "fisrtName", "age")
''''   "[id] AS [id], [fisrtName] AS [fisrtName], [age] AS [age]"
''''
'''' @ClassMethod
'''' This method can also be used on the default instance
''''
'@Description "Fields quoted and aliased: [<FieldName>] AS [<FieldName>],..."
Public Function FieldsQA(ParamArray FieldNameList() As Variant) As String
Attribute FieldsQA.VB_Description = "Fields quoted and aliased: [<FieldName>] AS [<FieldName>],..."
    Guard.Expression UBound(FieldNameList) >= 0, "SQLiteDB", "At least one field name is required"
    
    Dim FieldNames As Variant
    FieldNames = UnfoldParamArray(FieldNameList)
    
    Dim AliasedFields() As String
    ReDim AliasedFields(LBound(FieldNames) To UBound(FieldNames))
    Dim FieldIndex As Long
    For FieldIndex = LBound(FieldNames) To UBound(FieldNames)
        AliasedFields(FieldIndex) = FieldNames(FieldIndex) & "] AS [" & FieldNames(FieldIndex)
    Next FieldIndex
    FieldsQA = "[" & Join(AliasedFields, "], [") & "]"
End Function
