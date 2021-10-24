Attribute VB_Name = "GuardFixtures"
'@Folder "Common.Guard"
Option Explicit

Public Enum ErrNo
    PassedNoErr = 0&
    InvalidProcedureCallOrArgumentErr = 5&
    OutOfMemoryErr = 7&
    SubscriptOutOfRange = 9&
    TypeMismatchErr = 13&
    BadFileNameOrNumberErr = 52&
    FileNotFoundErr = 53&
    PermissionDeniedErr = 70&
    PathFileAccessErr = 75&
    PathNotFoundErr = 76&
    ObjectNotSetErr = 91&
    ObjectRequiredErr = 424&
    InvalidObjectUseErr = 425&
    MemberNotExistErr = 438&
    ActionNotSupportedErr = 445&
    KeyAlreadyExistsErr = 457&
    InvalidParameterErr = 1004&
    NoObject = 31004&
        
    CustomErr = VBA.vbObjectError + 1000&
    NotImplementedErr = VBA.vbObjectError + 1001&
    IncompatibleArraysErr = VBA.vbObjectError + 1002&
    IncompatibleStatusErr = VBA.vbObjectError + 1003&
    DefaultInstanceErr = VBA.vbObjectError + 1011&
    NonDefaultInstanceErr = VBA.vbObjectError + 1012&
    EmptyStringErr = VBA.vbObjectError + 1013&
    SingletonErr = VBA.vbObjectError + 1014&
    UnknownClassErr = VBA.vbObjectError + 1015&
    ObjectSetErr = VBA.vbObjectError + 1091&
    ExpectedArrayErr = VBA.vbObjectError + 2013&
    InvalidCharacterErr = VBA.vbObjectError + 2014&
    ConsistencyCheckErr = VBA.vbObjectError + 2024&
    IntegrityCheckErr = VBA.vbObjectError + 2034&
    ConnectionNotOpenedErr = vbObjectError + 3000
    StatementNotPreparedErr = vbObjectError + 3001
    TextStreamReadErr = &H80070021
    OLE_DB_ODBC_Err = &H80004005
    AdoFeatureNotAvailableErr = ADODB.ErrorValueEnum.adErrFeatureNotAvailable
    AdoInTransactionErr = ADODB.ErrorValueEnum.adErrInTransaction
    AdoInvalidTransactionErr = ADODB.ErrorValueEnum.adErrInvalidTransaction
    AdoConnectionStringErr = ADODB.ErrorValueEnum.adErrProviderNotFound
    AdoInvalidParamInfoErr = ADODB.ErrorValueEnum.adErrInvalidParamInfo
    AdoProviderFailedErr = ADODB.ErrorValueEnum.adErrProviderFailed
End Enum

Public Type TError
    Number As ErrNo
    Name As String
    Source As String
    Message As String
    Description As String
    Trapped As Boolean
End Type


'@Ignore ProcedureNotUsed
'@Description("Re-raises the current error, if there is one.")
Public Sub RethrowOnError()
Attribute RethrowOnError.VB_Description = "Re-raises the current error, if there is one."
    With VBA.Err
        If .Number <> 0 Then
            Debug.Print "Error " & .Number, .Description
            .Raise .Number
        End If
    End With
End Sub


'@Description("Formats and raises a run-time error.")
Public Sub RaiseError(ByRef errorDetails As TError)
Attribute RaiseError.VB_Description = "Formats and raises a run-time error."
    With errorDetails
        Dim Message As Variant
        Message = Array("Error:", _
            "name: " & .Name, _
            "number: " & .Number, _
            "message: " & .Message, _
            "description: " & .Description, _
            "source: " & .Source)
        Debug.Print Join(Message, vbNewLine & vbTab)
        VBA.Err.Raise .Number, .Source, .Message
    End With
End Sub
