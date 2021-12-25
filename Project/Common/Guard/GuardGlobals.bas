Attribute VB_Name = "GuardGlobals"
'@Folder "Common.Guard"
Option Explicit

Public Enum ADODBErrorValueEnum
    ADOadErrFeatureNotAvailable = 3251&
    ADOadErrInTransaction = 3246&
    ADOadErrInvalidTransaction = 3714&
    ADOadErrProviderNotFound = 3706&
    ADOadErrInvalidParamInfo = 3708&
    ADOadErrProviderFailed = 3000&
End Enum

Public Enum ErrNo
    PassedNoErr = 0&
    InvalidProcedureCallOrArgumentErr = 5&
    OutOfMemoryErr = 7&
    SubscriptOutOfRange = 9&
    TypeMismatchErr = 13&
    BadFileNameOrNumberErr = 52&
    FileNotFoundErr = 53&
    FileAlreadyExistsErr = 58&
    PermissionDeniedErr = 70&
    PathFileAccessErr = 75&
    PathNotFoundErr = 76&
    ObjectNotSetErr = 91&
    ObjectRequiredErr = 424&
    InvalidObjectUseErr = 425&
    ObjectCreateErr = 426&
    MemberNotExistErr = 438&
    ActionNotSupportedErr = 445&
    KeyAlreadyExistsErr = 457&
    InvalidParameterErr = 1004&
    NoObject = 31004&
        
    ConnectionNotOpenedErr = vbObjectError + 5001
    StatementNotPreparedErr = vbObjectError + 5002
    ConnectionOpenErr = vbObjectError + 5003
    ConnectionCloseErr = vbObjectError + 5004
    
    CustomErr = VBA.vbObjectError + 1000&
    NotImplementedErr = VBA.vbObjectError + 1001&
    IncompatibleArraysErr = VBA.vbObjectError + 1002&
    IncompatibleStatusErr = VBA.vbObjectError + 1003&
    DefaultInstanceErr = VBA.vbObjectError + 1011&
    NonDefaultInstanceErr = VBA.vbObjectError + 1012&
    EmptyStringErr = VBA.vbObjectError + 1013&
    SingletonErr = VBA.vbObjectError + 1014&
    UnknownClassErr = VBA.vbObjectError + 1015&
    EmptyOrNullErr = VBA.vbObjectError + 1016&
    ObjectSetErr = VBA.vbObjectError + 1091&
    ExpectedArrayErr = VBA.vbObjectError + 2013&
    InvalidCharacterErr = VBA.vbObjectError + 2014&
    ConsistencyCheckErr = VBA.vbObjectError + 2024&
    IntegrityCheckErr = VBA.vbObjectError + 2034&
    TextStreamReadErr = &H80070021
    OLE_DB_ODBC_Err = &H80004005
    AdoFeatureNotAvailableErr = ADOadErrFeatureNotAvailable
    AdoInTransactionErr = ADOadErrInTransaction
    AdoInvalidTransactionErr = ADOadErrInvalidTransaction
    AdoConnectionStringErr = ADOadErrProviderNotFound
    AdoInvalidParamInfoErr = ADOadErrInvalidParamInfo
    AdoProviderFailedErr = ADOadErrProviderFailed
End Enum

Public Type TError
    Number As ErrNo
    Name As String
    Source As String
    Message As String
    Stack As String
    Description As String
    Trapped As Boolean
End Type
