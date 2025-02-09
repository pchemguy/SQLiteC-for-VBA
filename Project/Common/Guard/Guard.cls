VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Guard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Class for common guards."
'@Folder "Common.Guard"
'@ModuleDescription "Class for common guards."
'@PredeclaredId
'@IgnoreModule ProcedureNotUsed, ParameterNotUsed
Option Explicit
Option Compare Text

Private Const MsgExpectedErrNotRaised As String = "Expected error was not raised."
Private Const MsgUnexpectedErrRaised As String = "Unexpected error was raised."

Private Type TGuard
    ModuleName As String
End Type
'@Ignore MoveFieldCloserToUsage
Private this As TGuard


'@Description("Default class factory")
Public Function Create(ByVal ModuleName As String) As Guard
Attribute Create.VB_Description = "Default class factory"
    Dim Instance As Guard
    Set Instance = New Guard
    Instance.Init ModuleName
    Set Create = Instance
End Function


Friend Sub Init(ByVal ModuleName As String)
    Guard.EmptyString ModuleName
    this.ModuleName = ModuleName
End Sub


'@Description("Returns instance reference")
Public Property Get Self() As Guard
Attribute Self.VB_Description = "Returns instance reference"
    Set Self = Me
End Property


'@Description("Returns class reference")
Public Property Get Class() As Guard
Attribute Class.VB_Description = "Returns class reference"
    Set Class = Guard
End Property


'@Description("Raises a run-time error if the specified string is empty.")
Public Sub EmptyString(ByVal TEXT As Variant)
Attribute EmptyString.VB_Description = "Raises a run-time error if the specified string is empty."
    Dim errorDetails As TError
    With errorDetails
        .Trapped = (VarType(TEXT) <> vbString)
        If .Trapped Then
            .Name = "TypeMismatchErr"
            .Number = ErrNo.TypeMismatchErr
            .Source = TypeName(TEXT) & " type"
            .Message = "String required"
            .Description = "Variable type: " & TypeName(TEXT) & ". String required."
            RaiseError errorDetails
        End If
        
        .Trapped = (TEXT = vbNullString)
        If .Trapped Then
            .Name = "EmptyStringErr"
            .Number = ErrNo.EmptyStringErr
            .Source = "String variable"
            .Message = "String variable empty"
            .Description = .Message
            RaiseError errorDetails
        End If
    End With
End Sub


'@Description("Raises a run-time error if the specified object reference is Nothing.")
Public Sub NullReference(ByVal instanceVar As Variant, _
                Optional ByVal Source As String, _
                Optional ByVal Message As String)
Attribute NullReference.VB_Description = "Raises a run-time error if the specified object reference is Nothing."
    Dim errorDetails As TError
    With errorDetails
        .Trapped = Not IsObject(instanceVar)
        If .Trapped Then
            .Name = "ObjectRequiredErr"
            .Number = ErrNo.ObjectRequiredErr
            .Source = TypeName(instanceVar) & " Type"
            .Message = "Object required"
            .Description = "Variable type: " & TypeName(instanceVar) & ". Object required."
            RaiseError errorDetails
        End If
        
        .Trapped = instanceVar Is Nothing
        If .Trapped Then
            .Name = "ObjectNotSetErr"
            .Number = ErrNo.ObjectNotSetErr
            .Source = "Variable is Nothing"
            .Message = "Object variable not set"
            .Description = "Variable type: " & TypeName(instanceVar) & ". Object not set."
            RaiseError errorDetails
        End If
    End With
End Sub


'@Description("Raises a run-time error if the specified object reference is not Nothing.")
Public Sub NonNullReference(ByVal instanceVar As Variant, _
                   Optional ByVal Source As String, _
                   Optional ByVal Message As String)
Attribute NonNullReference.VB_Description = "Raises a run-time error if the specified object reference is not Nothing."
    Dim errorDetails As TError
    With errorDetails
        .Trapped = Not IsObject(instanceVar)
        If .Trapped Then
            .Name = "ObjectRequiredErr"
            .Number = ErrNo.ObjectRequiredErr
            .Source = TypeName(instanceVar) & " Type"
            .Message = "Object required"
            .Description = "Variable type: " & TypeName(instanceVar) & ". Object required."
            RaiseError errorDetails
        End If
        
        .Trapped = Not instanceVar Is Nothing
        If .Trapped Then
            .Name = "ObjectSetErr"
            .Number = ErrNo.ObjectSetErr
            .Source = TypeName(instanceVar) & " Type"
            .Message = "Object variable already set"
            .Description = "Variable type: " & TypeName(instanceVar) & ". Object is not Nothing."
            RaiseError errorDetails
        End If
    End With
End Sub


'@Description("Raises a run-time error if the specified instance isn't the default instance (singleton).")
Public Sub Singleton(ByVal instanceVar As Object)
Attribute Singleton.VB_Description = "Raises a run-time error if the specified instance isn't the default instance (singleton)."
    Debug.Assert IsObject(instanceVar)
    Guard.NullReference instanceVar
        
    Dim classVar As Object
    Set classVar = instanceVar.Class
    
    Dim errorDetails As TError
    With errorDetails
        .Trapped = Not instanceVar Is classVar
        If .Trapped Then
            .Name = "SingletonErr"
            .Number = ErrNo.SingletonErr
            .Source = TypeName(instanceVar) & " type"
            .Message = "Singleton already instantiated"
            .Description = "Creation of non-default instances is not supported. Use predeclared instance."
            RaiseError errorDetails
        End If
    End With
End Sub


'@Description("Raises a run-time error if the specified instance isn't the default instance.")
Public Sub NonDefaultInstance(ByVal instanceVar As Object)
Attribute NonDefaultInstance.VB_Description = "Raises a run-time error if the specified instance isn't the default instance."
    Guard.NullReference instanceVar
    
    Dim classVar As Object
    Set classVar = instanceVar.Class
    
    Dim errorDetails As TError
    With errorDetails
        .Trapped = Not instanceVar Is classVar
        If .Trapped Then
            .Name = "NonDefaultInstanceErr"
            .Number = ErrNo.NonDefaultInstanceErr
            .Source = TypeName(classVar) & " Class"
            .Message = "Default (" & TypeName(classVar) & ") object instance must be used"
            .Description = vbNullString
            RaiseError errorDetails
        End If
    End With
End Sub


'@Description("Raises a run-time error if the specified expression is false.")
Public Sub Expression(ByVal AssertIsTrue As Boolean, ByVal Source As String, ByVal Message As String)
Attribute Expression.VB_Description = "Raises a run-time error if the specified expression is false."
    Dim errorDetails As TError
    With errorDetails
        .Trapped = Not AssertIsTrue
        If .Trapped Then
            .Name = "CustomErr"
            .Number = ErrNo.CustomErr
            .Source = Source
            .Message = Message
            .Description = vbNullString
            RaiseError errorDetails
        End If
    End With
End Sub


'@Description("Raises a run-time error if the specified expression is false.")
Public Sub ExpressionErr(ByVal TrapIfFalse As Boolean, _
                         ByVal ErrorCode As ErrNo, _
                         ByVal Source As String, _
                         ByVal Message As String)
Attribute ExpressionErr.VB_Description = "Raises a run-time error if the specified expression is false."
    Dim errorDetails As TError
    With errorDetails
        .Trapped = Not TrapIfFalse
        If .Trapped Then
            .Name = "CustomErr"
            .Number = ErrorCode
            .Source = Source
            .Message = Message
            .Description = vbNullString
            RaiseError errorDetails
        End If
    End With
End Sub


'@Description("Raises a run-time error if the specified expression is true.")
Public Sub ExpressionTrueErr(ByVal TrapIfTrue As Boolean, _
                             ByVal ErrorCode As ErrNo, _
                             ByVal Source As String, _
                             ByVal Message As String)
Attribute ExpressionTrueErr.VB_Description = "Raises a run-time error if the specified expression is true."
    Dim errorDetails As TError
    With errorDetails
        .Trapped = TrapIfTrue
        If .Trapped Then
            .Name = "CustomErr"
            .Number = ErrorCode
            .Source = Source
            .Message = Message
            .Description = vbNullString
            RaiseError errorDetails
        End If
    End With
End Sub


'@Description("Raises a run-time InvalidCharcterErr if the specified expression is false.")
Public Sub InvalidCharacter(ByVal AssertIsTrue As Boolean, ByVal Source As String, ByVal Message As String)
Attribute InvalidCharacter.VB_Description = "Raises a run-time InvalidCharcterErr if the specified expression is false."
    Dim errorDetails As TError
    With errorDetails
        .Trapped = Not AssertIsTrue
        If .Trapped Then
            .Name = "InvalidCharacterErr"
            .Number = ErrNo.InvalidCharacterErr
            .Source = Source
            .Message = Message
            .Description = vbNullString
            RaiseError errorDetails
        End If
    End With
End Sub


'@Description("Raises a run-time ActionNotSupportedErr if the specified expression is false.")
Public Sub ActionNotSupported(ByVal AssertIsTrue As Boolean, ByVal Source As String, ByVal Message As String)
Attribute ActionNotSupported.VB_Description = "Raises a run-time ActionNotSupportedErr if the specified expression is false."
    Dim errorDetails As TError
    With errorDetails
        .Trapped = Not AssertIsTrue
        If .Trapped Then
            .Name = "ActionNotSupportedErr"
            .Number = ErrNo.ActionNotSupportedErr
            .Source = Source
            .Message = Message
            .Description = vbNullString
            RaiseError errorDetails
        End If
    End With
End Sub


'@Description("Raises a run-time TypeMismatchErr if the specified expression is false.")
Public Sub TypeMismatch(ByVal AssertIsTrue As Boolean, ByVal Source As String, ByVal Message As String)
Attribute TypeMismatch.VB_Description = "Raises a run-time TypeMismatchErr if the specified expression is false."
    Dim errorDetails As TError
    With errorDetails
        .Trapped = Not AssertIsTrue
        If .Trapped Then
            .Name = "TypeMismatchErr"
            .Number = ErrNo.TypeMismatchErr
            .Source = Source
            .Message = Message
            .Description = vbNullString
            RaiseError errorDetails
        End If
    End With
End Sub


'@Description("Raises a run-time error if the specified instance is the default instance.")
Friend Sub DefaultInstance(ByVal instanceVar As Object)
Attribute DefaultInstance.VB_Description = "Raises a run-time error if the specified instance is the default instance."
    Guard.NullReference instanceVar
    
    Dim classVar As Object
    Set classVar = instanceVar.Class
    
    Dim errorDetails As TError
    With errorDetails
        .Trapped = instanceVar Is classVar
        If .Trapped Then
            .Name = "DefaultInstanceErr"
            .Number = ErrNo.DefaultInstanceErr
            .Source = TypeName(classVar) & " Class"
            .Message = "Non-default (" & TypeName(classVar) & ") object instance must be used"
            .Description = vbNullString
            RaiseError errorDetails
        End If
    End With
End Sub


'@Description("Raises a run-time error if the specified file does not exist.")
Public Sub FileNotExist(ByVal FilePathName As String)
Attribute FileNotExist.VB_Description = "Raises a run-time error if the specified file does not exist."
    Dim fsObject As Scripting.FileSystemObject
    Set fsObject = New Scripting.FileSystemObject
        
    Dim errorDetails As TError
    With errorDetails
        .Trapped = Not fsObject.FileExists(FilePathName)
        If .Trapped Then
            .Name = "FileNotFoundErr"
            .Number = ErrNo.FileNotFoundErr
            .Source = FilePathName
            .Message = "File not found " & FilePathName
            .Description = vbNullString
            RaiseError errorDetails
        End If
    End With
End Sub


'@Description("Raises a run-time error if transaction is requested but not supported. (Trapped on True)")
Public Sub InvalidTransaction(ByVal TransactoinsDisabled As Boolean)
Attribute InvalidTransaction.VB_Description = "Raises a run-time error if transaction is requested but not supported. (Trapped on True)"
    Debug.Assert VarType(TransactoinsDisabled) = vbBoolean
    
    Dim errorDetails As TError
    With errorDetails
        .Trapped = TransactoinsDisabled
        If .Trapped Then
            .Name = "NotInTransaction"
            .Number = ErrNo.AdoInvalidTransactionErr
            .Source = vbNullString
            .Message = "Current database does not support transaction"
            .Description = "Attempted operation not permmited when transactions are not supported."
            RaiseError errorDetails
        End If
    End With
End Sub


'@Description("Raises a run-time error if the specified connection has active transaction. (Trapped on True)")
Public Sub InTransaction(ByVal HasActiveTransaction As Boolean)
Attribute InTransaction.VB_Description = "Raises a run-time error if the specified connection has active transaction. (Trapped on True)"
    Debug.Assert VarType(HasActiveTransaction) = vbBoolean
    
    Dim errorDetails As TError
    With errorDetails
        .Trapped = HasActiveTransaction
        If .Trapped Then
            .Name = "InTransaction"
            .Number = ErrNo.AdoInTransactionErr
            .Source = vbNullString
            .Message = "Current database connection already in transaction"
            .Description = "Attempted operation not permmited when transaction in progress (e.g. starting a new transaction)."
            RaiseError errorDetails
        End If
    End With
End Sub


'@Description("Raises a run-time error if the specified connection has no active transaction. (Trapped on False)")
Public Sub NotInTransaction(ByVal HasActiveTransaction As Boolean)
Attribute NotInTransaction.VB_Description = "Raises a run-time error if the specified connection has no active transaction. (Trapped on False)"
    Debug.Assert VarType(HasActiveTransaction) = vbBoolean
    
    Dim errorDetails As TError
    With errorDetails
        .Trapped = Not HasActiveTransaction
        If .Trapped Then
            .Name = "NotInTransaction"
            .Number = ErrNo.AdoInvalidTransactionErr
            .Source = vbNullString
            .Message = "Current database connection is not in transaction"
            .Description = "Attempted operation not permmited when no  active transaction in progress (e.g. commit or rollback)."
            RaiseError errorDetails
        End If
    End With
End Sub


'@Description("Raises a run-time error if the specified connection string is not valid or provider cannot be found. (Trapped on Nothing or vbNullString)")
Public Sub ConnectionString(ByVal dboOrConnStr As Variant)
Attribute ConnectionString.VB_Description = "Raises a run-time error if the specified connection string is not valid or provider cannot be found. (Trapped on Nothing or vbNullString)"
    Dim errorDetails As TError
    With errorDetails
        .Trapped = IsFalsy(dboOrConnStr)
        If .Trapped Then
            .Name = "ConnectionStringError"
            .Number = ErrNo.AdoConnectionStringErr
            .Source = "Argument is Nothing"
            .Message = "Either the connection string is invalid or provider not found"
            .Description = .Message
            RaiseError errorDetails
        End If
    End With
End Sub


'@Description("Raises a run-time error if the specified references are not both arrays having same base and size")
Public Sub IncompatibleArrays(ByVal ArrayA As Variant, ByVal ArrayB As Variant)
Attribute IncompatibleArrays.VB_Description = "Raises a run-time error if the specified references are not both arrays having same base and size"
    Dim errorDetails As TError
    With errorDetails
        .Trapped = Not (IsArray(ArrayA) And IsArray(ArrayB))
        If .Trapped Then
            .Name = "Expected array"
            .Number = ErrNo.ExpectedArrayErr
            .Source = "Bad arrays"
            .Message = "Expected two arrays"
            .Description = .Message
            RaiseError errorDetails
        End If
    End With
    
    With errorDetails
        .Trapped = (LBound(ArrayA) <> LBound(ArrayB))
        If .Trapped Then
            .Name = "Incompatible arrays"
            .Number = ErrNo.IncompatibleArraysErr
            .Source = "Bad arrays"
            .Message = "Provided arrays have different bases"
            .Description = .Message
            RaiseError errorDetails
        End If
    End With
    
    With errorDetails
        .Trapped = (UBound(ArrayA) <> UBound(ArrayB))
        If .Trapped Then
            .Name = "Incompatible arrays"
            .Number = ErrNo.IncompatibleArraysErr
            .Source = "Bad arrays"
            .Message = "Provided arrays have different sizes"
            .Description = .Message
            RaiseError errorDetails
        End If
    End With
End Sub


'@Description("Raises a run-time error if the specified references are not both arrays having same base and size")
Public Sub NotArray(ByVal ArrayArg As Variant)
Attribute NotArray.VB_Description = "Raises a run-time error if the specified references are not both arrays having same base and size"
    Dim errorDetails As TError
    With errorDetails
        .Trapped = Not IsArray(ArrayArg)
        If .Trapped Then
            .Name = "Expected array"
            .Number = ErrNo.ExpectedArrayErr
            .Source = "Scalar or object argument"
            .Message = "Expected array"
            .Description = .Message
            RaiseError errorDetails
        End If
    End With
End Sub


Friend Sub AssertExpectedError(ByVal Assert As Rubberduck.PermissiveAssertClass, _
                      Optional ByVal ExpectedErrorNo As ErrNo = ErrNo.PassedNoErr)
    Debug.Assert TypeOf Assert Is Rubberduck.PermissiveAssertClass
    
    Dim errorDetails As String
    errorDetails = " Error: #" & Err.Number & " - " & Err.Description
    
    Select Case Err.Number
        Case ExpectedErrorNo
            Assert.Succeed
        Case ErrNo.PassedNoErr
            Assert.Fail MsgExpectedErrNotRaised
        Case Else
            Assert.Fail MsgUnexpectedErrRaised & errorDetails
    End Select
    
    Err.Clear
End Sub


'@Ignore ProcedureNotUsed
'@Description("Re-raises the current error, if there is one.")
Friend Sub RethrowOnError()
Attribute RethrowOnError.VB_Description = "Re-raises the current error, if there is one."
    With VBA.Err
        If .Number <> 0 Then
            Debug.Print "Error " & .Number, .Description
            .Raise .Number, .Source, .Description
        End If
    End With
End Sub


'@Description("Formats and raises a run-time error.")
Friend Sub RaiseError(ByRef errorDetails As TError)
Attribute RaiseError.VB_Description = "Formats and raises a run-time error."
    With errorDetails
        Dim Message As Variant
        Message = Array("Error:", _
            "name: " & .Name, _
            "number: " & .Number, _
            "message: " & .Message, _
            "description: " & .Description, _
            "source: " & .Source)
        Logger.Logg Join(Message, vbNewLine & vbTab), , DEBUGLEVEL_ERROR
        VBA.Err.Raise .Number, .Source, .Message
    End With
End Sub
