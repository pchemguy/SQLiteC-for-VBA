VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Logging facility using a dictionary object as the database."
'@Folder "Common.Logger"
'@ModuleDescription "Logging facility using a dictionary object as the database."
'@PredeclaredId
'@Exposed
'@IgnoreModule ProcedureNotUsed
Option Explicit

Public Enum DebugLevelEnum
    DEBUGLEVEL_NONE = -1&
    DEBUGLEVEL_ERROR = 0&
    DEBUGLEVEL_WARN = 1&
    DEBUGLEVEL_INFO = 2&
    DEBUGLEVEL_VERB = 3&
    DEBUGLEVEL_MAX = &H7FFFFFFF
End Enum

Private Type TLogger
    LogDb As Scripting.Dictionary
    Timers As Scripting.Dictionary
    '''' Only use if message's debug level is not greater
    DebugLevelImmediate As DebugLevelEnum
    '''' Only use if message's debug level is not greater
    DebugLevelDatabase As DebugLevelEnum
    '''' Use to set id width
    RecordIdDigits As Long
    UseTimeStamp As Boolean
    UseIdPadding As Boolean
    LogType(0 To 3) As String * 4
End Type
Private this As TLogger


Private Sub Class_Initialize()
    With this
        Set .LogDb = New Scripting.Dictionary
        .LogDb.CompareMode = TextCompare
        Set .Timers = New Scripting.Dictionary
        .Timers.CompareMode = TextCompare
        .DebugLevelImmediate = DEBUGLEVEL_ERROR
        .DebugLevelDatabase = DEBUGLEVEL_MAX
        .RecordIdDigits = 4
        .UseIdPadding = False
        .UseTimeStamp = False
        .LogType(0) = "ERRO"
        .LogType(1) = "WARN"
        .LogType(2) = "INFO"
        .LogType(3) = "VERB"
    End With
End Sub

Public Property Get DebugLevelImmediate() As DebugLevelEnum
    DebugLevelImmediate = this.DebugLevelImmediate
End Property

Public Property Let DebugLevelImmediate(ByVal Value As DebugLevelEnum)
    this.DebugLevelImmediate = Value
End Property

Public Property Get DebugLevelDatabase() As DebugLevelEnum
    DebugLevelDatabase = this.DebugLevelDatabase
End Property

Public Property Let DebugLevelDatabase(ByVal Value As DebugLevelEnum)
    this.DebugLevelDatabase = Value
End Property

Public Property Set LogDatabase(ByVal Instance As Scripting.Dictionary)
    Set this.LogDb = Instance
End Property

Public Property Get LogDatabase() As Scripting.Dictionary
    Set LogDatabase = this.LogDb
End Property

Public Property Get UseIdPadding() As Boolean
    UseIdPadding = this.UseIdPadding
End Property

Public Property Let UseIdPadding(ByVal Value As Boolean)
    this.UseIdPadding = Value
End Property

Public Property Get UseTimeStamp() As Boolean
    UseTimeStamp = this.UseTimeStamp
End Property

Public Property Let UseTimeStamp(ByVal Value As Boolean)
    this.UseTimeStamp = Value
End Property

Friend Sub RecordIdDigits(ByVal Value As Long)
    this.RecordIdDigits = Value
End Sub

'@Description "Adds item to the log database"
Public Sub Logg(ByVal LogLine As String, _
      Optional ByRef LogDatabase As Scripting.Dictionary = Nothing, _
      Optional ByVal DebugLevel As DebugLevelEnum = DEBUGLEVEL_ERROR)
Attribute Logg.VB_Description = "Adds item to the log database"
    Dim UseDb As Boolean
    UseDb = (DebugLevel <= this.DebugLevelDatabase)
    Dim UseImmediate As Boolean
    UseImmediate = (DebugLevel <= this.DebugLevelImmediate)
    If UseDb Or UseImmediate Then
        Dim LogDb As Scripting.Dictionary
        Set LogDb = IIf(LogDatabase Is Nothing, this.LogDb, LogDatabase)
        Dim RecId As String
        RecId = CStr(LogDb.Count)
        Dim Padding As Long
        If this.UseIdPadding Then
            Padding = this.RecordIdDigits - Len(RecId)
        Else
            Padding = 0
        End If
        Dim TimeStamp As String
        TimeStamp = vbNullString
        If this.UseTimeStamp Then TimeStamp = GetTimeStampMs & " "
        RecId = TimeStamp & "#" & String(IIf(Padding > 0, Padding, 0), " ") & _
                RecId & " " & this.LogType(DebugLevel) & "#"
        Dim Message As String
        Message = RecId & vbTab & LogLine
    End If
    '@Ignore IndexedDefaultMemberAccess
    If UseDb Then LogDb(RecId) = Message
    If UseImmediate Then Debug.Print Message
End Sub

'@Description "Prints the log database to the immediate"
Public Sub PrintLog(Optional ByRef LogDatabase As Scripting.Dictionary = Nothing)
Attribute PrintLog.VB_Description = "Prints the log database to the immediate"
    Dim LogDb As Scripting.Dictionary
    Set LogDb = IIf(LogDatabase Is Nothing, this.LogDb, LogDatabase)
    If LogDb.Count > 0 Then Debug.Print Join(LogDb.Items, vbNewLine)
End Sub

Public Sub ClearLog(Optional ByRef LogDatabase As Scripting.Dictionary = Nothing)
    IIf(LogDatabase Is Nothing, this.LogDb, LogDatabase).RemoveAll
End Sub

'@Description "Starts named timer. Use target name [module, macro, etc.] as TimerId."
Public Sub TimerSet(ByVal TimerId As String)
Attribute TimerSet.VB_Description = "Starts named timer. Use target name [module, macro, etc.] as TimerId."
    With this.Timers
        Guard.ExpressionTrueErr .Exists(TimerId), InvalidParameterErr, _
                                "Logger/TimerSet", "Timer ID already exists."
        .Item(TimerId) = Timer()
    End With
End Sub

'@Description "Logs elapsed time and removes named timer."
Public Sub TimerLogClear(ByVal TimerId As String, _
                Optional ByVal ObjectCount As Long = 1)
Attribute TimerLogClear.VB_Description = "Logs elapsed time and removes named timer."
    With this.Timers
        Guard.ExpressionErr .Exists(TimerId), InvalidParameterErr, _
                            "Logger/TimerSet", "Timer ID does not exist."
        Dim Elapsed As Single
        Elapsed = Round((Timer() - .Item(TimerId)) * 1000, 0)
        Logg "===== TIMERID: " & TimerId, , DEBUGLEVEL_INFO
        Logg "===== ELAPSED: " & CStr(Elapsed) & " ms.", , DEBUGLEVEL_INFO
        If ObjectCount > 1 Then
            Logg "===== OBJECTS: " & CStr(ObjectCount), , DEBUGLEVEL_INFO
            Logg "===== AVERAGE: " & CStr(Round(Elapsed / ObjectCount, 0)) & _
                 " ms/object", , DEBUGLEVEL_INFO
        End If
        .Remove TimerId
    End With
End Sub
