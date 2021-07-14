Option Explicit

' private vars
Private mLevel As Integer
Private mLevelName As String
Private mLoggers As Collection
Private mMessage As String

' @properties
Property Let Level(value As Integer)
    mLevel = value
End Property

Property Let LevelName(Name As String)
    mLevelName = Name
End Property

Property Let message(value As String)
    mMessage = value
End Property

Property Get Level() As Integer
    Level = mLevel
End Property

Property Get LevelName() As String
    LevelName = mLevelName
End Property

Property Get message() As String
    message = mMessage
End Property

Property Get Loggers() As Collection
    Set Loggers = mLoggers
End Property

' @public methods
Public Sub AddLogger(Logger As ILogger)
    mLoggers.add Logger
End Sub

' @constructors

Private Sub Class_Initialize()
    Set mLoggers = New Collection
End Sub

Private Sub Class_Terminate()
    Set mLoggers = Nothing
End Sub
