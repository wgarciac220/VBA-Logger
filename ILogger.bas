Option Explicit

' @properties
Property Get Self() As ILogger
End Property

Property Get Name() As String
End Property

Property Let RequiredLevel(value As Integer)
End Property

Property Let Bindings(binding As LoggerConfig)
End Property

Property Get RequiredLevel() As Integer
End Property

' @public methods
Public Function SetNext(handler As ILogger) As ILogger
End Function

Public Sub LogMessage()
End Sub


