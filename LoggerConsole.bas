Option Explicit

' @implements
Implements ILogger

' @private vars
Private mNextHandler As ILogger
Private mConfig As New LoggerConfig
Private mLevelToLog As Integer

' @properties
Property Get ILogger_Self() As ILogger
    Set ILogger_Self = Me
End Property

Property Get ILogger_Name() As String
    ILogger_Name = "console"
End Property

Property Let ILogger_RequiredLevel(value As Integer)
    mLevelToLog = value
End Property

Property Let ILogger_Bindings(binding As LoggerConfig)
    Set mConfig = binding
End Property

Property Get ILogger_RequiredLevel() As Integer
    ILogger_RequiredLevel = mLevelToLog
End Property

' @private methods
Private Function FindLogger(coll As Collection) As ILogger
    Dim Item As ILogger
    
        For Each Item In coll
            If Item.Name = Me.ILogger_Name Then
                Set FindLogger = Item
                Exit For
            End If
        Next Item
    
End Function

' @public methods
Public Function ILogger_SetNext(handler As ILogger) As ILogger
    Set mNextHandler = handler
    Set ILogger_SetNext = handler
End Function

Public Sub ILogger_LogMessage()
    Dim temp As String
    Dim Logger As ILogger
    
        Set Logger = FindLogger(mConfig.Loggers)
        
        If Not Logger Is Nothing Then
            If mConfig.Level >= Logger.RequiredLevel Then
                temp = "CONSOLE::Logger:"
                temp = temp & mConfig.message
                
                Debug.Print temp
            End If
        End If

        If Not mNextHandler Is Nothing Then
            mNextHandler.Bindings = mConfig
            mNextHandler.LogMessage
        End If
        
End Sub





