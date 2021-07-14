Option Explicit

' @implements interface
Implements ILoggerLayout

' @enum
Private Enum Levels
    OFF = 0
    DEBGl = 1
    INFOl = 2
    WARNl = 3
    ERRRl = 4
    FATLl = 5
End Enum

' @private vars
Private mLogger As ILogger
Private mLevel As Levels
Private Config As LoggerConfig
Private mNewSession As Boolean

'   -- layout
Private mIncludeDate As Boolean
Private mIncludeTime As Boolean
Private mIncludeLevel As Boolean

' @properties
'   -- layout
Property Let ILoggerLayout_AddDate(add As Boolean)
    mIncludeDate = add
End Property

Property Let ILoggerLayout_AddTime(add As Boolean)
    mIncludeTime = add
End Property

Property Let ILoggerLayout_AddLevel(add As Boolean)
    mIncludeLevel = add
End Property

' @private methods
Public Sub Init()
    Dim Console As ILogger
    Dim MessageBox As ILogger
    Dim File As ILogger
    Dim LoggersList As Collection
    
        Set Config = New LoggerConfig
        Set LoggersList = New Collection
        
        ' create logger instance and assign default level
        Set Console = New LoggerConsole
            Console.RequiredLevel = Levels.INFOl
        
        Set MessageBox = New LoggerMessageBox
            MessageBox.RequiredLevel = Levels.ERRRl
        
        Set File = New LoggerFile
            File.RequiredLevel = Levels.DEBGl
            
                
        ' add loggers to list
        LoggersList.add Console
        LoggersList.add MessageBox
        LoggersList.add File
        
        ' map handlers for each logger
        MapHandlers LoggersList

        ' add only required loggers to the collection
        Config.AddLogger Console
        Config.AddLogger MessageBox
        Config.AddLogger File
                        
        Set mLogger = Console

End Sub

Private Sub MapHandlers(Loggers As Collection)
    Dim Item As ILogger
    Dim Previous As ILogger
    
        For Each Item In Loggers
            If Not Previous Is Nothing Then
                Previous.SetNext Item
            End If
            
            Set Previous = Item
        Next Item
End Sub

Private Sub LoggerConfiguration()
        
        mIncludeDate = True
        mIncludeTime = True
        mIncludeLevel = True
    
End Sub

Private Function EnumName(Level As Levels) As String
    EnumName = Array("OFF", "DEBG", "INFO", "WARN", "ERRR", "FATL")(Level)
End Function

Private Function StringFormat(ParamArray arr() As Variant) As String
    Dim i As Long
    Dim temp As String

        temp = CStr(arr(0))
        
        For i = 1 To UBound(arr)
            temp = Replace(temp, "{" & i - 1 & "}", CStr(arr(i)))
        Next
        
        StringFormat = temp
        
End Function

Private Sub LogMessage(message As String)
    Dim temp As String
    Dim Level As String
    
        LoggerConfiguration

        Level = EnumName(mLevel)
    
        temp = temp & IIf(mIncludeLevel, StringFormat(" {0} -", Level), "")
        temp = temp & IIf(mIncludeDate, StringFormat(" {0} -", Format(Date, "mm/dd/yyyy")), "")
        temp = temp & IIf(mIncludeTime, StringFormat(" {0} -", Format(Time, "hh:mm:ss AM/PM")), "")
        temp = temp & StringFormat(" {0}", message)
        
        Config.Level = mLevel
        Config.LevelName = Level
        Config.message = temp
        
        mLogger.Bindings = Config
        mLogger.LogMessage
        
        mNewSession = False
        
End Sub

' @public methods
Public Sub DEBG(message As String)
    mLevel = DEBGl
    LogMessage message
End Sub

Public Sub INFO(message As String)
    mLevel = INFOl
    LogMessage message
End Sub

Public Sub WARN(message As String)
    mLevel = WARNl
    LogMessage message
End Sub

Public Sub ERRR(message As String)
    mLevel = ERRRl
    LogMessage message
End Sub

Public Sub FATL(message As String)
    mLevel = FATLl
    LogMessage message
End Sub

' @constructors
Private Sub Class_Initialize()
End Sub

