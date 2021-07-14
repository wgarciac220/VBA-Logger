Option Explicit

' @implements interface
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
    ILogger_Name = "file"
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
Private Function AppendData(message As String, File As String)
    Dim FileNum As Integer
    
        FileNum = FreeFile
        
        Open File For Append As #FileNum
        Print #FileNum, message
        Close #FileNum

End Function

Private Function FindLogger(coll As Collection) As ILogger
    Dim Item As ILogger
    
        For Each Item In coll
            If Item.Name = Me.ILogger_Name Then
                Set FindLogger = Item
                Exit For
            End If
        Next Item
    
End Function

'   customize
Private Function StringFormat(ParamArray arr() As Variant) As String
    Dim i As Long
    Dim temp As String

        temp = CStr(arr(0))
        
        For i = 1 To UBound(arr)
            temp = Replace(temp, "{" & i - 1 & "}", CStr(arr(i)))
        Next
        
        StringFormat = temp
        
End Function

Private Function BytesToMegabytes(Bytes As Double) As Double
    Dim size As Double
    
        size = (Bytes / 1024) / 1024
        BytesToMegabytes = Format(size, "###,###,##0.00")
  
End Function

Private Function ValidateFileSize(File As String) As String
    Dim temp As String
    Dim parent As String
    
        temp = StringFormat("Log_{0}.txt", Format(Now, "mmddyyyy_hhmmss"))
        
        If Path.FileExists(File) Then
            If BytesToMegabytes(Path.GetFile(File).size) > 10 Then
                parent = Path.GetFile(File).ParentFolder
                Path.GetFile(File).Name = temp
                File = Path.Combine(parent, "Log.txt")
            End If
        End If

        ValidateFileSize = File
        
End Function

' @public methods
Public Function ILogger_SetNext(handler As ILogger) As ILogger
    Set mNextHandler = handler
    Set ILogger_SetNext = handler
End Function

Public Sub ILogger_LogMessage()
    Dim temp As String
    Dim File As String
    Dim Logger As ILogger
    
        Set Logger = FindLogger(mConfig.Loggers)
        
        If Not Logger Is Nothing Then
            If mConfig.Level >= Logger.RequiredLevel Then
                temp = "File::Logger:"
                temp = temp & mConfig.message
                
                File = Path.Combine(ThisWorkbook.Path, "Log.txt")
                File = ValidateFileSize(File)
        
                AppendData temp, File
            End If
        End If
        
        If Not mNextHandler Is Nothing Then
            mNextHandler.Bindings = mConfig
            mNextHandler.LogMessage
        End If
        
End Sub

