# VBA-Logger

A logger for VBA, based on the desired level, logs data into the immediate window, a txt file, or in a form of a messagebox. 
This project uses Chain of Responsibility design pattern.

```VBA
Public Sub Main()
    
    Logger.Init
    Logger.DEBG "Debug message"
    Logger.INFO "Info message"
    Logger.WARN "Warning message"
    Logger.ERRR "Error message"
    Logger.FATL "Fatal message"
        
End Sub
```
