Attribute VB_Name = "time"
Option Explicit

    Dim startTime, endTime As Double
    Dim i As Long
    
Sub measureTime()
    startTime = Timer
    
    For i = 0 To 100000000
    'ŠÔ‚ª‚©‚©‚éˆ—
    Next i
    
    endTime = Timer
    
    Debug.Print "startTimeF" & startTime & "•b"
    Debug.Print "EndTimeF" & endTime & "•b"
    Debug.Print "End - StartF" & endTime - startTime & "•b"
    
End Sub

Sub getTime()
    startTime = Now
    Debug.Print Format(Now, "hh:mm:ss")
    Debug.Print Format(Now, "Medium Time")
    
    For i = 0 To 300000000
    'ŠÔ‚ª‚©‚©‚éˆ—
    Next i
    
    endTime = Now
    
    Debug.Print "End - StartF" & endTime - startTime & "•b"
    
End Sub

