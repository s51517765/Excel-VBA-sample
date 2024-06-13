Attribute VB_Name = "time"
Option Explicit

    Dim startTime, endTime As Double
    Dim i As Long
    
Sub measureTime()
    startTime = Timer
    
    For i = 0 To 100000000
    '時間がかかる処理
    Next i
    
    endTime = Timer
    
    Debug.Print "startTime：" & startTime & "秒"
    Debug.Print "EndTime：" & endTime & "秒"
    Debug.Print "End - Start：" & endTime - startTime & "秒"
    
End Sub

Sub getTime()
    startTime = Now
    Debug.Print Format(Now, "hh:mm:ss")
    Debug.Print Format(Now, "Medium Time")
    
    For i = 0 To 300000000
    '時間がかかる処理
    Next i
    
    endTime = Now
    
    Debug.Print "End - Start：" & endTime - startTime & "秒"
    
End Sub

