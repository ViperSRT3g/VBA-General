VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "class_Timer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private StartTime As Date
Private Previous As Long
Private PausedTime As Long
Private Paused As Boolean

Public Function Start()
    'Returns the current date/time that this timer was started
    StartTime = Now()
    Previous = 0
    PausedTime = 0
    Paused = False
    Start = StartTime
End Function

Public Sub Pause()
    If Not Paused Then
        Previous = GetTickCount
        Paused = True
    End If
End Sub

Public Sub UnPause()
    If Paused Then
        PausedTime = PausedTime + Abs(GetTickCount - Previous)
        Paused = False
    End If
End Sub

Public Function StopTimer() As Date
    'Always returns the elapsed time from when the timer began without
    'including the amount of time this timer may have been paused!
    'Calling this function more than once yields the current elapsed time
    'since the timer began! This can be used as an interval recording function!
    StopTimer = Abs((Now() - StartTime) - Ticks2Time(PausedTime))
End Function

Private Function Ticks2Time(ByVal Ticks As Long) As Date
    'Converts ticks to the number of elapsed seconds
    Ticks2Time = (Ticks / 1000) * CDbl(TimeValue("00:00:01"))
End Function
