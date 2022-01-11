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

#If Win64 And VBA7 Then
    Private Declare PtrSafe Function GetTickCount64 Lib "kernel32" () As LongPtr
    Private Previous As LongPtr
    Private PausedTime As LongPtr
#Else
    Private Declare Function GetTickCount Lib "kernel32" () As Long
    Private Previous As Long
    Private PausedTime As Long
#End If

Private StartTime As Date
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
        Previous = GetTicks
        Paused = True
    End If
End Sub

Public Sub UnPause()
    If Paused Then
        PausedTime = PausedTime + Abs(GetTicks - Previous)
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

#If Win64 And VBA7 Then
    Private Function Ticks2Time(ByVal Ticks As LongPtr) As Date
#Else
    Private Function Ticks2Time(ByVal Ticks As Long) As Date
#End If
    'Converts ticks to the number of elapsed seconds
    Ticks2Time = (Ticks / 1000) * CDbl(TimeValue("00:00:01"))
End Function

#If Win64 And VBA7 Then
    Private Function GetTicks() As LongPtr
        GetTicks = GetTickCount64
#Else
    Private Function GetTicks() As Long
        GetTicks = GetTickCount
#End If
End Function
