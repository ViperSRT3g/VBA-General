Attribute VB_Name = "mod_Maximize"
Option Explicit

Public Enum WindowEnum
    LeftMonitor = 0
    RightMonitor = 1
End Enum

Public Function MaximizeApp(ByVal Monitor As WindowEnum) As Boolean
    With Application
        .WindowState = xlNormal
        .Left = 2000 * Monitor
        .WindowState = xlMaximized
        MaximizeApp = (.WindowState = xlMaximized)
    End With
End Function
