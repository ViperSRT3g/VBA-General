Attribute VB_Name = "mod_UserformAPIs"
Option Explicit

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Function SetAsForeground(ByRef TargetUserform As UserForm) As Boolean
    On Error GoTo ErrorHandler
    If Len(TargetUserform.Caption) = 0 Then Exit Function
    Dim XLHWnd As Long: XLHWnd = Application.hwnd
    Dim MeHWnd As Long: MeHWnd = FindWindow("ThunderDFrame", TargetUserform.Caption)
    If (MeHWnd > 0) And (XLHWnd > 0) Then
        If SetParent(MeHWnd, XLHWnd) > 0 And SetForegroundWindow(MeHWnd) > 0 Then
            SetAsForeground = True
        End If
    End If
ErrorHandler:
End Function

