Attribute VB_Name = "mod_ListboxMouseScroll"
Option Explicit

Private Type POINTAPI
     x As Long
     y As Long
End Type

Private Type MOUSEHOOKSTRUCT
     pt As POINTAPI
     hwnd As Long
     wHitTestCode As Long
     dwExtraInfo As Long
End Type

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long

Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" ( _
    ByVal idHook As Long, _
    ByVal lpfn As Long, _
    ByVal hmod As Long, _
    ByVal dwThreadId As Long) As Long

Private Declare Function CallNextHookEx Lib "user32" ( _
    ByVal hHook As Long, _
    ByVal nCode As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Private Declare Function UnhookWindowsHookEx Lib "user32" ( _
    ByVal hHook As Long) As Long

Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Private Declare Function WindowFromPoint Lib "user32" ( _
    ByVal xPoint As Long, _
    ByVal yPoint As Long) As Long

Private Declare Function GetCursorPos Lib "user32.dll" ( _
    ByRef lpPoint As POINTAPI) As Long

Private Const WH_MOUSE_LL As Long = 14
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const HC_ACTION As Long = 0
Private Const GWL_HINSTANCE As Long = (-6)

Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const VK_UP As Long = &H26
Private Const VK_DOWN As Long = &H28
Private Const WM_LBUTTONDOWN As Long = &H201

Private mLngMouseHook As Long
Private mListBoxHwnd As Long
Private mbHook As Boolean

Public Sub HookListBoxScroll()
    Dim lngAppInst As Long
    Dim hwndUnderCursor As Long
    Dim tPT As POINTAPI
    
    Call GetCursorPos(tPT)
    hwndUnderCursor = WindowFromPoint(tPT.x, tPT.y)
    If mListBoxHwnd <> hwndUnderCursor Then
        Call UnhookListBoxScroll
        mListBoxHwnd = hwndUnderCursor
        lngAppInst = GetWindowLong(mListBoxHwnd, GWL_HINSTANCE)
        Call PostMessage(mListBoxHwnd, WM_LBUTTONDOWN, 0&, 0&)
        If Not mbHook Then
            mLngMouseHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseProc, lngAppInst, 0)
            mbHook = mLngMouseHook <> 0
        End If
     End If
End Sub

Public Sub UnhookListBoxScroll()
    If mbHook Then
        Call UnhookWindowsHookEx(mLngMouseHook)
        mLngMouseHook = 0
        mListBoxHwnd = 0
        mbHook = False
    End If
End Sub

Private Function MouseProc(ByVal nCode As Long, ByVal wParam As Long, ByRef lParam As MOUSEHOOKSTRUCT) As Long
    On Error GoTo errH 'Resume Next
    If (nCode = HC_ACTION) Then
        If WindowFromPoint(lParam.pt.x, lParam.pt.y) = mListBoxHwnd Then
            If wParam = WM_MOUSEWHEEL Then
                MouseProc = True
                If lParam.hwnd > 0 Then
                    Call PostMessage(mListBoxHwnd, WM_KEYDOWN, VK_UP, 0)
                Else
                    Call PostMessage(mListBoxHwnd, WM_KEYDOWN, VK_DOWN, 0)
                End If
                Call PostMessage(mListBoxHwnd, WM_KEYUP, VK_UP, 0)
                Exit Function
            End If
        Else
            Call UnhookListBoxScroll
        End If
    End If
    MouseProc = CallNextHookEx(mLngMouseHook, nCode, wParam, ByVal lParam)
    Exit Function
errH:
    Call UnhookListBoxScroll
End Function

