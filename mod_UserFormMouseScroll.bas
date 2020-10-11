Attribute VB_Name = "mod_UserFormMouseScroll"
Option Explicit
 ' Based on code from Peter Thornton here:
 ' http://social.msdn.microsoft.com/Forums/en-US/7d584120-a929-4e7c-9ec2-9998ac639bea/mouse-scroll-in-userform-listbox-in-excel-2010?forum=isvvba
Private Type POINTAPI
    x                               As Long
    y                               As Long
End Type
Private Type MOUSEHOOKSTRUCT
    pt                              As POINTAPI
    hwnd                            As Long
    wHitTestCode                    As Long
    dwExtraInfo                     As Long
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
Declare Function GetActiveWindow Lib "user32" () As Long
 
Private Const WH_MOUSE_LL          As Long = 14
Private Const WM_MOUSEWHEEL        As Long = &H20A
Private Const HC_ACTION            As Long = 0
Private Const GWL_HINSTANCE        As Long = (-6)
 
Private Const WM_KEYDOWN           As Long = &H100
Private Const WM_KEYUP             As Long = &H101
Private Const VK_UP                As Long = &H26
Private Const VK_DOWN              As Long = &H28
Private Const WM_LBUTTONDOWN       As Long = &H201
 
Private Const cSCROLLCHANGE        As Long = 10
 
Private mLngMouseHook              As Long
Private mFormHwnd                  As Long
Private mbHook                     As Boolean
Dim mForm                          As Object
 
 
Public Sub HookFormScroll(oForm As Object)
    Dim lngAppInst                  As Long
    Dim hwndUnderCursor             As Long
     
    Set mForm = oForm
    hwndUnderCursor = FindWindow("ThunderDFrame", oForm.Caption)
    Debug.Print "Form window: " & hwndUnderCursor
    If mFormHwnd <> hwndUnderCursor Then
        UnhookFormScroll
        Debug.Print "Unhook old proc"
        mFormHwnd = hwndUnderCursor
        lngAppInst = GetWindowLong(mFormHwnd, GWL_HINSTANCE)
        If Not mbHook Then
            mLngMouseHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseProc, lngAppInst, 0)
            mbHook = mLngMouseHook <> 0
            If mbHook Then Debug.Print "Form hooked"
        End If
    End If
End Sub
 
Public Sub UnhookFormScroll()
    If mbHook Then
        UnhookWindowsHookEx mLngMouseHook
        mLngMouseHook = 0
        mFormHwnd = 0
        mbHook = False
    End If
End Sub
 
Private Function MouseProc(ByVal nCode As Long, ByVal wParam As Long, ByRef lParam As MOUSEHOOKSTRUCT) As Long
    On Error GoTo errH 'Resume Next
    If (nCode = HC_ACTION) Then
        If GetActiveWindow = mFormHwnd Then
             
            If wParam = WM_MOUSEWHEEL Then
                MouseProc = True
                If lParam.hwnd > 0 Then
                    mForm.ScrollTop = Application.Max(0, mForm.ScrollTop - cSCROLLCHANGE)
                Else
                    mForm.ScrollTop = Application.Min(mForm.ScrollHeight - mForm.InsideHeight, mForm.ScrollTop + cSCROLLCHANGE)
                End If
                Exit Function
            End If
        End If
         
    End If
    MouseProc = CallNextHookEx(mLngMouseHook, nCode, wParam, ByVal lParam)
    Exit Function
errH:
    UnhookFormScroll
End Function

