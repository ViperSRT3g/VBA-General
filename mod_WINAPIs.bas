Attribute VB_Name = "mod_WINAPIs"
Option Explicit

Public Enum WindowStyle
    WS_BORDER = &H800000        'The window has a thin-line border.
    WS_CAPTION = &HC00000       'The window has a title bar (includes the WS_BORDER style).
    WS_DISABLED = &H8000000     'The window is initially disabled. A disabled window cannot receive input from the user. To change this after a window has been created, use the EnableWindow function.
    WS_MAXIMIZE = &H1000000     'The window is initially maximized.
    WS_MAXIMIZEBOX = &H10000    'The window has a maximize button. Cannot be combined with the WS_EX_CONTEXTHELP style. The WS_SYSMENU style must also be specified.
    WS_MINIMIZE = &H20000000    'The window is initially minimized.
    WS_MINIMIZEBOX = &H20000    'The window has a minimize button. Cannot be combined with the WS_EX_CONTEXTHELP style. The WS_SYSMENU style must also be specified.
    WS_SIZEBOX = &H40000        'The window has a sizing border.
    WS_SYSMENU = &H80000        'The window has a window menu on its title bar. The WS_CAPTION style must also be specified.
End Enum

Public Enum ExtendedWindowStyle
    WS_EX_APPWINDOW = &H40000       'Forces a top-level window onto the taskbar when the window is visible.
    WS_EX_TOOLWINDOW = &H80         'The window is intended to be used as a floating toolbar.
    WS_EX_TOPMOST = &H8             'The window should be placed above all non-topmost windows and should stay above them, even when the window is deactivated. To add or remove this style, use the SetWindowPos function.
    WS_EX_TRANSPARENT = &H20        'The window should not be painted until siblings beneath the window (that were created by the same thread) have been painted. The window appears transparent because the bits of underlying sibling windows have already been painted. To achieve transparency without these restrictions, use the SetWindowRgn function.
End Enum

Public Enum WindowPosition
    SWP_NOMOVE = &H2            'Retains the current position (ignores X and Y parameters).
    SWP_NOSIZE = &H1            'Retains the current size (ignores the cx and cy parameters).
    SWP_NOACTIVATE = &H10       'Does not activate the window. If this flag is not set, the window is activated and moved to the top of either the topmost or non-topmost group (depending on the setting of the hWndInsertAfter parameter).
    SWP_HIDEWINDOW = &H80       'Hides the window.
    SWP_SHOWWINDOW = &H40       'Displays the window.
    SWP_FRAMECHANGED = &H20     'Applies new frame styles set using the SetWindowLong function. Sends a WM_NCCALCSIZE message to the window, even if the window's size is not being changed. If this flag is not specified, WM_NCCALCSIZE is sent only when the window's size is being changed.
End Enum

Public Enum PlaySoundOptions
    SND_SYNC = &H0        '(Default) Play the sound synchronously. Code execution pauses until sound is complete.
    SND_ASYNC = &H1       'Play the sound asynchronously. Code execution does not wait for sound to complete.
    SND_NODEFAULT = &H2   'If the specified sound is not found, do not play the default sound (no sound is played).
    SND_MEMORY = &H4      'lpszSoundName is a memory file of the sound. Not used in VBA/VB6.
    SND_LOOP = &H8        'Continue playing sound in a loop until the next call to sndPlaySound.
    SND_NOSTOP = &H10     'Do not stop playing the current sound before playing the specified sound.
End Enum

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


#If Win64 And VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal Hwnd As Long) As Long
    Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal Hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal Hwnd As Long) As Long
#Else
    Private Declare Function DrawMenuBar Lib "user32" (ByVal Hwnd As Long) As Long
    Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Private Declare Function GetTickCount Lib "kernel32" () As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal Hwnd As Long) As Long
    Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal Hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare Function sndPlaySound32 Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
#End If

'Get Window Long Constants
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)

'Layered Window Attributes Constants
Private Const LWA_ALPHA = &H2&
Private Const LWA_COLORKEY = &H1


'Misc Constants
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOP = 0
Private Const ICON_SMALL = 0&
Private Const ICON_BIG = 1&
Private Const WM_SETICON = &H80
Private Const WS_EX_LAYERED = &H80000

'PUBLIC SUBROUTINES
Public Sub ModifyForm(ByVal WindowName As String, Optional FormResizable As Boolean, Optional EnableResizeButtons As Boolean = True, Optional ShowinTaskbar As Boolean = True, Optional IconFilePath As String)
    If Len(WindowName) = 0 Then Exit Sub
    Dim Hwnd As Long: Hwnd = GetHwnd(WindowName): If Hwnd = 0 Then Exit Sub
    If ValidIcon(IconFilePath) Then Call SetIcon(WindowName, IconFilePath)
    If EnableResizeButtons Then Call ResizeOptions(WindowName, FormResizable:=FormResizable)
    If FormResizable Then Call SetWinStyle(WindowName, WS_SIZEBOX)
    If ShowinTaskbar Then Call DisplayInTaskbar(WindowName)
End Sub


'PUBLIC FUNCTIONS
Public Function GetHwnd(ByVal WindowName As String) As Long
    If Len(WindowName) = 0 Then Exit Function
    GetHwnd = FindWindow("ThunderDFrame", WindowName)
End Function

Public Function SetIcon(ByVal WindowName As String, ByVal IconPath As String, Optional LargeIcon As Boolean = True) As Boolean
    If Len(IconPath) = 0 Then Exit Function
    If Not (FileExists(IconPath) And ValidIcon(IconPath)) Then Exit Function
    'If WindowName is passed, use that. Otherwise default to internal hwnd, if neither are valid Exit Function
    Dim Hwnd As Long: Hwnd = GetHwnd(WindowName): If Hwnd = 0 Then Exit Function
    
    Dim IconID As Long: IconID = ExtractIcon(0, IconPath, 0)
    SetIcon = SendMessage(Hwnd, WM_SETICON, IIf(LargeIcon, ICON_BIG, ICON_SMALL), IconID) > 0
End Function

Public Function SetWinStyle(ByVal WindowName As String, Optional ByVal WStyle As WindowStyle)
    Dim Hwnd As Long: Hwnd = GetHwnd(WindowName): If Hwnd = 0 Then Exit Function
    SetWinStyle = SetWindowLong(Hwnd, GWL_STYLE, GetWindowLong(Hwnd, GWL_STYLE) Or WStyle)
    Call DrawMenuBar(Hwnd) 'Redraws the menu bar of the specified window. If the menu bar changes after the system has created the window, this function must be called to draw the changed menu bar.
End Function

Public Function DisplayInTaskbar(ByVal WindowName As String) As Boolean
    Dim Hwnd As Long: Hwnd = GetHwnd(WindowName): If Hwnd = 0 Then Exit Function
    DisplayInTaskbar = SetWindowPos(Hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_HIDEWINDOW) And _
                       SetWindowLong(Hwnd, GWL_EXSTYLE, GetWindowLong(Hwnd, GWL_EXSTYLE) Or WS_SIZEBOX) And _
                       SetWindowPos(Hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_SHOWWINDOW)
    Call DrawMenuBar(Hwnd)
End Function

Public Function ResizeOptions(ByVal WindowName As String, Optional Minimize As Boolean = True, Optional Maximize As Boolean = True, Optional FormResizable As Boolean) As Boolean
    Dim Hwnd As Long: Hwnd = GetHwnd(WindowName): If Hwnd = 0 Then Exit Function
    Dim ROptions As Long
    If Minimize Then ROptions = ROptions Or WS_MINIMIZEBOX
    If Maximize Then ROptions = ROptions Or WS_MAXIMIZEBOX
    If FormResizable Then ROptions = ROptions Or WS_SIZEBOX
    ResizeOptions = SetWindowLong(Hwnd, GWL_STYLE, GetWindowLong(Hwnd, GWL_STYLE) Or ROptions)
    Call DrawMenuBar(Hwnd)
End Function

Public Function SetTransparency(ByVal WindowName As String, R As Integer, G As Integer, B As Integer) As Long
    Dim Hwnd As Long: Hwnd = GetHwnd(WindowName): If Hwnd = 0 Then Exit Function
    Dim Color As Long: Color = B * 65536 + G * 256 + R
    Dim RetVal As Long: RetVal = RetVal Or WS_EX_LAYERED
    If SetWindowLong(Hwnd, GWL_EXSTYLE, RetVal) Then SetTransparency = CBool(SetLayeredWindowAttributes(Hwnd, Color, 0, LWA_COLORKEY))
End Function

Public Function SetTranslucency(ByVal WindowName As String, ByVal Alpha As Byte) As Boolean
    Dim Hwnd As Long: Hwnd = GetHwnd(WindowName): If Hwnd = 0 Then Exit Function
    Dim lStyle As Long: lStyle = GetWindowLong(Hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    If SetWindowLong(Hwnd, GWL_EXSTYLE, lStyle) Then SetTranslucency = CBool(SetLayeredWindowAttributes(Hwnd, 0, CLng(Alpha), LWA_ALPHA))
End Function

Public Function SetTopmost(ByVal WindowName As String, ByVal Toggle As Boolean) As Boolean
    Dim Hwnd As Long: Hwnd = GetHwnd(WindowName): If Hwnd = 0 Then Exit Function
    SetTopmost = SetWindowPos(Hwnd, IIf(Toggle, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_FLAGS)
End Function

Public Function Intersect(ByRef Rect1 As RECT, Rect2 As RECT) As Boolean
    Dim ResultRect As RECT: Intersect = IntersectRect(ResultRect, Rect1, Rect2)
End Function


'PRIVATE FUNCTIONS
Private Function FileExists(ByVal FilePath As String) As Boolean
    With CreateObject("Scripting.FileSystemObject"): FileExists = .FileExists(FilePath): End With: End Function
Private Function FileExt(ByVal FilePath As String) As String
    With CreateObject("Scripting.FileSystemObject"): FileExt = .GetExtensionName(FilePath): End With: End Function
Private Function ValidIcon(ByVal IconPath As String) As Boolean
    ValidIcon = FileExt(IconPath) = "ico": End Function
