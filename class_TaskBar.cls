VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "class_TaskBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

#If Win64 And VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal Hwnd As Long) As Long
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
#End If

'Get Window Long Constants
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)

'Misc Constants
Private Const HWND_TOP = 0
Private Const WM_SETICON = &H80
Private Const ICON_SMALL = 0&
Private Const ICON_BIG = 1&

Private Hwnd As Long
Private FormCaption As String
Private IconPath_ As String

'CLASS PROPERTIES
Public Property Get FormName() As String 'Returns FormCaption
    FormName = FormCaption
End Property
Public Property Let FormName(ByVal Value As String)
    'Automatically establishes HWnd when valid FormCpation is passed
    If Len(Value) > 0 Then
        Hwnd = GetWindowID(Value)
        If Hwnd > 0 Then FormCaption = Value
    End If
End Property

Public Property Get WindowID() As Long 'Returns HWnd Value
    WindowID = Hwnd
End Property


'PUBLIC SUBROUTINES
Public Sub ModifyForm(ByVal FormName As String, Optional FormResizable As Boolean, Optional EnableResizeButtons As Boolean = True, Optional ShowinTaskbar As Boolean = True, Optional IconFilePath As String)
    Hwnd = GetWindowID(FormName)
    If Hwnd = 0 Then Exit Sub
    If ValidIcon(IconFilePath) Then Call SetIcon(IconFilePath)
    If EnableResizeButtons Then Call ResizeOptions(FormResizable:=FormResizable)
    If FormResizable Then Call SetWinStyle(WS_SIZEBOX)
    If ShowinTaskbar Then Call DisplayInTaskbar
End Sub


'PUBLIC CLASS FUNCTIONS
Public Function SetIcon(ByVal IconPath As String, Optional ByVal WindowName As String, Optional LargeIcon As Boolean = True) As Boolean
    If Not (Len(IconPath) > 0 And FileExists(IconPath) And ValidIcon(IconPath)) Then Exit Function
    'If WindowName is passed, use that. Otherwise default to internal hwnd, if neither are valid Exit Function
    Dim THwnd As Long: THwnd = IIf(Len(WindowName) > 0, GetWindowID(WindowName), Hwnd)
    If THwnd = 0 Then Exit Function
    
    Dim IconID As Long: IconID = ExtractIcon(0, IconPath, 0)
    SetIcon = SendMessage(THwnd, WM_SETICON, IIf(LargeIcon, ICON_BIG, ICON_SMALL), IconID) > 0
    If SetIcon Then IconPath_ = IconPath
End Function

Public Function SetWinStyle(Optional ByVal WStyle As WindowStyle)
    If Hwnd = 0 Then Exit Function
    SetWinStyle = SetWindowLong(Hwnd, GWL_STYLE, GetWindowLong(Hwnd, GWL_STYLE) Or WStyle)
    Call DrawMenuBar(Hwnd) 'Redraws the menu bar of the specified window. If the menu bar changes after the system has created the window, this function must be called to draw the changed menu bar.
End Function

Public Function DisplayInTaskbar() As Boolean
    If Hwnd = 0 Then Exit Function
    DisplayInTaskbar = SetWindowPos(Hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_HIDEWINDOW) And _
                       SetWindowLong(Hwnd, GWL_EXSTYLE, GetWindowLong(Hwnd, GWL_EXSTYLE) Or WS_SIZEBOX) And _
                       SetWindowPos(Hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_SHOWWINDOW)
    Call DrawMenuBar(Hwnd)
End Function

Public Function ResizeOptions(Optional Minimize As Boolean = True, Optional Maximize As Boolean = True, Optional FormResizable As Boolean) As Boolean
    If Hwnd = 0 Then Exit Function
    Dim ROptions As Long
    If Minimize Then ROptions = ROptions Or WS_MINIMIZEBOX
    If Maximize Then ROptions = ROptions Or WS_MAXIMIZEBOX
    If FormResizable Then ROptions = ROptions Or WS_SIZEBOX
    ResizeOptions = SetWindowLong(Hwnd, GWL_STYLE, GetWindowLong(Hwnd, GWL_STYLE) Or ROptions)
    Call DrawMenuBar(Hwnd)
End Function

Public Function GetWindowID(ByVal WindowName As String) As Long
    GetWindowID = FindWindow("ThunderDFrame", WindowName)
End Function


'PRIVATE CLASS FUNCTIONS
Private Function FileExists(ByVal FilePath As String) As Boolean
    With CreateObject("Scripting.FileSystemObject"): FileExists = .FileExists(FilePath): End With
End Function
Private Function FileExt(ByVal FilePath As String) As String
    With CreateObject("Scripting.FileSystemObject"): FileExt = .GetExtensionName(FilePath): End With
End Function

Private Function ValidIcon(ByVal IconPath As String) As Boolean
    ValidIcon = FileExt(IconPath) = "ico"
End Function
