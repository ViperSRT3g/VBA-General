Attribute VB_Name = "mod_WINAPIs"
Option Explicit

Public Enum WindowStyleFlags
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

Public Enum ExtendedWindowStyleFlags
    WS_EX_APPWINDOW = &H40000   'Forces a top-level window onto the taskbar when the window is visible.
    WS_EX_TOOLWINDOW = &H80     'The window is intended to be used as a floating toolbar.
    WS_EX_TOPMOST = &H8         'The window should be placed above all non-topmost windows and should stay above them, even when the window is deactivated. To add or remove this style, use the SetWindowPos function.
    WS_EX_TRANSPARENT = &H20    'The window should not be painted until siblings beneath the window (that were created by the same thread) have been painted. The window appears transparent because the bits of underlying sibling windows have already been painted. To achieve transparency without these restrictions, use the SetWindowRgn function.
End Enum

Public Enum WindowPositionFlags
    SWP_NOSIZE = &H1            'Retains the current size (ignores the cx and cy parameters).
    SWP_NOMOVE = &H2            'Retains the current position (ignores X and Y parameters).
    SWP_NOACTIVATE = &H10       'Does not activate the window. If this flag is not set, the window is activated and moved to the top of either the topmost or non-topmost group (depending on the setting of the hWndInsertAfter parameter).
    SWP_FRAMECHANGED = &H20     'Applies new frame styles set using the SetWindowLong function. Sends a WM_NCCALCSIZE message to the window, even if the window's size is not being changed. If this flag is not specified, WM_NCCALCSIZE is sent only when the window's size is being changed.
    SWP_SHOWWINDOW = &H40       'Displays the window.
    SWP_HIDEWINDOW = &H80       'Hides the window.
    SWP_NOREPOSITION = &H200    'Does not change the owner window's position in the Z order.
End Enum

Public Enum PlaySoundOptions
    SND_SYNC = &H0        '(Default) Play the sound synchronously. Code execution pauses until sound is complete.
    SND_ASYNC = &H1       'Play the sound asynchronously. Code execution does not wait for sound to complete.
    SND_NODEFAULT = &H2   'If the specified sound is not found, do not play the default sound (no sound is played).
    SND_MEMORY = &H4      'lpszSoundName is a memory file of the sound. Not used in VBA/VB6.
    SND_LOOP = &H8        'Continue playing sound in a loop until the next call to sndPlaySound.
    SND_NOSTOP = &H10     'Do not stop playing the current sound before playing the specified sound.
End Enum

Public Enum VirtualKeyCode
    VK_LBUTTON = &H1                'Left mouse button
    VK_RBUTTON = &H2                'Right mouse button
    VK_CANCEL = &H3                 'Control-break processing
    VK_MBUTTON = &H4                'Middle mouse button (three-button mouse)
    VK_XBUTTON1 = &H5               'X1 mouse button
    VK_XBUTTON2 = &H6               'X2 mouse button
    VK_BACK = &H8                   'BACKSPACE key
    VK_TAB = &H9                    'TAB key
    VK_CLEAR = &HC                  'CLEAR key
    VK_RETURN = &HD                 'ENTER key
    VK_SHIFT = &H10                 'SHIFT key
    VK_CONTROL = &H11               'CTRL key
    VK_MENU = &H12                  'ALT key
    VK_PAUSE = &H13                 'PAUSE key
    VK_CAPITAL = &H14               'CAPS LOCK key
    VK_KANA = &H15                  'IME Kana mode
    VK_HANGUEL = &H15               'IME Hanguel mode (maintained for compatibility; use VK_HANGUL)
    VK_HANGUL = &H15                'IME Hangul mode
    VK_IME_ON = &H16                'IME On
    VK_JUNJA = &H17                 'IME Junja mode
    VK_FINAL = &H18                 'IME final mode
    VK_HANJA = &H19                 'IME Hanja mode
    VK_KANJI = &H19                 'IME Kanji mode
    VK_IME_OFF = &H1A               'IME Off
    VK_ESCAPE = &H1B                'ESC key
    VK_CONVERT = &H1C               'IME convert
    VK_NONCONVERT = &H1D            'IME nonconvert
    VK_ACCEPT = &H1E                'IME accept
    VK_MODECHANGE = &H1F            'IME mode change request
    VK_SPACE = &H20                 'SPACEBAR
    VK_PRIOR = &H21                 'PAGE UP key
    VK_NEXT = &H22                  'PAGE DOWN key
    VK_END = &H23                   'END key
    VK_HOME = &H24                  'HOME key
    VK_LEFT = &H25                  'LEFT ARROW key
    VK_UP = &H26                    'UP ARROW key
    VK_RIGHT = &H27                 'RIGHT ARROW key
    VK_DOWN = &H28                  'DOWN ARROW key
    VK_SELECT = &H29                'SELECT key
    VK_PRINT = &H2A                 'PRINT key
    VK_EXECUTE = &H2B               'EXECUTE key
    VK_SNAPSHOT = &H2C              'PRINT SCREEN key
    VK_INSERT = &H2D                'INS key
    VK_DELETE = &H2E                'DEL key
    VK_HELP = &H2F                  'HELP key
    VK_ZERO = &H31                  '0 key
    VK_ONE = &H32                   '1 key
    VK_TWO = &H33                   '2 key
    VK_THREE = &H34                 '3 key
    VK_FOUR = &H35                  '4 key
    VK_FIVE = &H36                  '5 key
    VK_SIX = &H37                   '6 key
    VK_SEVEN = &H38                 '7 key
    VK_EIGHT = &H39                 '8 key
    VK_NINE = &H3A                  '9 key
    VK_A = &H41                     'A key
    VK_B = &H42                     'B key
    VK_C = &H43                     'C key
    VK_D = &H44                     'D key
    VK_E = &H46                     'E key
    VK_F = &H46                     'F key
    VK_G = &H47                     'G key
    VK_H = &H48                     'H key
    VK_I = &H49                     'I key
    VK_J = &H4A                     'J key
    VK_K = &H4B                     'K key
    VK_L = &H4C                     'L key
    VK_M = &H4D                     'M key
    VK_N = &H4E                     'N key
    VK_O = &H4F                     'O key
    VK_P = &H50                     'P key
    VK_Q = &H51                     'Q key
    VK_R = &H52                     'R key
    VK_S = &H53                     'S key
    VK_T = &H54                     'T key
    VK_U = &H55                     'U key
    VK_V = &H56                     'V key
    VK_W = &H57                     'W key
    VK_X = &H58                     'X key
    VK_Y = &H59                     'Y key
    VK_Z = &H5A                     'Z key
    VK_LWIN = &H5B                  'Left Windows key (Natural keyboard)
    VK_RWIN = &H5C                  'Right Windows key (Natural keyboard)
    VK_APPS = &H5D                  'Applications key (Natural keyboard)
    VK_SLEEP = &H5F                 'Computer Sleep key
    VK_NUMPAD0 = &H60               'Numeric keypad 0 key
    VK_NUMPAD1 = &H61               'Numeric keypad 1 key
    VK_NUMPAD2 = &H62               'Numeric keypad 2 key
    VK_NUMPAD3 = &H63               'Numeric keypad 3 key
    VK_NUMPAD4 = &H64               'Numeric keypad 4 key
    VK_NUMPAD5 = &H65               'Numeric keypad 5 key
    VK_NUMPAD6 = &H66               'Numeric keypad 6 key
    VK_NUMPAD7 = &H67               'Numeric keypad 7 key
    VK_NUMPAD8 = &H68               'Numeric keypad 8 key
    VK_NUMPAD9 = &H69               'Numeric keypad 9 key
    VK_MULTIPLY = &H6A              'Multiply key
    VK_ADD = &H6B                   'Add key
    VK_SEPARATOR = &H6C             'Separator key
    VK_SUBTRACT = &H6D              'Subtract key
    VK_DECIMAL = &H6E               'Decimal key
    VK_DIVIDE = &H6F                'Divide key
    VK_F1 = &H70                    'F1 key
    VK_F2 = &H71                    'F2 key
    VK_F3 = &H72                    'F3 key
    VK_F4 = &H73                    'F4 key
    VK_F5 = &H74                    'F5 key
    VK_F6 = &H75                    'F6 key
    VK_F7 = &H76                    'F7 key
    VK_F8 = &H77                    'F8 key
    VK_F9 = &H78                    'F9 key
    VK_F10 = &H79                   'F10 key
    VK_F11 = &H7A                   'F11 key
    VK_F12 = &H7B                   'F12 key
    VK_F13 = &H7C                   'F13 key
    VK_F14 = &H7D                   'F14 key
    VK_F15 = &H7E                   'F15 key
    VK_F16 = &H7F                   'F16 key
    VK_F17 = &H80                   'F17 key
    VK_F18 = &H81                   'F18 key
    VK_F19 = &H82                   'F19 key
    VK_F20 = &H83                   'F20 key
    VK_F21 = &H84                   'F21 key
    VK_F22 = &H85                   'F22 key
    VK_F23 = &H86                   'F23 key
    VK_F24 = &H87                   'F24 key
    VK_NUMLOCK = &H90               'NUM LOCK key
    VK_SCROLL = &H91                'SCROLL LOCK key
    VK_LSHIFT = &HA0                'Left SHIFT key
    VK_RSHIFT = &HA1                'Right SHIFT key
    VK_LCONTROL = &HA2              'Left CONTROL key
    VK_RCONTROL = &HA3              'Right CONTROL key
    VK_LMENU = &HA4                 'Left MENU key
    VK_RMENU = &HA5                 'Right MENU key
    VK_BROWSER_BACK = &HA6          'Browser Back key
    VK_BROWSER_FORWARD = &HA7       'Browser Forward key
    VK_BROWSER_REFRESH = &HA8       'Browser Refresh key
    VK_BROWSER_STOP = &HA9          'Browser Stop key
    VK_BROWSER_SEARCH = &HAA        'Browser Search key
    VK_BROWSER_FAVORITES = &HAB     'Browser Favorites key
    VK_BROWSER_HOME = &HAC          'Browser Start and Home key
    VK_VOLUME_MUTE = &HAD           'Volume Mute key
    VK_VOLUME_DOWN = &HAE           'Volume Down key
    VK_VOLUME_UP = &HAF             'Volume Up key
    VK_MEDIA_NEXT_TRACK = &HB0      'Next Track key
    VK_MEDIA_PREV_TRACK = &HB1      'Previous Track key
    VK_MEDIA_STOP = &HB2            'Stop Media key
    VK_MEDIA_PLAY_PAUSE = &HB3      'Play/Pause Media key
    VK_LAUNCH_MAIL = &HB4           'Start Mail key
    VK_LAUNCH_MEDIA_SELECT = &HB5   'Select Media key
    VK_LAUNCH_APP1 = &HB6           'Start Application 1 key
    VK_LAUNCH_APP2 = &HB7           'Start Application 2 key
    VK_OEM_1 = &HBA                 'Used for miscellaneous characters; it can vary by keyboard.
    VK_OEM_PLUS = &HBB              'For any country/region, the '+' key
    VK_OEM_COMMA = &HBC             'For any country/region, the ',' key
    VK_OEM_MINUS = &HBD             'For any country/region, the '-' key
    VK_OEM_PERIOD = &HBE            'For any country/region, the '.' key
    VK_OEM_2 = &HBF                 'Used for miscellaneous characters; it can vary by keyboard.
    VK_OEM_3 = &HC0                 'Used for miscellaneous characters; it can vary by keyboard.
    VK_OEM_4 = &HDB                 'Used for miscellaneous characters; it can vary by keyboard.
    VK_OEM_5 = &HDC                 'Used for miscellaneous characters; it can vary by keyboard.
    VK_OEM_6 = &HDD                 'Used for miscellaneous characters; it can vary by keyboard.
    VK_OEM_7 = &HDE                 'Used for miscellaneous characters; it can vary by keyboard.
    VK_OEM_8 = &HDF                 'Used for miscellaneous characters; it can vary by keyboard.
    VK_PROCESSKEY = &HE5            'IME PROCESS key
    VK_ATTN = &HF6                  'Attn key
    VK_CRSEL = &HF7                 'CrSel key
    VK_EXSEL = &HF8                 'ExSel key
    VK_EREOF = &HF9                 'Erase EOF key
    VK_PLAY = &HFA                  'Play key
    VK_ZOOM = &HFB                  'Zoom key
    VK_NONAME = &HFC                'Reserved
    VK_PA1 = &HFD                   'PA1 key
    VK_OEM_CLEAR = &HFE             'Clear key
End Enum

Public Enum ShowWindowFlags
    SW_HIDE
    SW_SHOWNORMAL
    SW_SHOWMINIMIZED
    SW_SHOWMAXIMIZED
    SW_SHOWNOACTIVATE
    SW_SHOW
    SW_MINIMIZE
    SW_SHOWMINNOACTIVE
    SW_SHOWNA
    SW_RESTORE
    SW_SHOWDEFAULT
    SW_FORCEMINIMIZE
End Enum

Public Enum FlashWindowFlags
    FLASHW_STOP = 0                             'Stops the window from flashing and restores it to its original state
    FLASHW_CAPTION = &H1&                       'Flashes the title bar on the window
    FLASHW_TRAY = &H2&                          'Flashes the button for the given window on the taskbar
    FLASHW_ALL = FLASHW_CAPTION Or FLASHW_TRAY  'Flashes the title bar and the taskbar button
    FLASHW_TIMER = &H4&                         'Flashes continuously until the FLASHW_STOP flag is set
    FLASHW_TIMERNOFG = &HC&                     'Flashes continuously as long as the window is in the background
End Enum

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type



#If Win64 And VBA7 Then
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare PtrSafe Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal VKey As Long) As Integer
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "User32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwflags As Long) As Long
    Private Declare PtrSafe Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Function sndPlaySound32 Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
    Private Declare PtrSafe Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal dwflags As FlashWindowFlags) As Long
#Else
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal VKey As Long) As Integer
    Private Declare Function GetTickCount Lib "kernel32" () As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function SetLayeredWindowAttributes Lib "User32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwflags As Long) As Long
    Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare Function sndPlaySound32 Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
    Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal dwflags As FlashWindowFlags) As Long
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
Public Sub ModifyForm(ByVal UserFormCaption As String, Optional FormResizable As Boolean, Optional EnableResizeButtons As Boolean = True, Optional ShowinTaskbar As Boolean = True, Optional IconFilePath As String)
    If Len(UserFormCaption) = 0 Then Exit Sub
    Dim hwnd As Long: hwnd = GetUserformHwnd(UserFormCaption): If hwnd = 0 Then Exit Sub
    If ValidIcon(IconFilePath) Then Call SetIcon(UserFormCaption, IconFilePath)
    If EnableResizeButtons Then Call ResizeOptions(UserFormCaption, FormResizable:=FormResizable)
    If FormResizable Then Call SetWinStyle(UserFormCaption, WS_SIZEBOX)
    If ShowinTaskbar Then Call DisplayInTaskbar(UserFormCaption)
End Sub


'PUBLIC FUNCTIONS
Public Function GetHwnd(ByVal ClassName As String, ByVal WindowName As String) As Long
    GetHwnd = FindWindow(ClassName, WindowName)
End Function

Public Function GetUserformHwnd(ByVal UserFormCaption As String) As Long
    If Len(UserFormCaption) = 0 Then Exit Function
    GetUserformHwnd = FindWindow("ThunderDFrame", UserFormCaption)
End Function

Public Function SetIcon(ByVal UserFormCaption As String, ByVal IconPath As String, Optional LargeIcon As Boolean = True) As Boolean
    If Len(IconPath) = 0 Then Exit Function
    If Not (FileExists(IconPath) And ValidIcon(IconPath)) Then Exit Function
    Dim hwnd As Long: hwnd = GetUserformHwnd(UserFormCaption): If hwnd = 0 Then Exit Function
    Dim IconID As Long: IconID = ExtractIcon(0, IconPath, 0)
    SetIcon = Not CBool(SendMessage(hwnd, WM_SETICON, IIf(LargeIcon, ICON_BIG, ICON_SMALL), IconID))
End Function

Public Function SetWinStyle(ByVal UserFormCaption As String, Optional ByVal WStyle As WindowStyleFlags) As Boolean
    Dim hwnd As Long: hwnd = GetUserformHwnd(UserFormCaption): If hwnd = 0 Then Exit Function
    SetWinStyle = SetWindowLong(hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) Or WStyle)
    Call DrawMenuBar(hwnd) 'Redraws the menu bar of the specified window. If the menu bar changes after the system has created the window, this function must be called to draw the changed menu bar.
End Function

Public Function SetWindowPosition(ByVal UserFormCaption As String, Optional ByVal WPFlags As WindowPositionFlags) As Boolean
    Dim hwnd As Long: hwnd = GetUserformHwnd(UserFormCaption): If hwnd = 0 Then Exit Function
    SetWindowPosition = SetWindowPos(hwnd, 0, 0, 0, 0, 0, WPFlags)
    Call DrawMenuBar(hwnd) 'Redraws the menu bar of the specified window. If the menu bar changes after the system has created the window, this function must be called to draw the changed menu bar.
End Function

Public Function DisplayInTaskbar(ByVal UserFormCaption As String) As Boolean
    Dim hwnd As Long: hwnd = GetUserformHwnd(UserFormCaption): If hwnd = 0 Then Exit Function
    DisplayInTaskbar = SetWindowPos(hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_HIDEWINDOW) And _
                       SetWindowLong(hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) Or WS_SIZEBOX) And _
                       SetWindowPos(hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_SHOWWINDOW)
    Call DrawMenuBar(hwnd)
End Function

Public Function ResizeOptions(ByVal UserFormCaption As String, Optional Minimize As Boolean = True, Optional Maximize As Boolean = True, Optional FormResizable As Boolean) As Boolean
    Dim hwnd As Long: hwnd = GetUserformHwnd(UserFormCaption): If hwnd = 0 Then Exit Function
    Dim ROptions As Long
    If Minimize Then ROptions = ROptions Or WS_MINIMIZEBOX
    If Maximize Then ROptions = ROptions Or WS_MAXIMIZEBOX
    If FormResizable Then ROptions = ROptions Or WS_SIZEBOX
    ResizeOptions = SetWindowLong(hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) Or ROptions)
    Call DrawMenuBar(hwnd)
End Function

Public Function SetTransparency(ByVal UserFormCaption As String, R As Integer, G As Integer, B As Integer) As Boolean
    Dim hwnd As Long: hwnd = GetUserformHwnd(UserFormCaption): If hwnd = 0 Then Exit Function
    Dim Color As LongPtr: Color = VBA.RGB(R, G, B)
    Dim RetVal As Long: RetVal = RetVal Or WS_EX_LAYERED
    If SetWindowLong(hwnd, GWL_EXSTYLE, RetVal) Then SetTransparency = CBool(SetLayeredWindowAttributes(hwnd, Color, 0, LWA_COLORKEY))
End Function

Public Function SetTranslucency(ByVal UserFormCaption As String, ByVal Alpha As Byte) As Boolean
    Dim hwnd As Long: hwnd = GetUserformHwnd(UserFormCaption): If hwnd = 0 Then Exit Function
    Dim lStyle As Long: lStyle = GetWindowLong(hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    If SetWindowLong(hwnd, GWL_EXSTYLE, lStyle) Then SetTranslucency = CBool(SetLayeredWindowAttributes(hwnd, 0, CLng(Alpha), LWA_ALPHA))
End Function

Public Function SetTopmost(ByVal UserFormCaption As String, ByVal Toggle As Boolean) As Boolean
    Dim hwnd As Long: hwnd = GetUserformHwnd(UserFormCaption): If hwnd = 0 Then Exit Function
    SetTopmost = SetWindowPos(hwnd, IIf(Toggle, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE)
End Function

Public Function SetAsForeground(ByVal UserFormCaption As String) As Boolean
    Dim hwnd As Long: hwnd = GetUserformHwnd(UserFormCaption): If hwnd = 0 Then Exit Function
    SetAsForeground = CBool(SetForegroundWindow(hwnd))
End Function

Public Function Intersect(ByRef Rect1 As RECT, Rect2 As RECT) As Boolean
    Dim ResultRect As RECT: Intersect = IntersectRect(ResultRect, Rect1, Rect2)
End Function

Public Function GetTicks() As Long
    GetTicks = GetTickCount
End Function

Public Function GetAKeyState(ByVal VKey As VirtualKeyCode) As Boolean
    GetAKeyState = GetAsyncKeyState(VKey)
End Function

Public Function ShowTargetWindow(ByVal UserFormCaption As String, ByVal ShowWindowOption As ShowWindowFlags) As Boolean
    Dim hwnd As Long: hwnd = GetUserformHwnd(UserFormCaption): If hwnd = 0 Then Exit Function
    ShowTargetWindow = CBool(ShowWindow(hwnd, ShowWindowOption))
End Function

Public Function FlashUserform(ByVal UserFormCaption As String, Optional ByVal dwflags As FlashWindowFlags = FLASHW_ALL) As Boolean
    Dim hwnd As Long: hwnd = GetUserformHwnd(UserFormCaption): If hwnd = 0 Then Exit Function
    FlashUserform = FlashWindow(hwnd, dwflags)
End Function

Public Function FlashWin(ByVal hwnd As Long, Optional ByVal dwflags As FlashWindowFlags = FLASHW_ALL) As Boolean
    FlashWin = FlashWindow(hwnd, dwflags)
End Function


'PRIVATE FUNCTIONS
Private Function FileExists(ByVal FilePath As String) As Boolean
    With CreateObject("Scripting.FileSystemObject"): FileExists = .FileExists(FilePath): End With: End Function
Private Function FileExt(ByVal FilePath As String) As String
    With CreateObject("Scripting.FileSystemObject"): FileExt = .GetExtensionName(FilePath): End With: End Function
Private Function ValidIcon(ByVal IconPath As String) As Boolean
    ValidIcon = FileExt(IconPath) = "ico": End Function
