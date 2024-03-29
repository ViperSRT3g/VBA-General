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

Public Enum ASSOCSTR
    ASSOCSTR_COMMAND = 1                    'A command string associated with a Shell verb
    ASSOCSTR_EXECUTABLE = 2                 'An executable from a Shell verb command string
    ASSOCSTR_FRIENDLYDOCNAME = 3            'The friendly name of a document type
    ASSOCSTR_FRIENDLYAPPNAME = 4            'The friendly name of an executable file
    ASSOCSTR_NOOPEN = 5                     'Ignore the information associated with the open subkey
    ASSOCSTR_SHELLNEWVALUE = 6              'Look under the ShellNew subkey.
    ASSOCSTR_DDECOMMAND = 7                 'A template for DDE commands
    ASSOCSTR_DDEIFEXEC = 8                  'The DDE command to use to create a process
    ASSOCSTR_DDEAPPLICATION = 9             'The application name in a DDE broadcast
    ASSOCSTR_DDETOPIC = 10                  'The topic name in a DDE broadcast
    ASSOCSTR_INFOTIP = 11                   'Corresponds to the InfoTip registry value
    ASSOCSTR_QUICKTIP = 12                  'Corresponds to the QuickTip registry value
    ASSOCSTR_TILEINFO = 13                  'Corresponds to the TileInfo registry value
    ASSOCSTR_CONTENTTYPE = 14               'Describes a general type of MIME file association
    ASSOCSTR_DEFAULTICON = 15               'Returns the path to the icon resources to use by default for this association
    ASSOCSTR_SHELLEXTENSION = 16            'For an object that has a Shell extension associated with it
    ASSOCSTR_DROPTARGET = 17                'For a verb invoked through COM and the IDropTarget interface
    ASSOCSTR_DELEGATEEXECUTE = 18           'For a verb invoked through COM and the IExecuteCommand interface
    ASSOCSTR_SUPPORTED_URI_PROTOCOLS = 19   'A string value of the URI protocol schemes
    ASSOCSTR_PROGID = 20                    'The ProgID provided by the app associated with the file type or URI scheme
    ASSOCSTR_APPID = 21                     'The AppUserModelID of the app associated with the file type or URI scheme
    ASSOCSTR_APPPUBLISHER = 22              'The publisher of the app associated with the file type or URI scheme
    ASSOCSTR_APPICONREFERENCE = 23          'The icon reference of the app associated with the file type or URI scheme
    ASSOCSTR_MAX = 24                       'last item in enum...
End Enum

Public Enum ASSOCF
    ASSOCF_NONE = &H0                   'None of the following options are set
    ASSOCF_INIT_NOREMAPCLSID = &H1      'Instructs IQueryAssociations interface methods not to map CLSID values to ProgID values
    ASSOCF_INIT_BYEXENAME = &H2         'Identifies the value of the pwszAssoc parameter of IQueryAssociations::Init as an executable file name
    ASSOCF_OPEN_BYEXENAME = &H2         'Identical to ASSOCF_INIT_BYEXENAME
    ASSOCF_INIT_DEFAULTTOSTAR = &H4     'Retrieve the comparable value from the * subkey when an IQueryAssociations method does not find the requested value
    ASSOCF_INIT_DEFAULTTOFOLDER = &H8   'Retrieve the comparable value from the Folder subkey
    ASSOCF_NOUSERSETTINGS = &H10        'Specifies that only HKEY_CLASSES_ROOT should be searched, and that HKEY_CURRENT_USER should be ignored
    ASSOCF_NOTRUNCATE = &H20            'Specifies that the return string should not be truncated
    ASSOCF_VERIFY = &H40                'This setting allows IQueryAssociations methods to read data from the user's hard disk for verification
    ASSOCF_REMAPRUNDLL = &H80           'Tells the method to ignore Rundll.exe and return information about its target
    ASSOCF_NOFIXUPS = &H100             'Instructs IQueryAssociations methods not to fix errors in the registry
    ASSOCF_IGNOREBASECLASS = &H200      'Specifies that the BaseClass value should be ignored
    ASSOCF_INIT_IGNOREUNKNOWN = &H400   'Specifies that the "Unknown" ProgID should be ignored; instead, fail
    ASSOCF_INIT_FIXED_PROGID = &H800    'Specifies that the supplied ProgID should be mapped using the system defaults
    ASSOCF_IS_PROTOCOL = &H1000         'Specifies that the value is a protocol, and should be mapped using the current user defaults
    ASSOCF_INIT_FOR_FILE = &H2000       'Specifies that the ProgID corresponds with a file extension based association. Use together with ASSOCF_INIT_FIXED_PROGID
End Enum

Public Enum NormalizationForm
    NormalizationOther = 0
    NormalizationC = 1
    NormalizationD = 2
    NormalizationKC = 5
    NormalizationKD = 6
End Enum

Public Enum FoldStringMapFlags
    MAP_FOLDCZONE = 16
    MAP_PRECOMPOSED = 32
    MAP_COMPOSITE = 64
    MAP_FOLDDIGITS = 128
End Enum

Public Enum SystemInfoType
    TotalRAMPercent = 1
    ProcessMem = 2
    ProcessCPU = 3
End Enum


Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


#If Win64 And VBA7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal nCount As Long)
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare PtrSafe Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal VKey As Long) As Integer
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "User32.dll" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwflags As Long) As Long
    Private Declare PtrSafe Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Function sndPlaySound32 Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
    Private Declare PtrSafe Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal dwflags As FlashWindowFlags) As Long
    Private Declare PtrSafe Function AssocQueryStringW Lib "shlwapi.dll" (ByVal Flags As ASSOCF, ByVal Str As ASSOCSTR, ByVal pszAssoc As Long, ByVal pszExtra As Long, ByVal pszOut As Long, ByRef pcchOut As Long) As Long
    Private Declare PtrSafe Function NormalizeString Lib "Normaliz" (ByVal normForm As Long, ByVal lpSrcString As LongPtr, ByVal cwSrcLength As Long, ByVal lpDstString As LongPtr, ByVal cwDstLength As Long) As Long
    Private Declare PtrSafe Function FoldString Lib "kernel32" Alias "FoldStringA" (ByVal dwMapFlags As FoldStringMapFlags, ByVal lpSrcStr As Long, ByVal cchSrc As Long, ByVal lpDestStr As Long, ByVal cchdest As Long) As Long
#Else
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal nCount As Long)
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
    Private Declare Function AssocQueryStringW Lib "shlwapi.dll" (ByVal Flags As ASSOCF, ByVal Str As ASSOCSTR, ByVal pszAssoc As Long, ByVal pszExtra As Long, ByVal pszOut As Long, ByRef pcchOut As Long) As Long
    Private Declare Function NormalizeString Lib "Normaliz" (ByVal normForm As Long, ByVal lpSrcString As LongPtr, ByVal cwSrcLength As Long, ByVal lpDstString As LongPtr, ByVal cwDstLength As Long) As Long
    Private Declare Function FoldString Lib "kernel32" Alias "FoldStringA" (ByVal dwMapFlags As FoldStringMapFlags, ByVal lpSrcStr As Long, ByVal cchSrc As Long, ByVal lpDestStr As Long, ByVal cchdest As Long) As Long
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
    Dim hWnd As Long: hWnd = GetUserformHwnd(UserFormCaption): If hWnd = 0 Then Exit Sub
    If ValidIcon(IconFilePath) Then Call SetIcon(UserFormCaption, IconFilePath)
    If EnableResizeButtons Then Call ResizeOptions(UserFormCaption, FormResizable:=FormResizable)
    If FormResizable Then Call SetWinStyle(UserFormCaption, WS_SIZEBOX)
    If ShowinTaskbar Then Call DisplayInTaskbar(UserFormCaption)
End Sub


'PUBLIC FUNCTIONS
'Exposes the FindWindow WINAPI directly for use outside of this module
Public Function GetHwnd(ByVal ClassName As String, ByVal WindowName As String) As Long
    GetHwnd = FindWindow(ClassName, WindowName)
End Function

'Returns the window handle for the given userform
Public Function GetUserformHwnd(ByVal UserFormCaption As String) As Long
    If Len(UserFormCaption) = 0 Then Exit Function
    GetUserformHwnd = FindWindow("ThunderDFrame", UserFormCaption)
End Function

'Returns the current window style value for the given userform
Public Function GetWinStyle(ByVal UserFormCaption As String) As Long
    Dim hWnd As Long: hWnd = GetUserformHwnd(UserFormCaption): If hWnd = 0 Then Exit Function
    GetWinStyle = GetWindowLong(hWnd, GWL_STYLE)
End Function

Public Function SetIcon(ByVal UserFormCaption As String, ByVal IconPath As String, Optional LargeIcon As Boolean = True) As Boolean
    If Len(IconPath) = 0 Then Exit Function
    With CreateObject("Scripting.FileSystemObject"): If Not (.FileExists(IconPath) And ValidIcon(IconPath)) Then Exit Function: End With: End Function
    
    Dim hWnd As Long: hWnd = GetUserformHwnd(UserFormCaption): If hWnd = 0 Then Exit Function
    Dim IconID As Long: IconID = ExtractIcon(0, IconPath, 0)
    SetIcon = Not CBool(SendMessage(hWnd, WM_SETICON, IIf(LargeIcon, ICON_BIG, ICON_SMALL), IconID))
End Function

'Sets the window style for a given userform. Flags can be combined in a single call
Public Function SetWinStyle(ByVal UserFormCaption As String, ByVal WStyle As WindowStyleFlags) As Boolean
    Dim hWnd As Long: hWnd = GetUserformHwnd(UserFormCaption): If hWnd = 0 Then Exit Function
    SetWinStyle = SetWindowLong(hWnd, GWL_STYLE, WStyle)
    Call DrawMenuBar(hWnd) 'Redraws the menu bar of the specified window. If the menu bar changes after the system has created the window, this function must be called to draw the changed menu bar.
End Function

Public Function SetWindowPosition(ByVal UserFormCaption As String, Optional ByVal WPFlags As WindowPositionFlags) As Boolean
    Dim hWnd As Long: hWnd = GetUserformHwnd(UserFormCaption): If hWnd = 0 Then Exit Function
    SetWindowPosition = SetWindowPos(hWnd, 0, 0, 0, 0, 0, WPFlags)
    Call DrawMenuBar(hWnd) 'Redraws the menu bar of the specified window. If the menu bar changes after the system has created the window, this function must be called to draw the changed menu bar.
End Function

Public Function DisplayInTaskbar(ByVal UserFormCaption As String) As Boolean
    Dim hWnd As Long: hWnd = GetUserformHwnd(UserFormCaption): If hWnd = 0 Then Exit Function
    DisplayInTaskbar = SetWindowPos(hWnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_HIDEWINDOW) And _
                       SetWindowLong(hWnd, GWL_EXSTYLE, GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_SIZEBOX) And _
                       SetWindowPos(hWnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_SHOWWINDOW)
    Call DrawMenuBar(hWnd)
End Function

Public Function ResizeOptions(ByVal UserFormCaption As String, Optional Minimize As Boolean = True, Optional Maximize As Boolean = True, Optional FormResizable As Boolean) As Boolean
    Dim hWnd As Long: hWnd = GetUserformHwnd(UserFormCaption): If hWnd = 0 Then Exit Function
    Dim ROptions As Long
    If Minimize Then ROptions = ROptions Or WS_MINIMIZEBOX
    If Maximize Then ROptions = ROptions Or WS_MAXIMIZEBOX
    If FormResizable Then ROptions = ROptions Or WS_SIZEBOX
    ResizeOptions = SetWindowLong(hWnd, GWL_STYLE, GetWindowLong(hWnd, GWL_STYLE) Or ROptions)
    Call DrawMenuBar(hWnd)
End Function

Public Function SetTransparency(ByVal UserFormCaption As String, R As Integer, G As Integer, B As Integer) As Boolean
    Dim hWnd As Long: hWnd = GetUserformHwnd(UserFormCaption): If hWnd = 0 Then Exit Function
    Dim Color As LongPtr: Color = VBA.RGB(R, G, B)
    Dim RetVal As Long: RetVal = RetVal Or WS_EX_LAYERED
    If SetWindowLong(hWnd, GWL_EXSTYLE, RetVal) Then SetTransparency = CBool(SetLayeredWindowAttributes(hWnd, Color, 0, LWA_COLORKEY))
End Function

Public Function SetTranslucency(ByVal UserFormCaption As String, ByVal Alpha As Byte) As Boolean
    Dim hWnd As Long: hWnd = GetUserformHwnd(UserFormCaption): If hWnd = 0 Then Exit Function
    Dim lStyle As Long: lStyle = GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    If SetWindowLong(hWnd, GWL_EXSTYLE, lStyle) Then SetTranslucency = CBool(SetLayeredWindowAttributes(hWnd, 0, CLng(Alpha), LWA_ALPHA))
End Function

Public Function SetTopmost(ByVal UserFormCaption As String, ByVal Toggle As Boolean) As Boolean
    Dim hWnd As Long: hWnd = GetUserformHwnd(UserFormCaption): If hWnd = 0 Then Exit Function
    SetTopmost = SetWindowPos(hWnd, IIf(Toggle, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE)
End Function

Public Function SetAsForeground(ByVal UserFormCaption As String) As Boolean
    Dim hWnd As Long: hWnd = GetUserformHwnd(UserFormCaption): If hWnd = 0 Then Exit Function
    SetAsForeground = CBool(SetForegroundWindow(hWnd))
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
    Dim hWnd As Long: hWnd = GetUserformHwnd(UserFormCaption): If hWnd = 0 Then Exit Function
    ShowTargetWindow = CBool(ShowWindow(hWnd, ShowWindowOption))
End Function

Public Function FlashUserform(ByVal UserFormCaption As String, Optional ByVal dwflags As FlashWindowFlags = FLASHW_ALL) As Boolean
    Dim hWnd As Long: hWnd = GetUserformHwnd(UserFormCaption): If hWnd = 0 Then Exit Function
    FlashUserform = FlashWindow(hWnd, dwflags)
End Function

Public Function FlashWin(ByVal hWnd As Long, Optional ByVal dwflags As FlashWindowFlags = FLASHW_ALL) As Boolean
    FlashWin = FlashWindow(hWnd, dwflags)
End Function

Public Function AssocQuery(FileExtension As String, Assoc As ASSOCSTR, Optional Flags As ASSOCF, Optional sExtra As String = "open") As String
    Dim Buffer As String: Buffer = String(1024, vbNullChar)
    Dim BufferSize As Long: BufferSize = Len(Buffer)

    Call AssocQueryStringW(Flags, Assoc, StrPtr(FileExtension), StrPtr(sExtra), StrPtr(Buffer), BufferSize)
    AssocQuery = Left(Buffer, BufferSize - 1)
End Function

'Utilized for importing files, or setting up links to files in Excel
Public Function ImportFile(FilePath As String, Target As Range, Optional FileLink As Boolean = True) As Boolean
    'FilePath - Path of the file to import
    'Target -   Range to place file object on
    'FileLink - Link to the file, or copy it into the workbook
    If Len(FilePath) = 0 Then Exit Function
    With CreateObject("Scripting.FileSystemObject")
        If Not .FileExists(FilePath) Then Exit Function
        Dim Output As OLEObject: Set Output = Target.Parent.OLEObjects.Add(FileName:=FilePath, Link:=FileLink, DisplayAsIcon:=True, _
            IconFileName:=AssocQuery("." & .GetExtensionName(FilePath), 2, 0, "open"), IconLabel:=.GetBaseName(FilePath), _
            Left:=Target.Left, Top:=Target.Top)
        Output.Width = Target.Width
    End With
    ImportFile = Not (Output Is Nothing)
    Set Output = Nothing
End Function

'Normalizes a given unicode string to W3C standards (https://unicode.org/reports/tr15/)
Public Function NormalizeUnicode(RawString As String, Optional Normalization As NormalizationForm = NormalizationC) As String
    If Len(RawString) = 0 Then Exit Function
    NormalizeUnicode = String(Len(RawString), 0)
    Call NormalizeString(Normalization, StrPtr(RawString), Len(RawString), StrPtr(NormalizeUnicode), Len(NormalizeUnicode))
End Function

'Normalizes and converts unicode characters to their transliterated English equivalents
Function CleanUnicode(RawString As String, Optional Normalization As NormalizationForm = NormalizationC) As String
    If Len(RawString) = 0 Then Exit Function
    Dim Index As Long
    
    CleanUnicode = NormalizeUnicode(RawString, Normalization)
    For Index = 0 To Len(RawString) * 2 - 2 Step 2
        Call FoldString(MAP_COMPOSITE, StrPtr(RawString) + Index, 1, StrPtr(CleanUnicode) + Index, 1)
    Next Index
End Function

'Returns the string stored at a given location in memory (Inverse of StrPtr() function)
Public Function PtrStr(StrPointer As LongPtr) As String
    If StrPointer = 0 Then Exit Function
    Dim Offset As Long, Buffer(1) As Byte, BufferStr As String
    Do
        Call CopyMemory(Buffer(0), ByVal StrPointer + Offset, 2)
        BufferStr = IIf(Not (Buffer(0) = 0 And Buffer(1) = 0), Buffer, vbNullString)
        PtrStr = PtrStr & BufferStr
        Offset = Offset + 2
    Loop Until Buffer(0) = 0 And Buffer(1) = 0
End Function

'Returns the corresponding System Information specified by the passed InfoType and Process name ("EXCEL.EXE"/"EXCEL") as examples
Public Function SystemInfo(InfoType As SystemInfoType, Optional ProcessName As String) As Variant
    Dim Query As Object, Item As Object
    
    Select Case InfoType
        Case TotalRAMPercent
            Set Query = GetObject("WinMgmts:root/cimv2").ExecQuery("Select * FROM Win32_PerfFormattedData_PerfOS_Memory")
        Case ProcessMem
            If Len(ProcessName) = 0 Then Exit Function
            Set Query = GetObject("winmgmts:").ExecQuery("SELECT WorkingSetSize FROM Win32_Process WHERE Name='" & ProcessName & "'")
        Case ProcessCPU
            If Len(ProcessName) = 0 Then Exit Function
            Set Query = GetObject("winmgmts:").ExecQuery("SELECT PercentProcessorTime FROM win32_PerfFormattedData_PerfProc_Process WHERE Name='" & ProcessName & "'")
    End Select
    
    For Each Item In Query
        Select Case InfoType
            Case TotalRAMPercent
                SystemInfo = Item.PercentCommittedBytesInUse
            Case ProcessMem 'Value in bytes
                SystemInfo = Item.WorkingSetSize
            Case ProcessCPU
                SystemInfo = Item.PercentProcessorTime
        End Select
    Next Item
End Function

'PRIVATE FUNCTIONS
Private Function ValidIcon(ByVal IconPath As String) As Boolean
    With CreateObject("Scripting.FileSystemObject"): ValidIcon = .GetExtensionName(IconPath) = "ico": End With
End Function
