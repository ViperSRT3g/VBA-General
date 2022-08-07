VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class_WINAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''=============================================================================
''' VBA-General : A toolbox for VBA features
''' -----------------------------------------
''' https://github.com/ViperSRT3g/VBA-General
''' -----------------------------------------
''' MIT License
'''
''' Permission is hereby granted, free of charge, to any person obtaining a copy
''' of this software and associated documentation files (the "Software"), to
''' deal in the Software without restriction, including without limitation the
''' rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
''' sell copies of the Software, and to permit persons to whom the Software is
''' furnished to do so, subject to the following conditions:
'''
''' The above copyright notice and this permission notice shall be included in
''' all copies or substantial portions of the Software.
'''
''' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
''' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
''' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
''' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
''' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
''' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
''' IN THE SOFTWARE.
'''=============================================================================

Option Explicit

'CONSTANTS DECLARATIONS
Private Const VT_BYREF = &H4000

'Layered Window Attributes Constants
Private Const LWA_ALPHA = &H2&
Private Const LWA_COLORKEY = &H1

'LOGFONT Type Constants
Private Const LF_FACESIZE = 32

'SendMessage Constants
Private Const WM_SETICON = &H80

'SendMessage Misc Constants
Private Const ICON_SMALL = 0&
Private Const ICON_BIG = 1&

'SystemParametersInfo Constants
Private Const SPI_GETNONCLIENTMETRICS = 41


'TYPE DECLARATIONS
Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(1 To LF_FACESIZE) As Byte
End Type

Private Type NONCLIENTMETRICS
    cbSize As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As LOGFONT
    iSMCaptionWidth As Long
    iSMCaptionHeight As Long
    lfSMCaptionFont As LOGFONT
    iMenuWidth As Long
    iMenuHeight As Long
    lfMenuFont As LOGFONT
    lfStatusFont As LOGFONT
    lfMessageFont As LOGFONT
    iPaddedBorderWidth As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type REMOTE_MEMORY
    memValue As Variant
    remoteVT As Variant 'Will be linked to the first 2 bytes of 'memValue' - see 'InitRemoteMemory'
    isInitialized As Boolean 'In case state is lost
End Type

Private Type SIZE
    cX As Long
    cY As Long
End Type



'ENUMERATION DECLARATIONS
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

Public Enum ExtendedWindowStyleFlags
    WS_EX_ACCEPTFILES = &H10             'The window accepts drag-drop files.
    WS_EX_APPWINDOW = &H40000            'Forces a top-level window onto the taskbar when the window is visible.
    WS_EX_CLIENTEDGE = &H200             'The window has a border with a sunken edge.
    WS_EX_COMPOSITED = &H2000000         'Paints all descendants of a window in bottom-to-top painting order using double-buffering. Bottom-to-top painting order allows a descendent window to have translucency (alpha) and transparency (color-key) effects, but only if the descendent window also has the WS_EX_TRANSPARENT bit set. Double-buffering allows the window and its descendents to be painted without flicker. This cannot be used if the window has a class style of either CS_OWNDC or CS_CLASSDC.
    WS_EX_CONTEXTHELP = &H400            'The title bar of the window includes a question mark. When the user clicks the question mark, the cursor changes to a question mark with a pointer. If the user then clicks a child window, the child receives a WM_HELP message. The child window should pass the message to the parent window procedure, which should call the WinHelp function using the HELP_WM_HELP command. The Help application displays a pop-up window that typically contains help for the child window. WS_EX_CONTEXTHELP cannot be used with the WS_MAXIMIZEBOX or WS_MINIMIZEBOX styles.
    WS_EX_CONTROLPARENT = &H10000        'The window itself contains child windows that should take part in dialog box navigation. If this style is specified, the dialog manager recurses into children of this window when performing navigation operations such as handling the TAB key, an arrow key, or a keyboard mnemonic.
    WS_EX_DLGMODALFRAME = &H1            'The window has a double border; the window can, optionally, be created with a title bar by specifying the WS_CAPTION style in the dwStyle parameter.
    WS_EX_LAYERED = &H80000              'The window is a layered window. This style cannot be used if the window has a class style of either CS_OWNDC or CS_CLASSDC. Windows 8: The WS_EX_LAYERED style is supported for top-level windows and child windows. Previous Windows versions support WS_EX_LAYERED only for top-level windows.
    WS_EX_LAYOUTRTL = &H400000           'If the shell language is Hebrew, Arabic, or another language that supports reading order alignment, the horizontal origin of the window is on the right edge. Increasing horizontal values advance to the left.
    WS_EX_LEFT = &H0                     'The window has generic left-aligned properties. This is the default.
    WS_EX_LEFTSCROLLBAR = &H4000         'If the shell language is Hebrew, Arabic, or another language that supports reading order alignment, the vertical scroll bar (if present) is to the left of the client area. For other languages, the style is ignored.
    WS_EX_LTRREADING = &H0               'The window text is displayed using left-to-right reading-order properties. This is the default.
    WS_EX_MDICHILD = &H40                'The window is a MDI child window.
    WS_EX_NOACTIVATE = &H8000000         'A top-level window created with this style does not become the foreground window when the user clicks it. The system does not bring this window to the foreground when the user minimizes or closes the foreground window. The window should not be activated through programmatic access or via keyboard navigation by accessible technology, such as Narrator. To activate the window, use the SetActiveWindow or SetForegroundWindow function. The window does not appear on the taskbar by default. To force the window to appear on the taskbar, use the WS_EX_APPWINDOW style.
    WS_EX_NOINHERITLAYOUT = &H100000     'The window does not pass its window layout to its child windows.
    WS_EX_NOPARENTNOTIFY = &H4           'The child window created with this style does not send the WM_PARENTNOTIFY message to its parent window when it is created or destroyed.
    WS_EX_NOREDIRECTIONBITMAP = &H200000 'The window does not render to a redirection surface. This is for windows that do not have visible content or that use mechanisms other than surfaces to provide their visual.
    WS_EX_RIGHT = &H1000                 'The window has generic "right-aligned" properties. This depends on the window class. This style has an effect only if the shell language is Hebrew, Arabic, or another language that supports reading-order alignment; otherwise, the style is ignored. Using the WS_EX_RIGHT style for static or edit controls has the same effect as using the SS_RIGHT or ES_RIGHT style, respectively. Using this style with button controls has the same effect as using BS_RIGHT and BS_RIGHTBUTTON styles.
    WS_EX_RIGHTSCROLLBAR = &H0           'The vertical scroll bar (if present) is to the right of the client area. This is the default.
    WS_EX_RTLREADING = &H2000            'If the shell language is Hebrew, Arabic, or another language that supports reading-order alignment, the window text is displayed using right-to-left reading-order properties. For other languages, the style is ignored.
    WS_EX_STATICEDGE = &H20000           'The window has a three-dimensional border style intended to be used for items that do not accept user input.
    WS_EX_TOOLWINDOW = &H80              'The window is intended to be used as a floating toolbar.
    WS_EX_TOPMOST = &H8                  'The window should be placed above all non-topmost windows and should stay above them, even when the window is deactivated. To add or remove this style, use the SetWindowPos function.
    WS_EX_TRANSPARENT = &H20             'The window should not be painted until siblings beneath the window (that were created by the same thread) have been painted. The window appears transparent because the bits of underlying sibling windows have already been painted. To achieve transparency without these restrictions, use the SetWindowRgn function.
    WS_EX_WINDOWEDGE = &H100             'The window has a border with a raised edge.
    WS_EX_OVERLAPPEDWINDOW = WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE 'The window is an overlapped window.
    WS_EX_PALETTEWINDOW = WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST 'The window is palette window, which is a modeless dialog box that presents an array of commands.
End Enum

Public Enum FlashWindowFlags
    FLASHW_STOP = 0                             'Stops the window from flashing and restores it to its original state
    FLASHW_CAPTION = &H1&                       'Flashes the title bar on the window
    FLASHW_TRAY = &H2&                          'Flashes the button for the given window on the taskbar
    FLASHW_ALL = FLASHW_CAPTION Or FLASHW_TRAY  'Flashes the title bar and the taskbar button
    FLASHW_TIMER = &H4&                         'Flashes continuously until the FLASHW_STOP flag is set
    FLASHW_TIMERNOFG = &HC&                     'Flashes continuously as long as the window is in the background
End Enum

Public Enum FoldStringMapFlags
    MAP_FOLDCZONE = 16
    MAP_PRECOMPOSED = 32
    MAP_COMPOSITE = 64
    MAP_FOLDDIGITS = 128
End Enum

Public Enum GetDeviceCapsEnum
    DRIVERVERSION = &H0     'The device driver version.
    TECHNOLOGY = &H2        'Device technology
    HORZSIZE = &H4          'Width, in millimeters, of the physical screen.
    VERTSIZE = &H6          'Height, in millimeters, of the physical screen.
    HORZRES = &H8           'Width, in pixels, of the screen; or for printers, the width, in pixels, of the printable area of the page.
    VERTRES = &HA           'Height, in raster lines, of the screen; or for printers, the height, in pixels, of the printable area of the page.
    LOGPIXELSX = &H58       'Number of pixels per logical inch along the screen width. In a system with multiple display monitors, this value is the same for all monitors.
    LOGPIXELSY = &H5A       'Number of pixels per logical inch along the screen height. In a system with multiple display monitors, this value is the same for all monitors.
    BITSPIXEL = &HC         'Number of adjacent color bits for each pixel.
    PLANES = &HE            'Number of color planes.
    NUMBRUSHES = &H10       'Number of device-specific brushes.
    NUMPENS = &H12          'Number of device-specific pens.
    NUMFONTS = &H16         'Number of device-specific fonts.
    NUMCOLORS = &H18        'Number of entries in the device's color table, if the device has a color depth of no more than 8 bits per pixel. For devices with greater color depths, 1 is returned.
    ASPECTX = &H28          'Relative width of a device pixel used for line drawing.
    ASPECTY = &H2A          'Relative height of a device pixel used for line drawing.
    PDEVICESIZE = &H1A      'Reserved.
    CLIPCAPS = &H24         'Flag that indicates the clipping capabilities of the device. If the device can clip to a rectangle, it is 1. Otherwise, it is 0.
    SIZEPALETTE = &H68      'Number of entries in the system palette. This index is valid only if the device driver sets the RC_PALETTE bit in the RASTERCAPS index and is available only if the driver is compatible with 16-bit Windows.
    NUMRESERVED = &H6A      'Number of reserved entries in the system palette. This index is valid only if the device driver sets the RC_PALETTE bit in the RASTERCAPS index and is available only if the driver is compatible with 16-bit Windows.
    COLORRES = &H6C         'Actual color resolution of the device, in bits per pixel. This index is valid only if the device driver sets the RC_PALETTE bit in the RASTERCAPS index and is available only if the driver is compatible with 16-bit Windows.
    RASTERCAPS = &H26       'Value that indicates the raster capabilities of the device
    CURVECAPS = &H1C        'Value that indicates the curve capabilities of the device
    LINECAPS = &H1E         'Value that indicates the line capabilities of the device
    POLYGONALCAPS = &H20    'Value that indicates the polygon capabilities of the device
    TEXTCAPS = &H22         'Value that indicates the text capabilities of the device
    ASPECTXY = &H2C         'Diagonal width of the device pixel used for line drawing.
End Enum

Public Enum GetWindowCmd
    GW_CHILD = &H5          'The retrieved handle identifies the child window at the top of the Z order, if the specified window is a parent window;
                            'otherwise, the retrieved handle is NULL. The function examines only child windows of the specified window.
                            'It does not examine descendant windows.
    GW_ENABLEDPOPUP = &H6   'The retrieved handle identifies the enabled popup window owned by the specified window (the search uses the first such
                            'window found using GW_HWNDNEXT); otherwise, if there are no enabled popup windows, the retrieved handle is that of the
                            'specified window.
    GW_HWNDFIRST = &H0      'The retrieved handle identifies the window of the same type that is highest in the Z order. If the specified window is a
                            'topmost window, the handle identifies a topmost window. If the specified window is a top-level window, the handle
                            'identifies a top-level window. If the specified window is a child window, the handle identifies a sibling window.
    GW_HWNDLAST = &H1       'The retrieved handle identifies the window of the same type that is lowest in the Z order. If the specified window is a
                            'topmost window, the handle identifies a topmost window. If the specified window is a top-level window, the handle
                            'identifies a top-level window. If the specified window is a child window, the handle identifies a sibling window.
    GW_HWNDNEXT = &H2       'The retrieved handle identifies the window below the specified window in the Z order. If the specified window is a topmost
                            'window, the handle identifies a topmost window. If the specified window is a top-level window, the handle identifies a
                            'top-level window. If the specified window is a child window, the handle identifies a sibling window.
    GW_HWNDPREV = &H3       'The retrieved handle identifies the window above the specified window in the Z order. If the specified window is a topmost
                            'window, the handle identifies a topmost window. If the specified window is a top-level window, the handle identifies a
                            'top-level window. If the specified window is a child window, the handle identifies a sibling window.
    GW_OWNER = &H4          'The retrieved handle identifies the specified window's owner window, if any. For more information, see Owned Windows.
End Enum

Public Enum GetWindowLongEnum
    GWLP_USERDATA = (-21)   'Retrieves the user data associated with the window.
                            'This data is intended for use by the application that created the window. Its value is initially zero.
    GWL_EXSTYLE = (-20)     'Retrieves the extended window styles.
    GWL_STYLE = (-16)       'Retrieves the window styles.
    GWLP_ID = (-12)         'Retrieves the identifier of the window.
    GWL_HWNDPARENT = (-8)   'Retrieves a handle to the parent window, if there is one.
    GWLP_HINSTANCE = (-6)   'Retrieves a handle to the application instance.
    GWLP_WNDPROC = (-4)     'Retrieves the pointer to the window procedure, or a handle representing the pointer to the window procedure.
                            'You must use the CallWindowProc function to call the window procedure.
End Enum

Public Enum hWndInsertAfterEnum
    HWND_TOPMOST = -1       'Places the window at the bottom of the Z order. If the hWnd parameter identifies a topmost window, the window loses its
                            'topmost status and is placed at the bottom of all other windows.
    HWND_NOTOPMOST = -2     'Places the window above all non-topmost windows (that is, behind all topmost windows). This flag has no effect if the
                            'window is already a non-topmost window.
    HWND_TOP = 0            'Places the window at the top of the Z order.
    HWND_BOTTOM = 1         'Places the window above all non-topmost windows. The window maintains its topmost position even when it is deactivated.
End Enum

Public Enum MouseEventFlags
    MOUSEEVENTF_ABSOLUTE = &H8000   'The dx and dy parameters contain normalized absolute coordinates. If not set, those parameters contain relative
                                    'data. The change in position since the last reported position
    MOUSEEVENTF_LEFTDOWN = &H2      'The left button is down.
    MOUSEEVENTF_LEFTUP = &H4        'The left button is up.
    MOUSEEVENTF_MIDDLEDOWN = &H20   'The middle button is down.
    MOUSEEVENTF_MIDDLEUP = &H40     'The middle button is up.
    MOUSEEVENTF_MOVE = &H1          'Movement occurred.
    MOUSEEVENTF_RIGHTDOWN = &H8     'The right button is down.
    MOUSEEVENTF_RIGHTUP = &H10      'The right button is up.
    MOUSEEVENTF_XDOWN = &H80        'An X button was pressed.
    MOUSEEVENTF_XUP = &H100         'An X button was released.
    MOUSEEVENTF_WHEEL = &H800       'The wheel has been moved, if the mouse has a wheel. The amount of movement is specified in dwData
    MOUSEEVENTF_HWHEEL = &H1000     'The wheel button is tilted.
End Enum

Public Enum NormalizationForm
    NormalizationOther = 0
    NormalizationC = 1
    NormalizationD = 2
    NormalizationKC = 5
    NormalizationKD = 6
End Enum

Public Enum PlaySoundOptions
    SND_SYNC = &H0        '(Default) Play the sound synchronously. Code execution pauses until sound is complete.
    SND_ASYNC = &H1       'Play the sound asynchronously. Code execution does not wait for sound to complete.
    SND_NODEFAULT = &H2   'If the specified sound is not found, do not play the default sound (no sound is played).
    SND_MEMORY = &H4      'lpszSoundName is a memory file of the sound. Not used in VBA/VB6.
    SND_LOOP = &H8        'Continue playing sound in a loop until the next call to sndPlaySound.
    SND_NOSTOP = &H10     'Do not stop playing the current sound before playing the specified sound.
End Enum

Public Enum ShowWindowFlags
    SW_HIDE             '[SW_HIDE] Hides the window and activates another window.
    SW_SHOWNORMAL       '[SW_SHOWNORMAL] Activates and displays a window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when displaying the window for the first time.
    SW_SHOWMINIMIZED    '[SW_SHOWMINIMIZED] Activates the window and displays it as a minimized window.
    SW_SHOWMAXIMIZED    '[SW_SHOWMAXIMIZED] Activates the window and displays it as a maximized window.
    SW_SHOWNOACTIVATE   '[SW_SHOWNOACTIVATE] Displays a window in its most recent size and position. This value is similar to SW_SHOWNORMAL, except that the window is not activated.
    SW_SHOW             '[SW_SHOW] Activates the window and displays it in its current size and position.
    SW_MINIMIZE         '[SW_MINIMIZE] Minimizes the specified window and activates the next top-level window in the Z order.
    SW_SHOWMINNOACTIVE  '[SW_SHOWMINNOACTIVE] Displays the window as a minimized window. This value is similar to SW_SHOWMINIMIZED, except the window is not activated.
    SW_SHOWNA           '[SW_SHOWNA] Displays the window in its current size and position. This value is similar to SW_SHOW, except that the window is not activated.
    SW_RESTORE          '[SW_RESTORE] Activates and displays the window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when restoring a minimized window.
    SW_SHOWDEFAULT      '[SW_SHOWDEFAULT] Sets the show state based on the SW_ value specified in the STARTUPINFO structure passed to the CreateProcess function by the program that started the application.
    SW_FORCEMINIMIZE    '[SW_FORCEMINIMIZE] Minimizes a window, even if the thread that owns the window is not responding. This flag should only be used when minimizing windows from a different thread.
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

Public Enum SetWindowPositionFlags
    SWP_NOSIZE = &H1            'Retains the current size (ignores the cx and cy parameters).
    SWP_NOMOVE = &H2            'Retains the current position (ignores X and Y parameters).
    SWP_NOACTIVATE = &H10       'Does not activate the window. If this flag is not set, the window is activated and moved to the top of either the
                                'topmost or non-topmost group (depending on the setting of the hWndInsertAfter parameter).
    SWP_DRAWFRAME = &H20        'Draws a frame (defined in the window's class description) around the window.
    SWP_FRAMECHANGED = &H20     'Applies new frame styles set using the SetWindowLong function. Sends a WM_NCCALCSIZE message to the window, even if
                                'the window's size is not being changed. If this flag is not specified, WM_NCCALCSIZE is sent only when the window's
                                'size is being changed.
    SWP_SHOWWINDOW = &H40       'Displays the window.
    SWP_HIDEWINDOW = &H80       'Hides the window.
    SWP_NOREPOSITION = &H200    'Does not change the owner window's position in the Z order.
End Enum

Public Enum WindowStyleFlags
    WS_BORDER = &H800000        'The window has a thin-line border.
    WS_CAPTION = &HC00000       'The window has a title bar (includes the WS_BORDER style).
    WS_CHILD = &H40000000       'The window is a child window. A window with this style cannot have a menu bar. This style cannot be used with the WS_POPUP style.
    WS_DISABLED = &H8000000     'The window is initially disabled. A disabled window cannot receive input from the user. To change this after a window
                                'has been created, use the EnableWindow function.
    WS_MAXIMIZE = &H1000000     'The window is initially maximized.
    WS_MAXIMIZEBOX = &H10000    'The window has a maximize button. Cannot be combined with the WS_EX_CONTEXTHELP style. The WS_SYSMENU style must also
                                'be specified.
    WS_MINIMIZE = &H20000000    'The window is initially minimized.
    WS_MINIMIZEBOX = &H20000    'The window has a minimize button. Cannot be combined with the WS_EX_CONTEXTHELP style. The WS_SYSMENU style must also
                                'be specified.
    WS_POPUP = &H80000000       'The window is a pop-up window. This style cannot be used with the WS_CHILD style.
    WS_SIZEBOX = &H40000        'The window has a sizing border.
    WS_SYSMENU = &H80000        'The window has a window menu on its title bar. The WS_CAPTION style must also be specified.
    WS_VISIBLE = &H10000000     'The window is initially visible. This style can be turned on and off by using the ShowWindow or SetWindowPos function.
End Enum

'WINAPI DECLARATIONS
#If Win64 And VBA7 Then
    Private Declare PtrSafe Function AssocQueryStringW Lib "shlwapi.dll" (ByVal flags As ASSOCF, ByVal Str As ASSOCSTR, ByVal pszAssoc As LongPtr, ByVal pszExtra As LongPtr, ByVal pszOut As LongPtr, ByRef pcchOut As Long) As Long
    Private Declare PtrSafe Function AtlAxGetControl Lib "atl" (ByVal hwnd As LongPtr, Unk As stdole.IUnknown) As Long
    Private Declare PtrSafe Function AtlAxWinInit Lib "Atl.dll" () As Long
    Private Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As LongPtr, ByVal hwnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Private Declare PtrSafe Function ColorHLSToRGB Lib "shlwapi.dll" (ByVal wHue As Integer, ByVal wLuminance As Integer, ByVal wSaturation As Integer) As Long
    Private Declare PtrSafe Function ColorRGBToHLS Lib "shlwapi" (ByVal clrRGB As Long, ByRef wHue As Integer, ByRef wLuminance As Integer, ByRef wSaturation As Integer) As Long
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    Private Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As ExtendedWindowStyleFlags, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As WindowStyleFlags, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal HwndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr
    Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
    Private Declare PtrSafe Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As LongPtr
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function FlashWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal dwFlags As FlashWindowFlags) As Long
    Private Declare PtrSafe Function FoldString Lib "kernel32" Alias "FoldStringA" (ByVal dwMapFlags As FoldStringMapFlags, ByVal lpSrcStr As LongPtr, ByVal cchSrc As Long, ByVal lpDestStr As LongPtr, ByVal cchdest As Long) As Long
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal VKey As VirtualKeyCode) As Long
    Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As GetDeviceCapsEnum) As Long
    Private Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
    Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As LongPtr
    Private Declare PtrSafe Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As LongPtr, ByVal wFlag As GetWindowCmd) As LongPtr
    Private Declare PtrSafe Function GetPixel Lib "gdi32" (ByVal hDC As LongPtr, ByVal X As Long, ByVal Y As Long) As Long
    Private Declare PtrSafe Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As LongPtr, ByVal lpString As String) As LongPtr
    Private Declare PtrSafe Function GetTickCount64 Lib "kernel32" () As LongPtr
    Private Declare PtrSafe Function GetWindow Lib "user32.dll" (ByVal hwnd As LongPtr, ByVal uCmd As GetWindowCmd) As LongPtr
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As GetWindowLongEnum) As LongPtr
    Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As Long
    Private Declare PtrSafe Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
    Private Declare PtrSafe Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByVal hUf As LongPtr) As Long
    Private Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As MouseEventFlags, ByVal dX As Long, ByVal dY As Long, ByVal dwData As Long, ByVal dwExtraInfo As LongPtr)
    Private Declare PtrSafe Function NormalizeString Lib "Normaliz" (ByVal normForm As Long, ByVal lpSrcString As LongPtr, ByVal cwSrcLength As Long, ByVal lpDstString As LongPtr, ByVal cwDstLength As Long) As Long
    Private Declare PtrSafe Function ReleaseCapture Lib "user32" () As Long
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As Long
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
    Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As LongPtr, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal hData As LongPtr) As Long
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As GetWindowLongEnum, ByVal dwNewLong As LongPtr) As LongPtr
    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As hWndInsertAfterEnum, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As SetWindowPositionFlags) As Boolean
    Private Declare PtrSafe Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String) As Long
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As ShowWindowFlags) As Long
    Private Declare PtrSafe Function sndPlaySound32 Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As PlaySoundOptions) As Long
    Private Declare PtrSafe Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
    Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
    Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal Point As LongLong) As LongPtr
#Else
    Private Declare Function AssocQueryStringW Lib "shlwapi.dll" (ByVal Flags As ASSOCF, ByVal Str As ASSOCSTR, ByVal pszAssoc As Long, ByVal pszExtra As Long, ByVal pszOut As Long, ByRef pcchOut As Long) As Long
    Private Declare Function AtlAxGetControl Lib "atl" (ByVal hwnd As Long, Unk As stdole.IUnknown) As Long
    Private Declare Function AtlAxWinInit Lib "Atl.dll" () As Long
    Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function ColorHLSToRGB Lib "shlwapi.dll" (ByVal wHue As Integer, ByVal wLuminance As Integer, ByVal wSaturation As Integer) As Long
    Private Declare Function ColorRGBToHLS Lib "shlwapi" (ByVal clrRGB As Long, ByRef wHue As Integer, ByRef wLuminance As Integer, ByRef wSaturation As Integer) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As ExtendedWindowStyleFlags, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As WindowStyleFlags, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
    Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
    Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal dwflags As FlashWindowFlags) As Long
    Private Declare Function FoldString Lib "kernel32" Alias "FoldStringA" (ByVal dwMapFlags As FoldStringMapFlags, ByVal lpSrcStr As Long, ByVal cchSrc As Long, ByVal lpDestStr As Long, ByVal cchdest As Long) As Long
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal VKey As VirtualKeyCode) As Long
    Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As GetDeviceCapsEnum) As Long
    Private Declare Function GetFocus Lib "user32" () As Long
    Private Declare Function GetForegroundWindow Lib "user32" () As Long
    Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
    Private Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As GetWindowCmd) As Long
    Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
    Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
    Private Declare Function GetTickCount Lib "kernel32" () As Long
    Private Declare Function GetWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal uCmd As GetWindowCmd) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As GetWindowLongEnum) As Long
    Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    Private Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
    Private Declare Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByVal hUf As Long) As Long
    Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As MouseEventFlags, ByVal dX As Long, ByVal dY As Long, ByVal dwData As Long, ByVal dwExtraInfo As Long)
    Private Declare Function NormalizeString Lib "Normaliz" (ByVal normForm As Long, ByVal lpSrcString As Long, ByVal cwSrcLength As Long, ByVal lpDstString As Long, ByVal cwDstLength As Long) As Long
    Private Declare Function ReleaseCapture Lib "user32" Alias "ReleaseCapture" () As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
    Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function SetLayeredWindowAttributes Lib "User32.dll" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwflags As Long) As Long
    Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As GetWindowLongEnum, ByVal dwNewLong As long) As Long
    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As hWndInsertAfterEnum, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As SetWindowPositionFlags) As Long
    Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As ShowWindowFlags) As Long
    Private Declare Function sndPlaySound32 Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As PlaySoundOptions) As Long
    Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
    Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
    Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
#End If


'CLASS METHODS
Private Sub Class_Initialize()
    With CreateObject("Scripting.FileSystemObject") 'Generate local folders for storing resources for this project
        If Not .FolderExists(.BuildPath(.GetSpecialFolder(2), "General-VBA")) Then .CreateFolder .BuildPath(.GetSpecialFolder(2), "General-VBA")
        If Not .FolderExists(.BuildPath(.GetSpecialFolder(2), "General-VBA\Resources")) Then .CreateFolder .BuildPath(.GetSpecialFolder(2), "General-VBA\Resources")
    End With
End Sub
'PRIVATE FUNCTIONS
'This method assures the required redirection for both the remote varType and the remote value at the same time thus removing any additional stack frames
'It can be used to both read from and write to memory by swapping the order of the last 2 parameters
'Source: https://github.com/cristianbuse/VBA-MemoryTools
#If Win64 And VBA7 Then
    Private Sub RemoteAssign(RM As REMOTE_MEMORY, memAddress As LongPtr, remoteVT As Variant, _
                             ByVal newVT As VbVarType, targetVariable As Variant, NewValue As Variant)
#Else
    Private Sub RemoteAssign(RM As REMOTE_MEMORY, memAddress As Long, remoteVT As Variant, _
                             ByVal newVT As VbVarType, targetVariable As Variant, NewValue As Variant)
#End If
    RM.memValue = memAddress
    remoteVT = newVT
    targetVariable = NewValue
    remoteVT = vbEmpty 'Stop linking to remote address, for safety
End Sub

'This function exists only because VBA Classes cannot contain or pass public UDTs between each other, minimizing the number of class modules
#If Win64 And VBA7 Then
    Private Function CopyUDT(SourceUDT As LongPtr, DestinationUDT As LongPtr, ByteLen As Long) As Boolean
#Else
    Private Function CopyUDT(SourceUDT As Long, DestinationUDT As Long, ByteLen As Long) As Boolean
#End If
    If Not LenB(SourceUDT) = ByteLen And LenB(DestinationUDT) = ByteLen Then Exit Function
    Dim Index As Long
    For Index = 0 To ByteLen - 1
        MemByte(DestinationUDT + Index) = MemByte(SourceUDT + Index)
    Next Index
    CopyUDT = (Index = ByteLen)
End Function



'PUBLIC FUNCTIONS
Public Function AssocQuery(FileExtension As String, Assoc As ASSOCSTR, Optional flags As ASSOCF, Optional sExtra As String = "open") As String
    Dim Buffer As String: Buffer = String(1024, vbNullChar)
    Dim BufferSize As Long: BufferSize = Len(Buffer)

    Call AssocQueryStringW(flags, Assoc, StrPtr(FileExtension), StrPtr(sExtra), StrPtr(Buffer), BufferSize)
    AssocQuery = Left(Buffer, BufferSize - 1)
End Function

#If Win64 And VBA7 Then
    Public Function AtlAxGetControl_(ByVal hwnd As LongPtr, Unk As stdole.IUnknown) As Long
#Else
    Public Function AtlAxGetControl_(ByVal hwnd As Long, Unk As stdole.IUnknown) As Long
#End If
    AtlAxGetControl_ = AtlAxGetControl(hwnd, Unk)
End Function

Public Function AtlAxWinInit_() As Long: AtlAxWinInit_ = AtlAxWinInit: End Function

#If Win64 And VBA7 Then
    Public Function CallWindowProc_(ByVal lpPrevWndFunc As LongPtr, ByVal hwnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
    Public Function CallWindowProc_(ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
    CallWindowProc_ = CallWindowProc(lpPrevWndFunc, hwnd, Msg, wParam, lParam)
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

Public Function ColorHLSToRGB_(wHue As Integer, wLuminance As Integer, wSaturation As Integer) As Long
    ColorHLSToRGB_ = ColorHLSToRGB(wHue, wLuminance, wSaturation): End Function

Public Function ColorRGBToHLS_(clrRGB As Long, wHue As Integer, wLuminance As Integer, wSaturation As Integer) As Long
    ColorRGBToHLS_ = ColorRGBToHLS(clrRGB, wHue, wLuminance, wSaturation): End Function

#If Win64 And VBA7 Then
    Public Function CreateWindowEx_(dwExStyle As ExtendedWindowStyleFlags, lpClassName As String, lpWindowName As String, dwStyle As WindowStyleFlags, _
                                    X As Long, Y As Long, nWidth As Long, nHeight As Long, _
                                    HwndParent As LongPtr, hMenu As LongPtr, hInstance As LongPtr, ByVal lpParam As LongPtr) As LongPtr
#Else
    Public Function CreateWindowEx_(dwExStyle As ExtendedWindowStyleFlags, lpClassName As String, lpWindowName As String, dwStyle As WindowStyleFlags, _
                                    X As Long, Y As Long, nWidth As Long, nHeight As Long, _
                                    HwndParent As Long, hMenu As Long, hInstance As Long, ByVal lpParam As Long) As Long
#End If
    CreateWindowEx_ = CreateWindowEx(dwExStyle, lpClassName, lpWindowName, dwStyle, X, Y, nWidth, nHeight, HwndParent, hMenu, hInstance, lpParam)
End Function

#If Win64 And VBA7 Then
    Public Function DestroyWindow_(hwnd As LongPtr) As Long
#Else
    Public Function DestroyWindow_(hwnd As Long) As Long
#End If
    DestroyWindow_ = DestroyWindow(hwnd)
End Function

Public Function DownloadFile(SourceURL As String, LocalFile As String) As Boolean
    If Len(SourceURL) = 0 Or Len(LocalFile) = 0 Then Exit Function
    Call URLDownloadToFile(0, SourceURL, LocalFile, 0, 0)
    DownloadFile = CreateObject("Scripting.FileSystemObject").FileExists(LocalFile)
End Function

#If Win64 And VBA7 Then
    Public Function DrawMenuBar_(hwnd As LongPtr)
#Else
    Public Function DrawMenuBar_(hwnd As Long)
#End If
    DrawMenuBar_ = DrawMenuBar(hwnd)
End Function

'Returns TRUE if a process with a given name is successfully terminated
'NOTE: Terminates ALL instances of a given process name
Public Function EndProcess(Optional ProcessName As String) As Boolean
    Dim ProcList As Collection
    If Len(ProcessName) > 0 Then
        Set ProcList = GetProcessList(ProcessName)
    Else
        Set ProcList = GetProcessList()
    End If
    If ProcList Is Nothing Then Exit Function
    
    Dim Output As Long, Process As Object
    
    For Each Process In ProcList
        If CBool(Process.Terminate = 0) Then Output = Output + 1
    Next Process
    
    Set Process = Nothing
    Set ProcList = Nothing
    EndProcess = Output > 0
End Function

Public Function ExtractIcon_(hInst As Long, lpszExeFileName As String, nIconIndex As Long) As Long
    ExtractIcon_ = ExtractIcon(hInst, lpszExeFileName, nIconIndex)
End Function

#If Win64 And VBA7 Then
    Public Function FindExecutable_(lpFile As String, lpDirectory As String, lpResult As String) As LongPtr
#Else
    Public Function FindExecutable_(lpFile As String, lpDirectory As String, lpResult As String) As Long
#End If
    FindExecutable_ = FindExecutable(lpFile, lpDirectory, lpResult)
End Function

#If Win64 And VBA7 Then
    Public Function FindWindow_(lpClassName As String, lpWindowName As String) As LongPtr
#Else
    Public Function FindWindow_(lpClassName As String, lpWindowName As String) As Long
#End If
    FindWindow_ = FindWindow(lpClassName, lpWindowName)
End Function

#If Win64 And VBA7 Then
    Public Function FlashWindow_(hwnd As LongPtr, dwFlags As FlashWindowFlags) As Long
#Else
    Public Function FlashWindow_(hwnd As Long, dwFlags As FlashWindowFlags) As Long
#End If
    FlashWindow_ = FlashWindow(hwnd, dwFlags)
End Function

Public Function GetAsyncKeyState_(VKey As VirtualKeyCode) As Long
    GetAsyncKeyState_ = GetAsyncKeyState(VKey)
End Function

#If Win64 And VBA7 Then
    Public Function GetCursorPosition(ReturnPointAPI As LongPtr) As Long
#Else
    Public Function GetCursorPosition(ReturnPointAPI As Long) As Long
#End If
    Dim Output As POINTAPI
    GetCursorPosition = GetCursorPos(Output)
    Call CopyUDT(VarPtr(Output), ReturnPointAPI, LenB(Output))
End Function

#If Win64 And VBA7 Then
    Public Function GetDC_(hwnd As LongPtr) As LongPtr
#Else
    Public Function GetDC_(hwnd As Long) As Long
#End If
    GetDC_ = GetDC(hwnd)
End Function

#If Win64 And VBA7 Then
    Public Function GetDeviceCapabilities(ByVal hDC As LongPtr, ByVal nIndex As GetDeviceCapsEnum) As Long
#Else
    Public Function GetDeviceCapabilities(ByVal hDC As Long, ByVal nIndex As GetDeviceCapsEnum) As Long
#End If
    GetDeviceCapabilities = GetDeviceCaps(hDC, nIndex)
End Function

Public Function GetEXE(FileName As String) As String
    Dim Output As String: Output = Space(255)
    Call FindExecutable_(FileName, vbNullString, Output)
    GetEXE = Left(Trim(Output), Len(Trim(Output)) - 1)
End Function

#If Win64 And VBA7 Then
    Public Function GetFocus_() As LongPtr
#Else
    Public Function GetFocus_() As Long
#End If
    GetFocus_ = GetFocus
End Function

#If Win64 And VBA7 Then
    Public Function GetHandleFromPoint(PointPtr As LongPtr) As LongPtr
    Dim Point As LongLong
    Call CopyUDT(PointPtr, VarPtr(Point), LenB(Point))
    GetHandleFromPoint = WindowFromPoint(Point)
#Else
    Public Function GetHandleFromPoint(PointPtr As Long) As Long
    Dim Point As POINTAPI
    Call CopyUDT(PointPtr, VarPtr(Point), LenB(Point))
    GetHandleFromPoint = WindowFromPoint(Point.X, Point.Y)
#End If
End Function

#If Win64 And VBA7 Then
    Public Function GetModuleHandle_(lpModuleName As String) As LongPtr
#Else
    Public Function GetModuleHandle_(lpModuleName As String) As Long
#End If
    GetModuleHandle_ = GetModuleHandle(lpModuleName)
End Function

#If Win64 And VBA7 Then
    Public Function GetNextWindow_(hwnd As LongPtr, wFlag As Long) As LongPtr
#Else
    Public Function GetNextWindow_(hwnd As Long, wFlag As Long) As Long
#End If
    GetNextWindow_ = GetNextWindow(hwnd, wFlag)
End Function

#If Win64 And VBA7 Then
    Public Function GetPixel_(hDC As LongPtr, X As Long, Y As Long) As Long
#Else
    Public Function GetPixel_(hDC As Long, X As Long, Y As Long) As Long
#End If
    GetPixel_ = GetPixel(hDC, X, Y)
End Function

'Returns a Collection of Processes based on a given Process name
'Returns a Collection of All Processes if no Process name is given
Public Function GetProcessList(Optional ProcessName As String) As Collection
    Dim WhereArg As String: WhereArg = IIf(Len(ProcessName) > 0, " WHERE Name LIKE '" & ProcessName & "'", vbNullString)
    Dim WIN_Services As Object: Set WIN_Services = GetObject("winmgmts:\\.\root\CIMV2")
    If WIN_Services Is Nothing Then Exit Function
    
    Dim WIN_ProcessSet As Object: Set WIN_ProcessSet = WIN_Services.ExecQuery("Select * FROM Win32_Process" & WhereArg, , 48)
    If WIN_ProcessSet Is Nothing Then Exit Function
    
    Dim Process As Object
    Set GetProcessList = New Collection
    For Each Process In WIN_ProcessSet
        GetProcessList.Add Process
    Next
    
    Set Process = Nothing
    Set WIN_ProcessSet = Nothing
    Set WIN_Services = Nothing
End Function

#If Win64 And VBA7 Then
    Public Function GetTick() As LongPtr
    GetTick = GetTickCount64
#Else
    Public Function GetTick() As Long
    GetTick = GetTickCount
#End If
End Function

#If Win64 And VBA7 Then
    Public Function GetWindow_(hwnd As LongPtr, uCmd As GetWindowCmd) As LongPtr
#Else
    Public Function GetWindow_(hwnd As LongPtr, uCmd As GetWindowCmd) As Long
#End If
    GetWindow_ = GetWindow(hwnd, uCmd)
End Function

#If Win64 And VBA7 Then
    Public Function GetWindowLong_(hwnd As LongPtr, nIndex As GetWindowLongEnum) As LongPtr
    GetWindowLong_ = GetWindowLongPtr(hwnd, nIndex)
#Else
    Public Function GetWindowLong_(hwnd As Long, nIndex As GetWindowLongEnum) As Long
    GetWindowLong_ = GetWindowLong(hwnd, nIndex)
#End If
End Function

#If Win64 And VBA7 Then
    Public Function GetWindowRect_(hwnd As LongPtr, lpRect As LongPtr) As Long
#Else
    Public Function GetWindowRect_(hwnd As Long, lpRect As Long) As Long
#End If
    Dim Output As RECT
    GetWindowRect_ = GetWindowRect(hwnd, Output)
    Call CopyMemory(ByVal lpRect, ByVal VarPtr(Output), LenB(Output))
End Function

#If Win64 And VBA7 Then
    Public Function GetUFHandle(ByVal TargetUserform As MSForms.UserForm) As LongPtr
#Else
    Public Function GetUFHandle(ByVal TargetUserform As MSForms.UserForm) As Long
#End If
    GetUFHandle = FindWindow("ThunderDFrame", TargetUserform.Caption)
End Function

#If Win64 And VBA7 Then
    Public Function IntersectRect_(lpDestRect As LongPtr, lpSrc1Rect As LongPtr, lpSrc2Rect As LongPtr) As Long
#Else
    Public Function IntersectRect_(lpDestRect As Long, lpSrc1Rect As Long, lpSrc2Rect As Long) As Long
#End If
    Dim Rect1 As RECT, Rect2 As RECT, DEST As RECT
    Call CopyUDT(lpSrc1Rect, VarPtr(Rect1), LenB(Rect1))
    Call CopyUDT(lpSrc2Rect, VarPtr(Rect2), LenB(Rect2))
    IntersectRect_ = IntersectRect(DEST, Rect1, Rect2)
    Call CopyUDT(VarPtr(DEST), lpDestRect, LenB(DEST))
End Function

'Read/Write a Byte from/to memory
'Source: https://github.com/cristianbuse/VBA-MemoryTools
#If Win64 And VBA7 Then
    Public Property Get MemByte(ByVal memAddress As LongPtr) As Byte
#Else
    Public Property Get MemByte(ByVal memAddress As Long) As Byte
#End If
    Static RM As REMOTE_MEMORY
    If Not RM.isInitialized Then 'Link .remoteVt to .memValue's first 2 bytes
        RM.remoteVT = VarPtr(RM.memValue): Call CopyMemory(RM.remoteVT, vbInteger + VT_BYREF, 2): RM.isInitialized = True
    End If
    Call RemoteAssign(RM, memAddress, RM.remoteVT, vbByte + VT_BYREF, MemByte, RM.memValue)
End Property
#If Win64 And VBA7 Then
    Public Property Let MemByte(ByVal memAddress As LongPtr, ByVal NewValue As Byte)
#Else
    Public Property Let MemByte(ByVal memAddress As Long, ByVal NewValue As Byte)
#End If
    Static RM As REMOTE_MEMORY
    If Not RM.isInitialized Then 'Link .remoteVt to .memValue's first 2 bytes
        RM.remoteVT = VarPtr(RM.memValue): Call CopyMemory(RM.remoteVT, vbInteger + VT_BYREF, 2): RM.isInitialized = True
    End If
    Call RemoteAssign(RM, memAddress, RM.remoteVT, vbByte + VT_BYREF, RM.memValue, NewValue)
End Property

#If Win64 And VBA7 Then
    Public Sub MouseEvent(ByVal dwFlags As MouseEventFlags, ByVal dX As Long, ByVal dY As Long, ByVal dwData As Long, ByVal dwExtraInfo As LongPtr)
#Else
    Public Sub MouseEvent(ByVal dwFlags As MouseEventFlags, ByVal dX As Long, ByVal dY As Long, ByVal dwData As Long, ByVal dwExtraInfo As Long)
#End If
    Call mouse_event(dwFlags, dX, dY, dwData, dwExtraInfo)
End Sub

'Normalizes a given unicode string to W3C standards (https://unicode.org/reports/tr15/)
Public Function NormalizeUnicode(RawString As String, Optional Normalization As NormalizationForm = NormalizationC) As String
    If Len(RawString) = 0 Then Exit Function
    NormalizeUnicode = String(Len(RawString), 0)
    Call NormalizeString(Normalization, StrPtr(RawString), Len(RawString), StrPtr(NormalizeUnicode), Len(NormalizeUnicode))
End Function

Public Function PixelsFromTwipsX(Twips As Long) As Long
    #If Win64 And VBA7 Then
        Dim hDC As LongPtr: hDC = GetDC(0)
    #Else
        Dim hDC As Long: hDC = GetDC(0)
    #End If
    Const TWIPSPERINCH = 1440
    PixelsFromTwipsX = (Twips / TWIPSPERINCH) * GetDeviceCaps(hDC, LOGPIXELSX)
    Call ReleaseDC(0, hDC)
End Function

Public Function PixelsFromTwipsY(Twips As Long) As Long
    #If Win64 And VBA7 Then
        Dim hDC As LongPtr: hDC = GetDC(0)
    #Else
        Dim hDC As Long: hDC = GetDC(0)
    #End If
    Const TWIPSPERINCH = 1440
    PixelsFromTwipsY = (Twips / TWIPSPERINCH) * GetDeviceCaps(hDC, LOGPIXELSY)
    Call ReleaseDC(0, hDC)
End Function

Public Function PlaySound(lpszSoundName As String, Optional uFlags As PlaySoundOptions) As Long
    PlaySound = sndPlaySound32(lpszSoundName, uFlags)
End Function

'Returns the string stored at a given location in memory (Inverse of StrPtr() function)
#If Win64 And VBA7 Then
    Public Function PtrStr(StrPointer As LongPtr) As String
#Else
    Public Function PtrStr(StrPointer As Long) As String
#End If
    If StrPointer = 0 Then Exit Function
    Dim Offset As Long, Buffer(1) As Byte, BufferStr As String
    Do
        Call CopyMemory(Buffer(0), ByVal StrPointer + Offset, 2)
        BufferStr = IIf(Not (Buffer(0) = 0 And Buffer(1) = 0), Buffer, vbNullString)
        PtrStr = PtrStr & BufferStr
        Offset = Offset + 2
    Loop Until Buffer(0) = 0 And Buffer(1) = 0
End Function

Public Function ReleaseCapture_() As Long
    ReleaseCapture_ = ReleaseCapture
End Function

#If Win64 And VBA7 Then
    Public Function ReleaseDC_(hwnd As LongPtr, hDC As LongPtr) As Long
#Else
    Public Function ReleaseDC_(hwnd As Long, hDC As Long) As Long
#End If
    ReleaseDC_ = ReleaseDC(hwnd, hDC)
End Function

#If Win64 And VBA7 Then
    Public Function SendMessage_(hwnd As LongPtr, wMsg As Long, wParam As Integer, lParam As Long) As Long
#Else
    Public Function SendMessage_(hwnd As Long, wMsg As Long, wParam As Integer, lParam As Long) As Long
#End If
    SendMessage_ = SendMessage(hwnd, wMsg, wParam, lParam)
End Function

Public Function SetCursorPosition(X As Long, Y As Long) As Long
    SetCursorPosition = SetCursorPos(X, Y)
End Function

#If Win64 And VBA7 Then
    Public Function SetForegroundWindow_(hwnd As LongPtr) As Long
#Else
    Public Function SetForegroundWindow_(hwnd As Long) As Long
#End If
    SetForegroundWindow_ = SetForegroundWindow(hwnd)
End Function

#If Win64 And VBA7 Then
    Public Function SetLayeredWindowAttributes_(hwnd As LongPtr, crKey As Long, bAlpha As Byte, dwFlags As Long) As Long
#Else
    Public Function SetLayeredWindowAttributes_(hwnd As Long, crKey As Long, bAlpha As Byte, dwFlags As Long) As Long
#End If
    SetLayeredWindowAttributes_ = SetLayeredWindowAttributes(hwnd, crKey, bAlpha, dwFlags)
End Function

#If Win64 And VBA7 Then
    Public Function SetProp_(ByVal hwnd As LongPtr, ByVal lpString As String, ByVal hData As LongPtr) As Long
#Else
    Public Function SetProp_(ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
#End If
    SetProp_ = SetProp(hwnd, lpString, hData)
End Function

#If Win64 And VBA7 Then
    Public Function SetWindowLong_(hwnd As LongPtr, nIndex As GetWindowLongEnum, dwNewLong As LongPtr) As LongPtr
    SetWindowLong_ = SetWindowLongPtr(hwnd, nIndex, dwNewLong)
#Else
    Public Function SetWindowLong_(hwnd As Long, nIndex As GetWindowLongEnum, dwNewLong As Long) As Long
    SetWindowLong_ = SetWindowLong(hwnd, nIndex, dwNewLong)
#End If
End Function

#If Win64 And VBA7 Then
    Public Function SetWindowPos_(hwnd As LongPtr, hWndInsertAfter As hWndInsertAfterEnum, X As Long, Y As Long, cX As Long, cY As Long, wFlags As SetWindowPositionFlags) As Boolean
#Else
    Public Function SetWindowPos_(hwnd As Long, hWndInsertAfter As hWndInsertAfterEnum, X As Long, Y As Long, cX As Long, cY As Long, wFlags As SetWindowPositionFlags) As Boolean
#End If
    SetWindowPos_ = SetWindowPos(hwnd, hWndInsertAfter, X, Y, cX, cY, wFlags)
End Function

#If Win64 And VBA7 Then
    Public Function SetWindowText_(hwnd As LongPtr, lpString As String) As Long
#Else
    Public Function SetWindowText_(hwnd As Long, lpString As String) As Long
#End If
    SetWindowText_ = SetWindowText(hwnd, lpString)
End Function

#If Win64 And VBA7 Then
    Public Function ShowWindow_(hwnd As LongPtr, nCmdShow As ShowWindowFlags) As Long
#Else
    Public Function ShowWindow_(hwnd As Long, nCmdShow As ShowWindowFlags) As Long
#End If
    ShowWindow_ = ShowWindow(hwnd, nCmdShow)
End Function

#If Win64 And VBA7 Then
    Public Function SystemParametersInfo_(uAction As Long, uParam As Long, ByVal lpvParam As LongPtr, fuWinIni As Long) As Long
#Else
    Public Function SystemParametersInfo_(uAction As Long, uParam As Long, ByVal lpvParam As Long, fuWinIni As Long) As Long
#End If
    Dim UDTInput As NONCLIENTMETRICS
    Call CopyMemory(ByVal (VarPtr(UDTInput)), ByVal lpvParam, LenB(UDTInput))
    SystemParametersInfo_ = SystemParametersInfo(uAction, uParam, UDTInput, fuWinIni)
    Call CopyMemory(ByVal lpvParam, ByVal VarPtr(UDTInput), LenB(UDTInput))
End Function

Public Function TwipsPerPixelX(ByVal Pixels As Long) As Long
    #If Win64 And VBA7 Then
        Dim hDC As LongPtr: hDC = GetDC(0)
    #Else
        Dim hDC As Long: hDC = GetDC(0)
    #End If
    Const TWIPSPERINCH = 1440
    TwipsPerPixelX = Pixels * (TWIPSPERINCH / GetDeviceCaps(hDC, LOGPIXELSX))
    Call ReleaseDC(0, hDC)
End Function

Public Function TwipsPerPixelY(ByVal Pixels As Long) As Long
    #If Win64 And VBA7 Then
        Dim hDC As LongPtr: hDC = GetDC(0)
    #Else
        Dim hDC As Long: hDC = GetDC(0)
    #End If
    Const TWIPSPERINCH = 1440
    TwipsPerPixelY = Pixels * (TWIPSPERINCH / GetDeviceCaps(hDC, LOGPIXELSY))
    Call ReleaseDC(0, hDC)
End Function
