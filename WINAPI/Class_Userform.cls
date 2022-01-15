VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class_Userform"
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

'SetWindowPos Constants
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOP = 0

'SystemParametersInfo Constants
Private Const SPI_GETNONCLIENTMETRICS = 41


'Notification Icon Styles
Public Enum KeyFlags
    KEY_SHIFT = &H1
    KEY_CTRL = &H2
    KEY_ALT = &H4
End Enum

Public Enum MButtonFlags
    FMBUTTONLEFT = &H1
    FMBUTTONRIGHT = &H2
    FMBUTTONMIDDLE = &H4
End Enum

Public Enum NotificationIconEnum
    NoIcon
    InfoIcon
    WarningIcon
    ErrorIcon
End Enum


'TYPE DECLARATIONS
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type SIZE
    cX As Long
    cY As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

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



'EXTERNAL VARIABLES
Public WINAPI As New Class_WINAPI

'INTERNAL VARIABLES
Private WithEvents UForm As MSForms.UserForm    'Userform Events Object
Attribute UForm.VB_VarHelpID = -1

#If Win64 And VBA7 Then
    Private Handle As LongPtr   'LongPtr Handle of the Userform
    Private BHandle As LongPtr  'LongPtr Handle of the WebBrowser (Gif Icon)
    Private WStyle As LongPtr   'LongPtr Window style flag variable
    Private EXStyle As LongPtr  'LongPtr Extended window style flag variable
#Else
    Private Handle As Long      'Long Handle of the Userform
    Private BHandle As Long     'Long Handle of the WebBrowser (Gif Icon)
    Private WStyle As Long      'Long Window style flag variable
    Private EXStyle As Long     'Long Extended window style flag variable
#End If

Private Form As Object          'Userform Event Action Object
Private Browser As Object       'WebBrowser Object (Gif Icon)
Private lShiftLeftPix As Long   'WebBrowser X Offset (Gif Icon)
Private lShiftTopPix As Long    'WebBrowser Y Offset (Gif Icon)
Private IconHandle As Long      'Long Handle of the Userform Icon
Private Transparency As Long    'Userform's Transparency setting
Private Translucency As Byte    'Userform's Translucency setting
Private BlankIconPath As String 'Filepath of a given icon file

'Property flags
Private IconEnabled_ As Boolean
Private WindowBorder_ As Boolean
Private WindowCaption_ As Boolean
Private WindowEdge_ As Boolean
Private MaximizeBox_ As Boolean
Private MinimizeBox_ As Boolean
Private Resizable_ As Boolean
Private SystemMenu_ As Boolean
Private Taskbar_ As Boolean
Private TopMost_ As Boolean
Private TooltipWindow_ As Boolean
Private Draggable_ As Boolean


'CLASS PROPERTIES
Public Property Get IconEnabled() As Boolean
    HasIcon = IconEnabled_
End Property

#If Win64 And VBA7 Then
    Public Property Get hwnd() As LongPtr
#Else
    Public Property Get hwnd() As Long
#End If
    hwnd = Handle
End Property

Public Property Get WindowBorder() As Boolean
    WindowBorder = WindowBorder_
End Property
Public Property Let WindowBorder(Value As Boolean)
    WindowBorder_ = Value
    If WindowBorder_ Then Call AddWindowStyle(WS_BORDER)
    If Not WindowBorder_ Then Call RemoveWindowStyle(WS_BORDER)
End Property

Public Property Get WindowCaption() As Boolean
    WindowCaption = WindowCaption_
End Property
Public Property Let WindowCaption(Value As Boolean)
    WindowCaption_ = Value
    If WindowCaption_ Then Call AddWindowStyle(WS_CAPTION)
    If Not WindowCaption_ Then Call RemoveWindowStyle(WS_CAPTION)
End Property

Public Property Get WindowEdge() As Boolean
    WindowEdge = WindowEdge_
End Property
Public Property Let WindowEdge(Value As Boolean)
    WindowEdge_ = Value
    If WindowEdge_ Then Call AddExtWindowStyle(WS_EX_WINDOWEDGE)
    If Not WindowEdge_ Then Call RemoveExtWindowStyle(WS_EX_WINDOWEDGE)
End Property

Public Property Get MaximizeBox() As Boolean
    MaximizeBox = MaximizeBox_
End Property
Public Property Let MaximizeBox(Value As Boolean)
    MaximizeBox_ = Value
    If MaximizeBox_ Then Call AddWindowStyle(WS_MAXIMIZEBOX)
    If Not MaximizeBox_ Then Call RemoveWindowStyle(WS_MAXIMIZEBOX)
End Property

Public Property Get MinimizeBox() As Boolean
    MinimizeBox = MinimizeBox_
End Property
Public Property Let MinimizeBox(Value As Boolean)
    MinimizeBox_ = Value
    If MinimizeBox_ Then Call AddWindowStyle(WS_MINIMIZEBOX)
    If Not MinimizeBox_ Then Call RemoveWindowStyle(WS_MINIMIZEBOX)
End Property

Public Property Get Resizable() As Boolean
    Resizable = Resizable_
End Property
Public Property Let Resizable(Value As Boolean)
    Resizable_ = Value
    If Resizable_ Then Call AddWindowStyle(WS_SIZEBOX)
    If Not Resizable_ Then Call RemoveWindowStyle(WS_SIZEBOX)
End Property

Public Property Get SystemMenu() As Boolean
    SystemMenu = SystemMenu_
End Property
Public Property Let SystemMenu(Value As Boolean)
    SystemMenu_ = Value
    If SystemMenu_ Then Call AddWindowStyle(WS_SYSMENU)
    If Not SystemMenu_ Then Call RemoveWindowStyle(WS_SYSMENU)
End Property

Public Property Get Taskbar() As Boolean
    Taskbar = Taskbar_
End Property
Public Property Let Taskbar(Value As Boolean)
    Taskbar_ = Value
    If Taskbar_ Then Call AddExtWindowStyle(WS_EX_APPWINDOW)
    If Not Taskbar_ Then Call RemoveExtWindowStyle(WS_EX_APPWINDOW)
    'Reset the TopMost state due to WS_EX_APPWINDOW adversely affecting it
    Call WINAPI.SetWindowPos_(Handle, IIf(TopMost_, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Property

Public Property Get TopMost() As Boolean
    TopMost = TopMost_
End Property
Public Property Let TopMost(Value As Boolean)
    TopMost_ = Value
    Call WINAPI.SetWindowPos_(Handle, IIf(TopMost_, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Property

Public Property Get TooltipWindow() As Boolean
    TooltipWindow = TooltipWindow_
End Property
Public Property Let TooltipWindow(Value As Boolean)
    TooltipWindow_ = Value
    If TooltipWindow_ Then Call AddExtWindowStyle(WS_EX_TOOLWINDOW)
    If Not TooltipWindow_ Then Call RemoveExtWindowStyle(WS_EX_TOOLWINDOW)
End Property

Public Property Get Draggable() As Boolean
    Draggable = Draggable_
End Property
Public Property Let Draggable(Value As Boolean)
    Draggable_ = Value
End Property


'CLASS METHODS
Private Sub Class_Initialize()
End Sub
Private Sub Class_Terminate()
    If BHandle > 0 Then Call WINAPI.DestroyWindow_(BHandle)
    Set Browser = Nothing
    Set Form = Nothing
    Set UForm = Nothing
    Set WINAPI = Nothing
End Sub


'USERFORM EVENTS
Private Sub UForm_Layout()
    Call UpdateIcon
End Sub
Public Sub RefreshIcon(Toggle As Boolean)
    If Not Toggle Then Exit Sub
    Call UpdateIcon
End Sub
Private Sub UForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Const WM_NCLBUTTONDOWN = &HA1
    Const HTCAPTION = 2
    If Button = FMBUTTONLEFT And Draggable_ Then
        Call WINAPI.ReleaseCapture_
        Call WINAPI.SendMessage_(Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If
End Sub


'PUBLIC SUB/FUNCTIONS
'Sets the given userform as the target for this class
Public Function SetUserform(TargetUserform As Object) As Boolean
    If TargetUserform Is Nothing Then Exit Function
    Set UForm = TargetUserform 'Set Event Listener
    Set Form = TargetUserform
    Handle = WINAPI.GetUFHandle(TargetUserform) 'Get Form Handle
    
    'Load Window Styling
    WStyle = WINAPI.GetWindowLong_(Handle, GWL_STYLE)
    EXStyle = WINAPI.GetWindowLong_(Handle, GWL_EXSTYLE)
    
    'Apply Window Style bit flags to class properties
    WindowBorder_ = (WStyle And WS_BORDER) = WS_BORDER
    WindowCaption_ = (WStyle And WS_CAPTION) = WS_CAPTION
    MaximizeBox_ = (WStyle And WS_MAXIMIZEBOX) = WS_MAXIMIZEBOX
    MinimizeBox_ = (WStyle And WS_MINIMIZEBOX) = WS_MINIMIZEBOX
    Resizable_ = (WStyle And WS_SIZEBOX) = WS_SIZEBOX
    SystemMenu_ = (WStyle And WS_SYSMENU) = WS_SYSMENU
    Taskbar_ = (EXStyle And WS_EX_APPWINDOW) = WS_EX_APPWINDOW
    TopMost_ = (EXStyle And WS_EX_TOPMOST) = WS_EX_TOPMOST
    TooltipWindow_ = (EXStyle And WS_EX_TOOLWINDOW) = WS_EX_TOOLWINDOW
    
    SetUserform = (Not Handle = 0)
End Function

'Adds window styles to the userform
Public Function AddWindowStyle(WindowStyle As WindowStyleFlags) As Boolean
    If Handle = 0 Then Exit Function
    WStyle = WINAPI.GetWindowLong_(Handle, GWL_STYLE) Or WindowStyle
    AddWindowStyle = WINAPI.SetWindowLong_(Handle, GWL_STYLE, WStyle)
    Call WINAPI.DrawMenuBar_(Handle)
End Function
'Removes window styles from the userform
Public Function RemoveWindowStyle(WindowStyle As WindowStyleFlags) As Boolean
    If Handle = 0 Then Exit Function
    WStyle = WINAPI.GetWindowLong_(Handle, GWL_STYLE) And Not WindowStyle
    RemoveWindowStyle = WINAPI.SetWindowLong_(Handle, GWL_STYLE, WStyle)
    Call WINAPI.DrawMenuBar_(Handle)
End Function

'Adds extended window styles to the userform
Public Function AddExtWindowStyle(ExtendedWindowStyle As ExtendedWindowStyleFlags) As Boolean
    If Handle = 0 Then Exit Function
    EXStyle = WINAPI.GetWindowLong_(Handle, GWL_EXSTYLE) Or ExtendedWindowStyle
    
    Call WINAPI.ShowWindow_(Handle, SW_HIDE)
    AddExtWindowStyle = WINAPI.SetWindowLong_(Handle, GWL_EXSTYLE, EXStyle)
    Call WINAPI.ShowWindow_(Handle, SW_SHOW)
    Call WINAPI.DrawMenuBar_(Handle)
End Function
'Removes extended window styles from the userform
Public Function RemoveExtWindowStyle(ExtendedWindowStyle As ExtendedWindowStyleFlags) As Boolean
    If Handle = 0 Then Exit Function
    EXStyle = WINAPI.GetWindowLong_(Handle, GWL_EXSTYLE) And Not ExtendedWindowStyle
    
    Call WINAPI.ShowWindow_(Handle, SW_HIDE)
    RemoveExtWindowStyle = WINAPI.SetWindowLong_(Handle, GWL_EXSTYLE, EXStyle)
    Call WINAPI.ShowWindow_(Handle, SW_SHOW)
    Call WINAPI.DrawMenuBar_(Handle)
End Function

'Sets the Transparency bit of a form's style flag
Public Function SetTransparency(Enable As Boolean, Optional Color As Long) As Boolean
    If Handle = 0 Then Exit Function
    Transparency = Color
    EXStyle = WINAPI.GetWindowLong_(Handle, GWL_EXSTYLE)
    EXStyle = IIf(Enable, EXStyle Or WS_EX_LAYERED, EXStyle And Not WS_EX_LAYERED)
    SetTransparency = WINAPI.SetWindowLong_(Handle, GWL_EXSTYLE, EXStyle)
    If SetTransparency And Enable Then SetTransparency = WINAPI.SetLayeredWindowAttributes_(Handle, Transparency, Translucency, LWA_COLORKEY)
End Function

'Sets the Translucency bit of a form's style flag
Public Function SetTranslucency(Enable As Boolean, Optional ByVal Alpha As Byte) As Boolean
    If Handle = 0 Then Exit Function
    Translucency = Alpha
    EXStyle = WINAPI.GetWindowLong_(Handle, GWL_EXSTYLE) Or WS_EX_LAYERED
    EXStyle = IIf(Enable, EXStyle Or WS_EX_LAYERED, EXStyle And Not WS_EX_LAYERED)
    If WINAPI.SetWindowLong_(Handle, GWL_EXSTYLE, EXStyle) Then SetTranslucency = WINAPI.SetLayeredWindowAttributes_(Handle, Transparency, Translucency, LWA_ALPHA)
End Function

Public Function Show() As Boolean
    If Handle = 0 Then Exit Function
    Show = WINAPI.SetForegroundWindow_(Handle)
End Function

Public Function SetIcon(ByVal IconPath As String, Optional LargeIcon As Boolean = True) As Boolean
    If Handle = 0 Then Exit Function
    IconHandle = IIf(CreateObject("Scripting.FileSystemObject").FileExists(IconPath), WINAPI.ExtractIcon_(0, IconPath, 0), 0)
    SetIcon = Not WINAPI.SendMessage_(Handle, WM_SETICON, IIf(LargeIcon, ICON_BIG, ICON_SMALL), IconHandle)
    IconEnabled_ = SetIcon And (Not IconHandle = 0)
    If (Not IconEnabled_) And (BHandle > 0) Then Call WINAPI.ShowWindow_(BHandle, SW_HIDE)
End Function

Public Function SetGifIcon(Optional ByVal FilePath As String) As Boolean
    If Len(FilePath) > 0 Then
        If Not CreateObject("Scripting.FileSystemObject").FileExists(FilePath) Then Exit Function
        If CreateBrowser Then
            Browser.Document.Body.innerHTML = "<img style=""position:absolute;top:0px;left:0px;width:20px;height:20px"" src=""" & FilePath & """>"
            If CreateIcon Then Call SetIcon(BlankIconPath)
            Call UpdateIcon 'Updates the icon appearance
            SetGifIcon = CBool(BHandle)
        End If
    Else
        'Hide the browser window, and remove the form's icon
        SetGifIcon = WINAPI.ShowWindow_(BHandle, SW_HIDE)
        Call SetIcon("")
    End If
End Function

Public Function Flash(Optional ByVal FWFlag As FlashWindowFlags = FLASHW_ALL) As Boolean
    If Handle = 0 Then Exit Function
    Flash = WINAPI.FlashWindow_(Handle, FWFlag)
End Function

Public Function Notify(Title As String, Message As String, Optional ByVal NotificationIcon As NotificationIconEnum, Optional ByVal App As String = "Excel", Optional ByVal Duration As Integer = 10)
    Dim Icon As String
    Select Case NotificationIcon
        Case 0: Icon = "None"
        Case 1: Icon = "Info"
        Case 2: Icon = "Warning"
        Case 3: Icon = "Error"
        Case Else: Icon = "None"
    End Select
    
    Dim PSStr As String: PSStr = """powershell.exe"" -Command " & Chr(34) & "& { " & _
    "Add-Type -AssemblyName 'System.Windows.Forms'" & _
    "; $notification = New-Object System.Windows.Forms.NotifyIcon" & _
    "; $path = (Get-Process -id (get-process " & App & ").id).Path" & _
    "; $notification.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)" & _
    "; $notification.BalloonTipIcon  = [System.Windows.Forms.ToolTipIcon]::" & Icon & "" & _
    "; $notification.BalloonTipText = '" & Message & "'" & _
    "; $notification.BalloonTipTitle = '" & Title & "'" & _
    "; $notification.Visible = $true" & _
    "; $notification.ShowBalloonTip(" & Duration & ")" & _
    " }" & Chr(34)
    Notify = CreateObject("WScript.Shell").Run(PSStr, 0, False)
End Function


'PRIVATE SUB/FUNCTIONS
'Creates a 16x16 transparent icon for use in this module. ICO is converted to byte() and saved as Blank.ico
'Returns TRUE if the icon exists
Private Function CreateIcon() As Boolean
    Const ICO As String = "0000010001001010020001000100B0000000160000002800000010000000200000000100010000000000000000000000000000000000000000000000000000000000FFFFFF0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000FFFF0000"
    With CreateObject("Scripting.FileSystemObject")
        Dim d(197) As Byte, I As Long, FP As String: FP = .BuildPath(.GetSpecialFolder(2), "General-VBA\Resources\Blank.ico")
        If Not .FileExists(FP) Then: For I = 0 To Len(ICO) - 1 Step 2: d(I / 2) = CByte("&H" & Mid(ICO, I + 1, 2)): Next I: I = FreeFile: Open FP For Binary Lock Read Write As I: Put I, , d: Close I
        BlankIconPath = FP: CreateIcon = .FileExists(FP)
    End With
End Function

'Creates the Browser for displaying the GIF as an icon
'Returns TRUE if the browser is created, or already exists
Private Function CreateBrowser() As Boolean
    If BHandle = 0 Then
        'This function initializes ATL's control hosting code by registering the "AtlAxWin80" and "AtlAxWinLic80" window classes
        'plus a couple of custom window messages.
        Call WINAPI.AtlAxWinInit_
        'CreateWindowEx - Creates an overlapped, pop-up, or child window with an extended window style; This function is identical to the CreateWindow function.
        'GetModuleHandle - Retrieves a module handle for the specified module. The module must have been loaded by the calling process.
        BHandle = WINAPI.CreateWindowEx_(0, "AtlAxWin", "about:blank", WS_POPUP Or WS_VISIBLE Or WS_DISABLED, 0, 0, 0, 0, 0, 0, WINAPI.GetModuleHandle_(vbNullString), VarPtr(0))
        If BHandle Then
            Call WINAPI.SetWindowLong_(BHandle, GWL_HWNDPARENT, Handle) 'Sets Handle as the Parent handle for BHandle
            'SetWindowPos_ - Set newly created Browser to temporarily be the topmost window
            'SetWindowLong_ - Set Extended window style properties to reflect the topmost z-order
            'SetWindowPos_ - Remove topmost z-order in case the userform is meant to be the topmost window
            Call WINAPI.SetWindowPos_(BHandle, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
            Call WINAPI.SetWindowLong_(BHandle, GWL_EXSTYLE, WINAPI.GetWindowLong_(BHandle, GWL_EXSTYLE) Or WS_EX_COMPOSITED Or WS_EX_LAYERED Or WS_EX_TRANSPARENT Or WS_EX_TOPMOST)
            Call WINAPI.SetWindowPos_(BHandle, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
            Call WINAPI.SetWindowPos_(Handle, IIf(TopMost_, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
            'Adds a new entry or changes an existing entry in the property list of the specified window.
            'Sets BHandle as a new property of Handle with name "BHandle"
            Call WINAPI.SetProp_(Handle, "BHandle", BHandle)
            'Obtains a direct interface pointer to the control contained inside a specified window given its handle.
            Dim Unk As stdole.IUnknown
            Call WINAPI.AtlAxGetControl_(BHandle, Unk)
            Set Browser = Unk: Set Unk = Nothing
            With Browser: .Silent = True: .FullScreen = True: DoEvents: End With
        End If
    End If
    CreateBrowser = Not Browser Is Nothing
End Function

'Returns the Hex color value for the browser background
Private Function GetRGB() As String
    Dim DC As LongPtr, FormRECT As RECT
    Dim Color As Long, Hue As Integer, Luminance As Integer, Saturation As Integer

    Call WINAPI.GetWindowRect_(Handle, VarPtr(FormRECT))
    DC = WINAPI.GetDC_(0)
    Color = WINAPI.GetPixel_(DC, FormRECT.Left + 50, FormRECT.Top + 5)
    Call WINAPI.ReleaseDC_(0, DC)
    Call WINAPI.ColorRGBToHLS_(Color, Hue, Luminance, Saturation)
    Color = WINAPI.ColorHLSToRGB_(Hue, Luminance, Saturation)
    GetRGB = "#" & Right("0" & Hex(&HFF& And Color), 2) & Right("0" & Hex((&HFF00& And Color) \ 256), 2) & Right("0" & Hex((&HFF0000 And Color) \ 65536), 2)
End Function

'Updates the position and background color of the browser window for the gif icon
Private Sub UpdateIcon()
    If Handle = 0 Or BHandle = 0 Then Exit Sub
    If IconEnabled_ Then
        Dim ChildRECT As RECT
        Dim NCM As NONCLIENTMETRICS
        Dim X As Long, Y As Long
        
        Call WINAPI.GetWindowRect_(WINAPI.GetNextWindow_(Handle, GW_CHILD), VarPtr(ChildRECT))
        NCM.cbSize = Len(NCM)
        If WINAPI.SystemParametersInfo_(SPI_GETNONCLIENTMETRICS, 0, VarPtr(NCM), 0) Then
            If lShiftLeftPix < 0 Then lShiftLeftPix = 0
            If lShiftLeftPix >= 5 Then lShiftLeftPix = 5
            If lShiftTopPix < -5 Then lShiftTopPix = -5
            If lShiftTopPix >= 5 Then lShiftTopPix = 5
            X = ChildRECT.Left + lShiftLeftPix + 5
            Y = ChildRECT.Top - NCM.iCaptionHeight - NCM.iPaddedBorderWidth + lShiftTopPix
            Call WINAPI.SetWindowPos_(BHandle, HWND_TOP, X, Y, NCM.iCaptionHeight - 2, NCM.iCaptionHeight - 2, SWP_NOACTIVATE Or SWP_SHOWWINDOW)
        End If
        
        If Not Browser Is Nothing Then Browser.Document.Body.bgcolor = GetRGB
        Call WINAPI.ShowWindow_(BHandle, SW_SHOWNORMAL) 'Brings the Browser to the front
    Else
        Call WINAPI.ShowWindow_(BHandle, SW_HIDE) 'Hides the Browser
    End If
    Call WINAPI.SetForegroundWindow_(Handle) 'Keeps Userform as the active window
End Sub
