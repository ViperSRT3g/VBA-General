Attribute VB_Name = "mod_Hotkeys"
Option Explicit

Private Type KBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type

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

Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
    ByVal hWnd1 As Long, _
    ByVal hWnd2 As Long, _
    ByVal lpsz1 As String, _
    ByVal lpsz2 As String) As Long

Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SetTimer Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal nIDEvent As Long, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As Long) As Long

Private Declare Function KillTimer Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal uIDEvent As Long) As Long

Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function EmptyClipboard Lib "user32" () As Long

Private Declare Function CloseClipboard Lib "user32" () As Long

Private Const WH_KEYBOARD_LL = 13
Private Const HC_ACTION = 0
Private Const GWL_HINSTANCE = (-6)
Private Const VK_SNAPSHOT = &H2C

Private blnHookEnabled As Boolean
Private hhkLowLevelKybd As Long
Private lWBhwnd As Long
Private lTimerID As Long
Private oNewApp As Application

'-------------------
'Public routines .
'-------------------
Public Sub SetHook()

    Dim owb As Workbook
    
    'Create an invisible xl instance
    'to run the keyboard hook callback.
    'this is to avoid accidental VBIDE crash.
    Set oNewApp = New Application
    
    oNewApp.EnableEvents = False
    Set owb = oNewApp.Workbooks.Open(ThisWorkbook.FullName, ReadOnly:=True)
    oNewApp.Run "Hook_KeyBoard", ThisWorkbook.FullName

End Sub

Public Sub RemoveHook()

    'Proper cleanup.
    If Not oNewApp Is Nothing Then
        oNewApp.Run "Unhook_KeyBoard"
        oNewApp.Quit
    End If

End Sub

Public Sub Hook_KeyBoard(WbPathName As String)

    Dim lXLhwnd, lXLHinstance As Long
    Dim owb As Workbook
    
    'get a temp pointer to calling WB.
    Set owb = GetObject(WbPathName)
    
    'store the hwnd of the calling WB.
    lWBhwnd = FindWindowEx(owb.Parent.hwnd, 0, "XLDESK", vbNullString)
    lWBhwnd = FindWindowEx(lWBhwnd, 0, "EXCEL7", owb.Name)
    
    'release ref to calling WB.
    Set owb = Nothing
    
    'hook the keyboard system wide.
    lXLhwnd = FindWindow("XLMAIN", Application.Caption)
    lXLHinstance = GetWindowLong(lXLhwnd, GWL_HINSTANCE)
    
    If blnHookEnabled = False Then
        hhkLowLevelKybd = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf LowLevelKeyboardProc, lXLHinstance, 0)
        blnHookEnabled = True
        'unhook the keyboard in case it wasn't properly
        'unhooked in the Before_Close event.
        lTimerID = SetTimer(0, 0, 1, AddressOf MonitorWorkBook)
    End If

End Sub

Public Sub Unhook_KeyBoard()

    If hhkLowLevelKybd <> 0 Then UnhookWindowsHookEx hhkLowLevelKybd
    hhkLowLevelKybd = 0
    blnHookEnabled = False

End Sub

'------------------
'Private routines.
'------------------
Private Function LowLevelKeyboardProc(ByVal nCode As Long, ByVal wParam As Long, ByRef lParam As KBDLLHOOKSTRUCT) As Long

    If nCode = HC_ACTION Then
        'never process the Print Key.
        If lParam.vkCode = VK_SNAPSHOT Then
            LowLevelKeyboardProc = -1
            SetTimer FindWindow("XLMAIN", Application.Caption), 0, 1, AddressOf ClearTheClipboard
            Exit Function
        End If
    End If
    
    LowLevelKeyboardProc = CallNextHookEx(0, nCode, wParam, ByVal lParam)

End Function

Private Sub ClearTheClipboard()

    'display msgbox only once.
 
    KillTimer FindWindow("XLMAIN", Application.Caption), 0
    If FindWindowEx(0, 0, "#32770", "Microsoft Excel") Then Exit Sub
    MsgBox "The Print Key is disabled.", vbSystemModal + vbInformation
 
    'Optionally clear the clipboard.
    OpenClipboard (0&)
    EmptyClipboard
    CloseClipboard
 
    'run any other code here...

End Sub

Private Sub MonitorWorkBook()
 
    'safety net in case the calling WB didn't
    'unhook the keyboard before closing.
    If IsWindow(lWBhwnd) = 0 Then
        DoEvents
        KillTimer 0, lTimerID
        Unhook_KeyBoard
        Application.Quit
    End If
 
End Sub
