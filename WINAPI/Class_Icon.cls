VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class_Icon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents UfEvents As MSForms.UserForm
Attribute UfEvents.VB_VarHelpID = -1

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type Size
    cx As Long
    cy As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const LF_FACESIZE = 32
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

#If VBA7 Then
    #If Win64 Then
        Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    #Else
        Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As LongPtr
    #End If
    Private Declare PtrSafe Function AtlAxWinInit Lib "Atl.dll" () As Long
    Private Declare PtrSafe Function AtlAxGetControl Lib "atl" (ByVal hwnd As LongPtr, Unk As stdole.IUnknown) As Long
    Private Declare PtrSafe Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr
    Private Declare PtrSafe Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal hData As LongPtr) As Long
    Private Declare PtrSafe Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As LongPtr, ByVal lpString As String) As LongPtr
    Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As LongPtr
    Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As Long
    Private Declare PtrSafe Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As LongPtr, ByVal wFlag As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare PtrSafe Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
    Private Declare PtrSafe Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDc As LongPtr, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hDc As LongPtr) As Long
    Private Declare PtrSafe Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByVal hUf As LongPtr) As Long
#Else
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function AtlAxWinInit Lib "Atl.dll" () As Long
    Private Declare Function AtlAxGetControl Lib "atl" (ByVal hwnd As Long, Unk As stdole.IUnknown) As Long
    Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
    Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
    Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
    Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
    Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    Private Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long
    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
    Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDc As Long) As Long
    Private Declare Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByVal hUf As Long) As Long
#End If


Private Sub Class_Terminate()
    Set UfEvents = Nothing
End Sub

Private Sub UfEvents_Layout()
    Call SetIconPos(UfEvents)
End Sub

Public Function AddAnimatedIcon(ByVal Form As Object, ByVal IconURL As String) As Boolean
    Set UfEvents = Form
    AddAnimatedIcon = CreateIcon(Form, IconURL)
End Function

Private Function CreateIcon(ByVal Form As Object, ByVal url As String) As Boolean

    Const GWL_HWNDPARENT = (-8)
    Const WS_VISIBLE = &H10000000
    Const WS_POPUP = &H80000000
    Const WS_DISABLED = &H8000000
    Const WS_EX_TOOLWINDOW = &H80

    #If Win64 Then
        Dim hWebCtrl As LongPtr, hParent As LongPtr, hDc As LongPtr
    #Else
        Dim hWebCtrl As Long, hParent As Long, hDc As Long
    #End If

    Dim oWbrowser  As Object, Unk As stdole.IUnknown
    Dim tTextSize As Size, sCaption  As String, lStyle As Long

    Call IUnknown_GetWindow(Form, VarPtr(hParent))
    hDc = GetDC(hParent)
    sCaption = Form.Caption
    Call GetTextExtentPoint32(hDc, sCaption, Len(sCaption), tTextSize)
    Call ReleaseDC(hParent, hDc)
    Form.Caption = Space(1 + tTextSize.cx / Len(sCaption)) & sCaption
    Call AtlAxWinInit
    lStyle = WS_POPUP Or WS_VISIBLE Or WS_DISABLED Or WS_EX_TOOLWINDOW
    hWebCtrl = CreateWindowEx(0, "AtlAxWin", "about:blank", lStyle, _
    0, 0, 0, 0, 0, 0, GetModuleHandle(vbNullString), ByVal 0)
    Call SetWindowLong(hWebCtrl, GWL_HWNDPARENT, hParent)
    If hWebCtrl Then
        Call SetProp(hParent, "hWebCtrl", hWebCtrl)
        Call AtlAxGetControl(hWebCtrl, Unk)
        Set oWbrowser = Unk
        With oWbrowser
            .Silent = True
            DoEvents
            .Document.Body.innerHTML = "<img style=""position:absolute;top:0px;left:0px;width:" & Fix(.Width) & _
            "px;height:" & Fix(.Height) & "px"" src=""" & url & "?" & ObjPtr(Form) & """/>"
            CreateIcon = CBool(Len(Dir(url))) And CBool(hWebCtrl)
        End With
    End If
   
End Function


Private Sub SetIconPos(ByVal Form As Object)

    Const SWP_SHOWWINDOW = &H40
    Const SWP_NOACTIVATE = &H10
    Const SPI_GETNONCLIENTMETRICS = 41
    Const GW_CHILD = 5

    #If Win64 Then
        Dim hParent As LongPtr, hwnd As LongPtr
    #Else
        Dim hParent As Long, hwnd As Long
    #End If

    Dim tParentRect As RECT, tChildRect As RECT
    Dim tNCM As NONCLIENTMETRICS

    Call IUnknown_GetWindow(Form, VarPtr(hParent))
    hwnd = GetProp(hParent, "hWebCtrl")
    Call GetWindowRect(GetNextWindow(hParent, GW_CHILD), tChildRect)
    Call GetWindowRect(hParent, tParentRect)
    tNCM.cbSize = Len(tNCM)
    If SystemParametersInfo(SPI_GETNONCLIENTMETRICS, 0, tNCM, 0) Then
        With tParentRect
            Call SetWindowPos(hwnd, 0, tChildRect.Left + 5, tChildRect.Top - tNCM.iCaptionHeight - tNCM.iPaddedBorderWidth, _
            tNCM.iCaptionHeight, tNCM.iCaptionHeight, SWP_NOACTIVATE + SWP_SHOWWINDOW)
        End With
    End If
End Sub

