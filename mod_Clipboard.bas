Attribute VB_Name = "mod_Clipboard"
Option Explicit

Private Declare Function OpenClipboard Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32.dll" () As Long
Private Declare Function CloseClipboard Lib "user32.dll" () As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare Function GetClipboardData Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long

Public Const GMEM_MOVEABLE As Long = &H2
Public Const GMEM_ZEROINIT As Long = &H40
Public Const CF_UNICODETEXT As Long = &HD

Public Function SetClipboard(ByVal RawString As String)
    Dim StringPointer As Long, BufferLen As Long, LockHandle As Long
    Call OpenClipboard(0&)
    Call EmptyClipboard
    BufferLen = LenB(RawString) + 2&
    StringPointer = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, BufferLen)
    LockHandle = GlobalLock(StringPointer)
    Call lstrcpy(LockHandle, StrPtr(RawString))
    Call GlobalUnlock(StringPointer)
    Call SetClipboardData(CF_UNICODETEXT, StringPointer)
    Call CloseClipboard
    SetClipboard = (GetClipboard = RawString)
End Function

Public Function GetClipboard() As String
    Dim StringPointer As Long, BufferLen As Long, LockHandle As Long
    Dim ClipData As String
    Const CF_UNICODETEXT As Long = 13&
    Call OpenClipboard(0&)
    If IsClipboardFormatAvailable(CF_UNICODETEXT) Then
        StringPointer = GetClipboardData(CF_UNICODETEXT)
        If StringPointer Then
            LockHandle = GlobalLock(StringPointer)
            BufferLen = GlobalSize(StringPointer)
            ClipData = String$(BufferLen \ 2& - 1&, vbNullChar)
            Call lstrcpy(StrPtr(ClipData), LockHandle)
            Call GlobalUnlock(StringPointer)
        End If
        GetClipboard = ClipData
    End If
    Call CloseClipboard
End Function
