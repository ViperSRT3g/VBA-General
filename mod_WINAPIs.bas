Attribute VB_Name = "mod_WINAPIs"
Option Explicit

Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

#If Win64 And VBA7 Then
    Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
    Public Declare PtrSafe Function IntersectRect Lib "user32" (lpDestRect As Rect, lpSrc1Rect As Rect, lpSrc2Rect As Rect) As Long
    Public Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Public Declare PtrSafe Function SetWindowPos Lib "user32" _
        (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
        ByVal x As Long, ByVal y As Long, _
        ByVal cx As Long, ByVal cy As Long, _
        ByVal wFlags As Long) As Long
    Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
    Public Declare PtrSafe Function sndPlaySound32 Lib "winmm.dll" Alias "sndPlaySoundA" ( _
        ByVal lpszSoundName As String, _
        ByVal uFlags As Long) As Long
#Else
    Public Declare Function GetTickCount Lib "kernel32" () As Long
    Public Declare Function IntersectRect Lib "user32" (lpDestRect As Rect, lpSrc1Rect As Rect, lpSrc2Rect As Rect) As Long
    Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Public Declare Function SetWindowPos Lib "user32" _
        (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
        ByVal x As Long, ByVal y As Long, _
        ByVal cx As Long, ByVal cy As Long, _
        ByVal wFlags As Long) As Long
    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
    Public Declare Function sndPlaySound32 Lib "winmm.dll" Alias "sndPlaySoundA" ( _
        ByVal lpszSoundName As String, _
        ByVal uFlags As Long) As Long
#End If

Public Const SND_SYNC = &H0        ' (Default) Play the sound synchronously. Code execution pauses until sound is complete.
Public Const SND_ASYNC = &H1       ' Play the sound asynchronously. Code execution does not wait for sound to complete.
Public Const SND_NODEFAULT = &H2   ' If the specified sound is not found, do not play the default sound (no sound is played).
Public Const SND_MEMORY = &H4      ' lpszSoundName is a memory file of the sound. Not used in VBA/VB6.
Public Const SND_LOOP = &H8        ' Continue playing sound in a loop until the next call to sndPlaySound.
Public Const SND_NOSTOP = &H10     ' Do not stop playing the current sound before playing the specified sound.
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const SWP_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Function ExcelTopmost(ByVal Toggle As Boolean) As Long
    ExcelTopmost = SetWindowPos(Application.hwnd, IIf(Toggle, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_FLAGS)
End Function

Public Function Intersect(ByRef Rect1 As Rect, Rect2 As Rect) As Boolean
    Dim ResultRect As Rect: Intersect = IntersectRect(ResultRect, Rect1, Rect2)
End Function
