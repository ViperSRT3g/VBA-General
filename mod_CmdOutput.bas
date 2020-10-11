Attribute VB_Name = "mod_CmdOutput"
Option Explicit

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Byte
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Const WAIT_INFINITE         As Long = (-1&)
Private Const STARTF_USESHOWWINDOW  As Long = &H1
Private Const STARTF_USESTDHANDLES  As Long = &H100
Private Const SW_HIDE               As Long = 0&

Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Sub GetStartupInfo Lib "kernel32" Alias "GetStartupInfoA" (lpStartupInfo As STARTUPINFO)
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long

Public Function Redirect(szBinaryPath As String, szCommandLn As String) As String
    Dim tSA_CreatePipe              As SECURITY_ATTRIBUTES
    Dim tSA_CreateProcessPrc        As SECURITY_ATTRIBUTES
    Dim tSA_CreateProcessThrd       As SECURITY_ATTRIBUTES
    Dim tSA_CreateProcessPrcInfo    As PROCESS_INFORMATION
    Dim tStartupInfo                As STARTUPINFO
    Dim hRead                       As Long
    Dim hWrite                      As Long
    Dim bRead                       As Long
    Dim abytBuff()                  As Byte
    Dim lngResult                   As Long
    Dim szFullCommand               As String
    Dim lngExitCode                 As Long
    Dim lngSizeOf                   As Long
    
    tSA_CreatePipe.nLength = Len(tSA_CreatePipe)
    tSA_CreatePipe.lpSecurityDescriptor = 0&
    tSA_CreatePipe.bInheritHandle = True
    
    tSA_CreateProcessPrc.nLength = Len(tSA_CreateProcessPrc)
    tSA_CreateProcessThrd.nLength = Len(tSA_CreateProcessThrd)
    
    If (CreatePipe(hRead, hWrite, tSA_CreatePipe, 0&) <> 0&) Then
        tStartupInfo.cb = Len(tStartupInfo)
        GetStartupInfo tStartupInfo
    
        With tStartupInfo
            .hStdOutput = hWrite
            .hStdError = hWrite
            .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
            .wShowWindow = SW_HIDE
        End With
        
        szFullCommand = """" & szBinaryPath & """" & " " & szCommandLn
        lngResult = CreateProcess(0&, szFullCommand, tSA_CreateProcessPrc, tSA_CreateProcessThrd, True, 0&, 0&, vbNullString, tStartupInfo, tSA_CreateProcessPrcInfo)
    
        If (lngResult <> 0&) Then
            lngResult = WaitForSingleObject(tSA_CreateProcessPrcInfo.hProcess, WAIT_INFINITE)
            lngSizeOf = GetFileSize(hRead, 0&)
            If (lngSizeOf > 0) Then
                ReDim abytBuff(lngSizeOf - 1)
                If ReadFile(hRead, abytBuff(0), UBound(abytBuff) + 1, bRead, ByVal 0&) Then
                    Redirect = StrConv(abytBuff, vbUnicode)
                End If
            End If
            Call GetExitCodeProcess(tSA_CreateProcessPrcInfo.hProcess, lngExitCode)
            CloseHandle tSA_CreateProcessPrcInfo.hThread
            CloseHandle tSA_CreateProcessPrcInfo.hProcess
                    
            If (lngExitCode <> 0&) Then Err.Raise vbObject + 1235&, "GetExitCodeProcess", "Non-zero Application exist code"
            
            CloseHandle hWrite
            CloseHandle hRead
        Else
            Err.Raise vbObject + 1236&, "CreateProcess", "CreateProcess Failed, Code: " & Err.LastDllError
        End If
    End If
End Function
