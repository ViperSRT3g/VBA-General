Attribute VB_Name = "mod_PingAddress"
Option Explicit

Private Declare Function GetHostByName Lib "wsock32.dll" Alias "gethostbyname" (ByVal strDomainName As String) As Long
Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired&, lpWSAdata As WSAdata) As Long
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal HANDLE As Long) As Boolean
Private Declare Function IcmpSendEcho Lib "ICMP" ( _
    ByVal IcmpHandle As Long, _
    ByVal DestAddress As Long, _
    ByVal RequestData As String, _
    ByVal RequestSize As Integer, _
    RequestOptns As IP_OPTION_INFORMATION, _
    ReplyBuffer As IP_ECHO_REPLY, _
    ByVal ReplySize As Long, _
    ByVal TimeOut As Long) As Boolean
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Private Type WSAdata 'wsock32.dll data
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To 255) As Byte
    szSystemStatus(0 To 128) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Private Type Hostent 'Memory copy
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type

Private Type IP_OPTION_INFORMATION 'Optional ICMP info
    TTL As Byte
    Tos As Byte
    Flags As Byte
    OptionsSize As Long
    OptionsData As String * 128
End Type

Private Type IP_ECHO_REPLY 'ICMP return data
    Address(0 To 3) As Byte
    Status As Long
    RoundTripTime As Long
    DataSize As Integer
    Reserved As Integer
    data As Long
    Options As IP_OPTION_INFORMATION
End Type
 
Public Function PingAddress(strDomainName) As String
    Dim hFile As Long, AddrList As Long, Address As Long
    Dim rIP As String
    Dim lpWSAdata As WSAdata: Call WSAStartup(&H101, lpWSAdata) 'Start wsock32.dll
    Dim hHostent As Hostent
    Dim OptInfo As IP_OPTION_INFORMATION
    Dim EchoReply As IP_ECHO_REPLY
    
    If GetHostByName(strDomainName) <> 0 Then 'Get Hostname from domain
        Call CopyMemory(hHostent.h_name, ByVal GetHostByName(strDomainName), Len(hHostent))
        Call CopyMemory(AddrList, ByVal hHostent.h_addr_list, 4)
        Call CopyMemory(Address, ByVal AddrList, 4)
    Else
        PingAddress = "Error, Unable to get host name"
        Exit Function
    End If
    
    'Get file handle for Internet Control Message Protocol Data
    hFile = IcmpCreateFile()
    If hFile = 0 Then
        PingAddress = "Error, Unable to Create File Handle"
        Exit Function
    End If
    
    'TTL is the Time To Live, i.e. number of hops the request will make before failure
    OptInfo.TTL = 255
    
    'send the request - return data to structures declared above
    If IcmpSendEcho(hFile, Address, String(32, "A"), 32, OptInfo, EchoReply, Len(EchoReply) & 8, 2000) Then
        rIP = CStr(EchoReply.Address(0)) & "." & _
              CStr(EchoReply.Address(1)) & "." & _
              CStr(EchoReply.Address(2)) & "." & _
              CStr(EchoReply.Address(3))
    Else: PingAddress = "Error, Timeout"
    End If
    If EchoReply.Status = 0 Then: PingAddress = rIP & "," & Trim$(CStr(EchoReply.RoundTripTime)) & "ms"
    Else: PingAddress = "Error, Failure"
    End If
    
    'Close data file handle, and wsock32.dll
    Call IcmpCloseHandle(hFile)
    Call WSACleanup
End Function
