Attribute VB_Name = "modWinSock"
'***Module that holds all Winsock functionality**************************************'

Option Explicit

'Winsock messages that will go to the window handler
Public Const FD_READ = &H1&
Public Const FD_CONNECT = &H10&
Public Const FD_CLOSE = &H20&
Public Const FD_ACCEPT = &H8&

Public Const WSA_DESCRIPTIONLEN = 256
Public Const WSA_DescriptionSize = WSA_DESCRIPTIONLEN + 1
Private Const SOMAXCONN = &H7FFFFFFF

Public Const WSA_SYS_STATUS_LEN = 128
Public Const WSA_SysStatusSize = WSA_SYS_STATUS_LEN + 1
Private Const WSADESCRIPTION_LEN = 257
Private Const WSASYS_STATUS_LEN = 129

'Type of socket
Private Const SOCK_STREAM = 1
Private Const AF_INET = 2
Private Const IPPROTO_TCP = 6

Public Const HOSTENT_SIZE = 16
Public Const FD_SETSIZE = 64

'Winsock Data structure
Public Type WSAData
    wVersion       As Integer                       'Version
    wHighVersion   As Integer                       'High Version
    szDescription  As String * WSADESCRIPTION_LEN   'Description
    szSystemStatus As String * WSASYS_STATUS_LEN    'Status of system
    iMaxSockets    As Integer                       'Maximum number of sockets allowed
    iMaxUdpDg      As Integer                       'Maximum UDP datagrams
    lpVendorInfo   As Long                          'Vendor Info
End Type

'Socket Address structure
Public Type SOCKADDR_IN
    sin_family       As Integer
    sin_port         As Integer
    sin_addr         As Long
    sin_zero(1 To 8) As Byte
End Type

Type HostEnt
    h_name      As Long
    h_aliases   As Long
    h_addrtype  As Integer
    h_length    As Integer
    h_addr_list As Long
End Type

Public Type timeval
  tv_sec  As Long   'seconds
  tv_usec As Long   'microseconds
End Type

Public Type fd_set
  fd_count                  As Long
  fd_array(1 To FD_SETSIZE) As Long
End Type

'Types used for Ping
Private Type ICMP_OPTIONS
    Ttl             As Byte
    Tos             As Byte
    Flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type

Public Type ICMP_ECHO_REPLY
    Address         As Long
    status          As Long
    RoundTripTime   As Long
    DataSize        As Long
   'Reserved        As Integer
    DataPointer     As Long
    Options         As ICMP_OPTIONS
    Data            As String * 250
End Type

'Ping functions
Public Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Public Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal HANDLE As Long) As Boolean
Public Declare Function IcmpSendEcho Lib "ICMP" (ByVal IcmpHandle As Long, ByVal DestAddress As Long, ByVal RequestData As String, ByVal RequestSize As Integer, RequestOptns As ICMP_OPTIONS, ReplyBuffer As ICMP_ECHO_REPLY, ByVal ReplySize As Long, ByVal TimeOut As Long) As Boolean

'Winsock Initialization and termination
Public Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSAData) As Long
Public Declare Function WSACleanup Lib "ws2_32.dll" () As Long

'Socket Functions
Public Declare Function Socket Lib "ws2_32.dll" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
Public Declare Function WSACloseSocket Lib "ws2_32.dll" Alias "closesocket" (ByVal s As Long) As Long
Public Declare Function WSAAsyncSelect Lib "ws2_32.dll" (ByVal s As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long

'For checking the sockets state
Public Declare Function SocketState Lib "ws2_32.dll" Alias "select" (ByVal nfds As Long, ByRef readfds As Any, ByRef writefds As Any, ByRef exceptfds As Any, ByRef TimeOut As Any) As Long

'Reciving and sending data on winsock functions
Private Declare Function WSARecv Lib "ws2_32.dll" Alias "recv" (ByVal s As Long, ByRef Buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
Private Declare Function WSASend Lib "ws2_32.dll" Alias "send" (ByVal s As Long, ByRef Buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long

'Winsock API functions to create a listening server
Private Declare Function WSABind Lib "ws2_32.dll" Alias "bind" (ByVal s As Long, ByRef name As SOCKADDR_IN, ByRef namelen As Long) As Long
Private Declare Function WSAListen Lib "ws2_32.dll" Alias "listen" (ByVal s As Long, ByVal backlog As Long) As Long
Private Declare Function WSAAccept Lib "ws2_32.dll" Alias "accept" (ByVal s As Long, ByRef addr As SOCKADDR_IN, ByRef addrlen As Long) As Long

'Functions to convert numeric datatypes from and to VB so winsock and VB agree
Public Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer
Public Declare Function htonl Lib "ws2_32.dll" (ByVal hostlong As Long) As Long
Public Declare Function ntohl Lib "ws2_32.dll" (ByVal netlong As Long) As Long
Public Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Integer) As Integer

'Hostname resolving functions
Public Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long
Public Declare Function gethostbyaddr Lib "wsock32.dll" (haddr As Long, ByVal hnlen As Long, ByVal addrtype As Long) As Long
Public Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal inn As Long) As Long

'Memory functions
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Public Declare Function lstrcpyA Lib "kernel32" (ByVal RetVal As String, ByVal Ptr As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

'Window functions
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long


Public lngListenSocket As Long          'The listening socketdescriptor
Public blnListening As Boolean          'Flag to tell if listening or not
Public colClient As New Collection      'Collection of connected Clients
Public colNetIPLong As New Collection   'Collection of IP

'Public Const WINSOCK_MESSAGE = 1025     'Tells what kind of messages to listen for
Public WindowHandle As Long             'Handle to the window that recieves the winsock messages
Public MyIllegalSrvInstance As clsServer 'To have access to the created instance of clsserver

Private Const WM_USER = &H400
Public Const WINSOCK_MESSAGE  As Long = WM_USER + &H401 'Tells what kind of messages to listen for

'**Initiate WinSock to be used
Public Function StartWinsock() As Boolean
Dim StartUpData As WSAData, blnRet As Boolean

If Not WSAStartup(&H101, StartUpData) Then
    blnRet = True
    blnListening = True
Else
    blnRet = False
End If
StartWinsock = blnRet
End Function

'**Uninitiate WinSock and close all connection
Public Sub EndWinSock()
Dim i As Long, client As Variant
If blnListening Then
    On Error Resume Next
    For Each client In colClient
        CloseSocket CLng(client)
    Next
    
    Set colClient = Nothing
    Set colNetIPLong = Nothing
    
    WSACloseSocket lngListenSocket
    WSACleanup
    blnListening = False
End If
End Sub

'**Start listening for connections on specified Port
Public Function StartListen(ByVal lngPort As Long) As Boolean
Dim udtSocketAddress As SOCKADDR_IN, blnRet As Boolean

'Create a socket to listen on
lngListenSocket = Socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)

'Tell WinSock where to send the traffic and what to listen for
Call WSAAsyncSelect(lngListenSocket, WindowHandle, WINSOCK_MESSAGE, FD_CONNECT Or FD_READ Or FD_CLOSE Or FD_ACCEPT)

If lngListenSocket <> -1 Then
    With udtSocketAddress
        .sin_addr = inet_addr("0.0.0.0")                'Accept any IP
        .sin_port = htons(UnsignedToInteger(lngPort))   'The listening Port
        .sin_family = AF_INET                           'Listen to Internet family traffic (TCP etc)
    End With
    
    'Bind The listening port so it is dedicated
    Call WSABind(lngListenSocket, udtSocketAddress, LenB(udtSocketAddress))
    
    'Tell WinSock to start listen on our port and allow as many connection as supported
    Call WSAListen(lngListenSocket, SOMAXCONN)
    blnRet = True
Else
    blnRet = False
End If
StartListen = blnRet
End Function

'***Create a new socket that will have the same properties as the listening Socket
'***Like where the traffic is being listen for, IP´s allowed, and type of traffic
Public Sub AcceptConnection(ByVal SockDescriptor As Long)
Dim SckAddr As SOCKADDR_IN, tempSocket As Long, IP As String, Host As String
'Get a socketdescriptor for the new connection
tempSocket = WSAAccept(SockDescriptor, SckAddr, LenB(SckAddr))
'Tell WinSock where to send the traffic and what to listen for
'Call WSAAsyncSelect(tempSocket, WindowHandle, WINSOCK_MESSAGE, FD_CONNECT Or FD_READ Or FD_CLOSE Or FD_ACCEPT)

'If the socket is valid
If tempSocket <> -1 Then
    'Add it to the collections
    colClient.Add CStr(tempSocket), CStr(tempSocket)
    colNetIPLong.Add CStr(SckAddr.sin_addr), CStr(tempSocket)
    
    'Pass the new connection to the instance of clsServer that is already created
    MyIllegalSrvInstance.WinSockAccept tempSocket
End If

End Sub

'**Called from the messagehandler when the messagehandler was triggered by data arrival from a connection
Public Sub DataArrival(ByVal SockDescriptor As Long)
Dim sTemp As String, lRet As Long, szBuf As String

'Recive the data from the connection, Loop until all data is received
Do
    szBuf = String(256, 0)
    lRet = WSARecv(SockDescriptor, ByVal szBuf, Len(szBuf), 0)
    If lRet > 0 Then sTemp = sTemp + Left$(szBuf, lRet)
Loop Until lRet <= 0

'If the recieved data is bigger than 0 byte
If LenB(sTemp) > 0 Then
    'Pass the sending client and the data to the instance of clsServer that is already created
    MyIllegalSrvInstance.WinSockDataArrival SockDescriptor, sTemp
End If
End Sub

'**Close the socket with the specified socket descriptor
Public Sub CloseSocket(ByVal SockDescriptor As Long)
On Error GoTo Errhandler

'Remove it from the client collection
colClient.Remove CStr(SockDescriptor)
colNetIPLong.Remove CStr(SockDescriptor)

'Pass the closing client to the instance of clsServer that is already created
MyIllegalSrvInstance.WinSockCloseSocket SockDescriptor

'Close the socket
WSACloseSocket SockDescriptor
Exit Sub
Errhandler:
'Close the socket
WSACloseSocket SockDescriptor
End Sub

'**Send data to the client with the specified socket descriptor
Public Function SendData(ByVal sckDescriptor As Long, Message As String) As Long
Dim arrMessage() As Byte, strTemp As String

On Error GoTo Errhandler

    arrMessage = ""
    strTemp = StrConv(Message, vbFromUnicode)
    arrMessage = strTemp
    
    If UBound(arrMessage) > -1 Then
        SendData = WSASend(sckDescriptor, arrMessage(0), (UBound(arrMessage) - LBound(arrMessage) + 1), 0)
    End If

Errhandler:
End Function

'**Get the connections IPNumber
Public Function GetAscIP(ByVal Address As Long) As String
Dim strRet As String, lRet As Long
On Error GoTo Errhandler
lRet = inet_ntoa(Address)
strRet = String$(lstrlen(ByVal lRet), 0)
lstrcpyA ByVal strRet, ByVal lRet
GetAscIP = strRet
Errhandler:
End Function

'**Get the connections Hostname
Public Function GetHostByAddress(ByVal Address As Long) As String
Dim lLength As Long, lRet As Long

lRet = gethostbyaddr(Address, 4, AF_INET)
If lRet <> 0 Then
    CopyMemory lRet, ByVal lRet, 4
    lLength = lstrlen(lRet)
    If lLength > 0 Then
        GetHostByAddress = Space$(lLength)
        CopyMemory ByVal GetHostByAddress, ByVal lRet, lLength
    End If
Else
    GetHostByAddress = "UNKNOWN"
End If

Errhandler:
End Function

'Check the specified sockets state (Doesnt work!?)
Public Function IsSocketOK(ByVal sckDescriptor As Long) As Boolean
Dim udtRead_fd As fd_set, udtWrite_fd As fd_set, udtError_fd As fd_set
Dim lRet As Long
On Error GoTo Errhandler

udtWrite_fd.fd_count = 1
udtWrite_fd.fd_array(1) = sckDescriptor
lRet = SocketState(0&, udtRead_fd, udtWrite_fd, udtError_fd, 0&)

IsSocketOK = CBool(lRet)

Errhandler:
End Function

'Ping the client but beware that it isn´t async, short timeoutms is advisable
Public Function Ping(ByVal IPLong As Long, Optional TimeOutMS As Integer = 500) As Long
Dim hFile As Long, lpWSAdata As WSAData
Dim hHostent As HostEnt, AddrList As Long
Dim Address As Long, rIP As String
Dim OptInfo As ICMP_OPTIONS
Dim EchoReply As ICMP_ECHO_REPLY

On Error GoTo Errhandler

hFile = IcmpCreateFile()
If hFile <> 0 Then
    OptInfo.Ttl = 255
    If IcmpSendEcho(hFile, IPLong, String(32, "A"), 32, OptInfo, EchoReply, Len(EchoReply) + 8, TimeOutMS) Then
        If EchoReply.status = 0 Then Ping = EchoReply.RoundTripTime + 1
    End If
End If
IcmpCloseHandle hFile
Errhandler:
End Function

'Helper function converting long to integer
Public Function UnsignedToInteger(Value As Long) As Integer
If Value < 0 Or Value >= 65536 Then Err.Raise 6, Err.Source, Err.Description & vbCrLf & _
"Overflow, only numbers between 1 to 65536 is allowed as Port values!"

If Value <= 32767 Then
    UnsignedToInteger = Value
Else
    UnsignedToInteger = Value - 65536
End If

End Function

