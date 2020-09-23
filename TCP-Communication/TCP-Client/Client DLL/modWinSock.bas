Attribute VB_Name = "modWinSock"
'***Module that holds all Winsock functionality**************************************'

Option Explicit

Public Const INADDR_NONE = &HFFFF
Public Const SOCKET_ERROR = -1
Public Const INVALID_SOCKET = -1

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


'Winsock Initialization and termination
Public Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSAData) As Long
Public Declare Function WSACleanup Lib "ws2_32.dll" () As Long

'Socket Functions
Public Declare Function Socket Lib "ws2_32.dll" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
Public Declare Function WSACloseSocket Lib "ws2_32.dll" Alias "closesocket" (ByVal s As Long) As Long
Public Declare Function WSAAsyncSelect Lib "ws2_32.dll" (ByVal s As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Private Declare Function WSAConnect Lib "ws2_32.dll" Alias "connect" (ByVal s As Long, ByRef name As SOCKADDR_IN, ByVal namelen As Long) As Long

'For checking the sockets state
Public Declare Function SocketState Lib "ws2_32.dll" Alias "select" (ByVal nfds As Long, ByRef readfds As Any, ByRef writefds As Any, ByRef exceptfds As Any, ByRef timeout As Long) As Long

'Reciving and sending data on winsock functions
Private Declare Function WSARecv Lib "ws2_32.dll" Alias "recv" (ByVal s As Long, ByRef Buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
Private Declare Function WSASend Lib "ws2_32.dll" Alias "send" (ByVal s As Long, ByRef Buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long

'Functions to convert numeric datatypes from and to VB so winsock and VB agree
Public Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer

'Hostname resolving functions
Public Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long
Public Declare Function gethostbyname Lib "ws2_32.dll" (ByVal host_name As String) As Long

'Memory functions
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

Public MyIllegalInstance As clsClient

Private Const WM_USER = &H400
Public Const WINSOCK_MESSAGE  As Long = WM_USER + &H402 'Tells what kind of messages to listen for

'Public Const WINSOCK_MESSAGE = 1026
Public WindowHandle As Long


Public blnListening As Boolean
Public colClient As New Collection

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
Dim i As Long, obj As Variant
If blnListening Then
    For Each obj In colClient
        CloseSocket obj
    Next
    WSACleanup
    blnListening = False
End If
End Sub

Public Function ConnectTo(ByVal strRemoteHost As String, ByVal RemotePort As Long) As Long
Dim udtSocketAddress As SOCKADDR_IN, blnRet As Boolean, lngAdress As Long, sckDescriptor As Long
Dim conRet As Long

'Create a socket to listen on
sckDescriptor = Socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
If sckDescriptor > 0 Then
    'Tell WinSock where to send the traffic and what to listen for
    Call WSAAsyncSelect(sckDescriptor, WindowHandle, WINSOCK_MESSAGE, FD_CONNECT Or FD_READ Or FD_CLOSE)
    lngAdress = GetAddressLong(strRemoteHost)
    If lngAdress <> INADDR_NONE Then
        With udtSocketAddress
            .sin_addr = lngAdress
            .sin_port = htons(UnsignedToInteger(RemotePort))
            .sin_family = AF_INET
        End With
        Call WSAConnect(sckDescriptor, udtSocketAddress, LenB(udtSocketAddress))
        ConnectTo = sckDescriptor
    End If
End If
End Function

'**Called from the messagehandler when the messagehandler was triggered by data arrival from a connection
Public Sub DataArrival(ByVal SockDescriptor As Long)
Dim sTemp As String, lRet As Long, szBuf As String
On Error GoTo errHandler
'Recive the data from the connection, Loop until all data is received
Do
    szBuf = String(256, 0)
    lRet = WSARecv(SockDescriptor, ByVal szBuf, Len(szBuf), 0)
    If lRet > 0 Then sTemp = sTemp + Left$(szBuf, lRet)
Loop Until lRet <= 0

'If the recieved data is bigger than 0 byte
If LenB(sTemp) > 0 Then
    MyIllegalInstance.WinSockDataArrival SockDescriptor, sTemp
End If
errHandler:
End Sub

Public Sub ConnectionEstablished(ByVal sckDescriptor As Long)
MyIllegalInstance.WinSockConnection sckDescriptor
End Sub

'**Close the socket with the specified socket descriptor
Public Sub CloseSocket(ByVal SockDescriptor As Long)
On Error Resume Next
WSACloseSocket SockDescriptor
'Pass the closing client to the instance of clsServer that is already created
MyIllegalInstance.WinSockCloseSocket SockDescriptor
End Sub

'**Send data to the client with the specified socket descriptor
Public Function SendData(ByVal sckDescriptor As Long, ByVal Message As String) As Long
Dim arrMessage() As Byte, strTemp As String
On Error GoTo errHandler

If IsSocketOK(sckDescriptor) Then
    arrMessage = ""
    strTemp = StrConv(Message, vbFromUnicode)
    arrMessage = strTemp
    
    If UBound(arrMessage) > -1 Then
        SendData = WSASend(sckDescriptor, arrMessage(0), (UBound(arrMessage) - LBound(arrMessage) + 1), 0)
    End If
End If

errHandler:
End Function


'Check the specified sockets state
Public Function IsSocketOK(ByVal sckDescriptor As Long) As Boolean
Dim udtRead_fd As fd_set, udtWrite_fd As fd_set, udtError_fd As fd_set
Dim lRet As Long
On Error GoTo errHandler

udtWrite_fd.fd_count = 1
udtWrite_fd.fd_array(1) = sckDescriptor

lRet = SocketState(0&, udtRead_fd, udtWrite_fd, udtError_fd, 0&)

IsSocketOK = CBool(lRet)

errHandler:
End Function

'Get the remoteadress as network long
Public Function GetAddressLong(ByVal strHostName As String) As Long
Dim lngPtrToHOSTENT As Long
Dim udtHostent      As HostEnt
Dim lngPtrToIP      As Long
Dim lngAddress As Long

On Error GoTo errHandler

lngAddress = inet_addr(strHostName)
If lngAddress = INADDR_NONE Then
    lngPtrToHOSTENT = gethostbyname(strHostName)
    If lngPtrToHOSTENT <> 0 Then
        CopyMemory udtHostent, ByVal lngPtrToHOSTENT, LenB(udtHostent)
        CopyMemory lngPtrToIP, ByVal udtHostent.h_addr_list, 4
        CopyMemory lngAddress, ByVal lngPtrToIP, udtHostent.h_length
    Else
        lngAddress = INADDR_NONE
    End If
End If

GetAddressLong = lngAddress
Exit Function
errHandler:
GetAddressLong = INADDR_NONE
End Function

Public Function UnsignedToInteger(Value As Long) As Integer
If Value < 0 Or Value >= 65536 Then Err.Raise 6, Err.Source, Err.Description & vbCrLf & _
"Overflow, only numbers between 1 to 65536 is allowed as Port values!"

If Value <= 32767 Then
    UnsignedToInteger = Value
Else
    UnsignedToInteger = Value - 65536
End If

End Function


