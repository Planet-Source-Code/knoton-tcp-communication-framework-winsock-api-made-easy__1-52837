VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Event OnDataArrival(ByVal sckDescriptor As Long, Data As String)
Public Event OnConnection(ByVal sckDescriptor As Long)
Public Event OnConnectionClose(ByVal sckDescriptor As Long)

Private sc As cSubclass
Implements WinSubHook.iSubclass

Private Sub Class_Initialize()
'Set this instance of clsServer so that other modules in this project can used the exposed functions and subs
Set MyIllegalInstance = Me
'Create the window that will deal with our WinSock messages
WindowHandle = CreateWindowEx(0&, "STATIC", "SOCKET_WINDOW", 0&, 0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, ByVal 0&)
Set sc = New cSubclass
Call sc.Subclass(WindowHandle, Me)
Call sc.AddMsg(WINSOCK_MESSAGE, MSG_AFTER)

StartWinsock

End Sub

Private Sub Class_Terminate()
'Destroy the window that deal with our winsock messages
DestroyWindow WindowHandle
EndWinSock
End Sub

Public Function Connect(ByVal stRemoteHost As String, ByVal RemotePort As Long) As Long
Connect = ConnectTo(stRemoteHost, RemotePort)
End Function

Public Function Send(ByVal sckDescriptor As Long, ByVal strData As String) As Long
    Send = SendData(sckDescriptor, strData)
End Function

Public Sub CloseSocket(ByVal sckDescriptor As Long)
WSACloseSocket sckDescriptor
WinSockCloseSocket sckDescriptor
End Sub

Friend Sub WinSockDataArrival(ByVal sckDescriptor As Long, ByVal Data As String)
RaiseEvent OnDataArrival(sckDescriptor, Data)
End Sub

Friend Sub WinSockCloseSocket(ByVal sckDescriptor As Long)
RaiseEvent OnConnectionClose(sckDescriptor)
colClient.Remove CStr(sckDescriptor)
End Sub

Friend Sub WinSockConnection(ByVal sckDescriptor As Long)
colClient.Add CStr(sckDescriptor), CStr(sckDescriptor)
RaiseEvent OnConnection(sckDescriptor)
End Sub

Private Sub iSubclass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As WinSubHook.eMsg, ByVal wParam As Long, ByVal lParam As Long)
    'If it is our WinSock message
    If uMsg = WINSOCK_MESSAGE Then
        'Send the message to be processed
        ProcessMessage wParam, lParam
    End If
End Sub

Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, hWnd As Long, uMsg As WinSubHook.eMsg, wParam As Long, lParam As Long)
'
End Sub
'**The function that will Process the Winsock messages
Private Function ProcessMessage(ByVal wParam As Long, ByVal lParam As Long)
    Select Case lParam
        Case FD_CONNECT
            ConnectionEstablished wParam
        Case FD_READ
            DataArrival wParam
        Case FD_CLOSE
            CloseSocket wParam
    End Select
End Function

