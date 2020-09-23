VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*******Class that expose the events to the creator of the class*********************'
'******               And expose helper functions               *********************'
'************************************************************************************'

Option Explicit
'Event to be triggered when data arrives
Public Event OnDataArrival(ByVal sckDescriptor As Long, ByVal Data As String)
Public Event OnNewConnection(ByVal sckDescriptor As Long)
Public Event OnSocketClose(ByVal sckDescriptor As Long)
Private sc As cSubclass
Implements WinSubHook.iSubclass


'**Exposed function to start serving the server
Public Function StartServing(ByVal lngPort As Long) As Boolean
StartWinsock
StartServing = StartListen(lngPort)
End Function

'**Exposed function to stop serving the server
Public Sub StopServing()
EndWinSock
End Sub

Private Sub Class_Initialize()
'Set this instance of clsServer so that other modules in this project can used the exposed functions and subs
Set MyIllegalSrvInstance = Me
'Create a window to get a handle that will recieve WinSock messages
WindowHandle = CreateWindowEx(0&, "STATIC", "SOCKET_WINDOW", 0&, 0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, ByVal 0&)

Set sc = New cSubclass
Call sc.Subclass(WindowHandle, Me)
Call sc.AddMsg(WINSOCK_MESSAGE, MSG_AFTER)
End Sub

Private Sub Class_Terminate()
'Destroy the window that recieves the winsock messages
DestroyWindow WindowHandle
End Sub

'**Exposed function to get number of connections
Public Function NoOfConnections() As Long
NoOfConnections = colClient.Count
End Function

'**Exposed function to send data to a specific connection
Public Function SendDataToClient(ByVal sckDesc As Long, ByVal Message As String) As Long
Dim ret As Long
If LenB(Message) > 0 Then
    ret = SendData(sckDesc, Message)
    
    If ret = -1 Then
        CloseSocket sckDesc
    End If
End If
SendDataToClient = ret
End Function

'**Exposed function to send data to all connected connections
Public Sub SendDataToAll(ByVal Message As String)
Dim Obj As Variant, ArrCloseClient() As Long, i As Integer

On Error Resume Next
ReDim ArrCloseClient(0)
If LenB(Message) > 0 Then
    For Each Obj In colClient
        If SendData(CLng(Obj), Message) = -1 Then
            If ArrCloseClient(0) <> 0 Then
                ReDim Preserve ArrCloseClient(UBound(ArrCloseClient) + 1)
                ArrCloseClient(UBound(ArrCloseClient)) = CLng(Obj)
            Else
                ArrCloseClient(0) = CLng(Obj)
            End If
        End If
    Next
    
    If ArrCloseClient(0) <> 0 Then
        For i = 0 To UBound(ArrCloseClient)
            CloseSocket ArrCloseClient(i)
        Next
    End If
End If
End Sub

'**Exposed function to close the socket with the specified socket descriptor
Public Sub CloseSocketHandle(ByVal sckDesc As Long)
CloseSocket sckDesc
End Sub

'**Get clients IP
Public Function GetUserIP(ByVal sckDesc As Long) As String
GetUserIP = GetAscIP(colNetIPLong(CStr(sckDesc)))
End Function

'**Get Clients Hostname
Public Function GetUserHostName(ByVal sckDesc As Long) As String
GetUserHostName = GetHostByAddress(CLng(colNetIPLong(CStr(sckDesc))))
End Function

'**Get Clients Ping
Public Function GetUserPing(ByVal sckDesc As Long, Optional TimeOutMS As Integer = 500) As Integer
Dim IPLong As Long
IPLong = CLng(colNetIPLong(CStr(sckDesc)))
GetUserPing = Ping(IPLong, TimeOutMS)
End Function

'**Called from AcceptConnection in modWinsock
Friend Sub WinSockAccept(ByVal sckDescriptor As Long)
'Raise the event to the creator of this class
RaiseEvent OnNewConnection(sckDescriptor)
End Sub

Friend Sub WinSockDataArrival(ByVal sckDescriptor As Long, ByVal Data)
'Raise the event to the creator of this class
RaiseEvent OnDataArrival(sckDescriptor, Data)
End Sub

Friend Sub WinSockCloseSocket(ByVal sckDescriptor As Long)
'Raise the event to the creator of this class
RaiseEvent OnSocketClose(sckDescriptor)
End Sub

'**Triggered by winsubhook when messages arrives
Private Sub iSubclass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As WinSubHook.eMsg, ByVal wParam As Long, ByVal lParam As Long)
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
    Case FD_ACCEPT
        AcceptConnection wParam
    Case FD_READ
        DataArrival wParam
    Case Else
        CloseSocket wParam
End Select
End Function

