Attribute VB_Name = "modMessageHandler"
'***Module that create a message handler to recive all our WinSock messages**********'
'***                 and pass them on to be processed                      **********'
'************************************************************************************'
Option Explicit

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const GWL_WNDPROC = (-4)
Public Const WINSOCK_MESSAGE = 1026

Public WindowHandle As Long
Public prevProc As Long
Public MyIllegalInstance As clsClient

'**Create the messagehandler
Public Sub CreateMessageHandler()
'Create a invisible window
WindowHandle = CreateWindowEx(0&, "STATIC", "SOCKET_WINDOW", 0&, 0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, ByVal 0&)
'Tell the system to send our message to that window and get the handler to the original messagehandler
prevProc = SetWindowLong(WindowHandle, GWL_WNDPROC, AddressOf WindowProc)
End Sub

'**Destroy the windowhandler
Public Sub TerminateMessageHandler()
If prevProc <> 0 Then
    'Tell the system that the original messagehandler should be used
    SetWindowLong WindowHandle, GWL_WNDPROC, prevProc
    'Destroy the window that we created
    DestroyWindow WindowHandle
    prevProc = 0
End If
End Sub

'**The Function that will retrieve the messages
Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    'If it is our WinSock message
    If uMsg = WINSOCK_MESSAGE Then
        'Send the message to be processed
        ProcessMessage wParam, lParam
    Else
        'If it is another message pass it to the original messagehandler
        WindowProc = CallWindowProc(prevProc, hWnd, uMsg, wParam, lParam)
    End If
End Function

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


