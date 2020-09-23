VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCP-Server  Connections: 0"
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Text            =   "412"
      Top             =   300
      Width           =   675
   End
   Begin VB.CommandButton Command4 
      Caption         =   "SendDataToAll"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3540
      TabIndex        =   4
      Top             =   660
      Width           =   1395
   End
   Begin VB.TextBox txtSendToAll 
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   900
      Width           =   3435
   End
   Begin VB.CommandButton Command3 
      Caption         =   "No of connection"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3540
      TabIndex        =   2
      Top             =   120
      Width           =   1395
   End
   Begin VB.CommandButton Command2 
      Caption         =   "StopListen"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "StartListen"
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Data to be sent to all connected clients"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   660
      Width           =   3435
   End
   Begin VB.Label Label1 
      Caption         =   "Port"
      Height          =   195
      Left            =   1320
      TabIndex        =   6
      Top             =   120
      Width           =   675
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents myHub As TCPServer.clsServer
Attribute myHub.VB_VarHelpID = -1
Private ColClient As New Collection
Private NickList As String

Private Sub Command1_Click()
Dim Port As Long
If Not IsNumeric(txtPort.Text) Then
    MsgBox "Port must be numeric !!"
    Exit Sub
End If
Port = txtPort.Text
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command1.Enabled = False
myHub.StartServing Port
End Sub

Private Sub Command2_Click()
myHub.StopServing
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command1.Enabled = True
End Sub

Private Sub Command3_Click()
Form1.Caption = " TCP-Server  Connections: " & myHub.NoOfConnections
End Sub

Private Sub Command4_Click()
Dim Data As String
Data = txtSendToAll.Text
If Data <> "" Then
    myHub.SendDataToAll Data
End If
End Sub

Private Sub Form_Load()
Set myHub = New TCPServer.clsServer
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
myHub.StopServing
Set myHub = Nothing
End Sub

Private Sub myHub_OnDataArrival(ByVal sckDescriptor As Long, ByVal Data As String)
Dim client As clsClient, tempSocket As Long, splitData As Variant, NickName As String
On Error GoTo errHandler

Set client = ColClient.Item(CStr(sckDescriptor))
NickName = client.NickName

Select Case NickName
    Case ""
        If Mid(Data, 1, 5) = "$NICK" Then
            Dim Nick As String
            Nick = Mid(Data, 6)
            If IsNickFree(Nick) Then
                client.NickName = Nick
                NickList = NickList & "|" & Nick
                myHub.SendDataToAll "$NICKLIST" & NickList
            Else
                myHub.SendDataToClient sckDescriptor, "Your nick is already taken, please change it !"
                myHub.CloseSocketHandle sckDescriptor
                Exit Sub
            End If
        Else
            myHub.SendDataToClient sckDescriptor, "You must send your nickname first !"
        End If
    Case Else
        If Mid(Data, 1, 7) = "$SENDPM" Then
            splitData = Split(Data, "|")
            tempSocket = GetSocketHandleByNick(splitData(1))
            If tempSocket <> 0 Then
                myHub.SendDataToClient tempSocket, "<PM FROM: " & NickName & "> " & splitData(2)
            End If
        Else
            myHub.SendDataToAll "<" & NickName & "> " & Data
        End If
End Select

errHandler:
End Sub

Private Sub myHub_OnNewConnection(ByVal sckDescriptor As Long)
Dim client As New clsClient
client.SocketHandle = sckDescriptor
ColClient.Add client, CStr(sckDescriptor)


Form1.Caption = " TCP-Server  Connections: " & myHub.NoOfConnections

End Sub


Private Sub myHub_OnSocketClose(ByVal sckDescriptor As Long)
Dim client As clsClient, Nick As String
Set client = ColClient(CStr(sckDescriptor))
Nick = client.NickName
Set client = Nothing
ColClient.Remove CStr(sckDescriptor)

If Nick <> "" Then
    NickList = Replace(NickList, "|" & Nick, "")
    myHub.SendDataToAll "$NICKLIST" & NickList
End If

Form1.Caption = " TCP-Server  Connections: " & myHub.NoOfConnections
End Sub

Private Function IsNickFree(Nick As String) As Boolean
Dim client As clsClient
IsNickFree = True
For Each client In ColClient
If client.NickName = Nick Then
    IsNickFree = False
    Exit For
End If
Next

Exit Function
errHandler:
IsNickFree = False
End Function

Private Function GetSocketHandleByNick(ByVal Nick As String) As Long
Dim client As clsClient
For Each client In ColClient
    If client.NickName = Nick Then
        GetSocketHandleByNick = client.SocketHandle
        Exit For
    End If
Next
End Function

Private Function GetNickList() As String
Dim strRet As String, client As clsClient
For Each client In ColClient
    If client.NickName <> "" Then
        If strRet <> "" Then
            strRet = client.NickName & "|"
        Else
            strRet = strRet & client.NickName & "|"
        End If
    End If
Next
If Right(strRet, 1) = "|" Then strRet = Mid(strRet, 1, Len(strRet) - 1)
GetNickList = strRet
End Function
