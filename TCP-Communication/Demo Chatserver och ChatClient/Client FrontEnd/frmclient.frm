VERSION 5.00
Begin VB.Form frmClient 
   Caption         =   "Client"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSendPM 
      Caption         =   "PM"
      Height          =   495
      Left            =   7140
      TabIndex        =   14
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ListBox lstUsers 
      Height          =   1620
      Left            =   6840
      Sorted          =   -1  'True
      TabIndex        =   13
      ToolTipText     =   "Click User to enable PM"
      Top             =   300
      Width           =   1815
   End
   Begin VB.TextBox txtNickname 
      Height          =   285
      Left            =   5460
      TabIndex        =   10
      Top             =   300
      Width           =   1335
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   3240
      TabIndex        =   7
      Top             =   300
      Width           =   795
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   300
      Width           =   1875
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   495
      Left            =   7140
      TabIndex        =   5
      Top             =   3060
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   3240
      Width           =   6735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear Chat"
      Height          =   495
      Left            =   7140
      TabIndex        =   3
      Top             =   1980
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtChat 
      Height          =   2415
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   720
      Width           =   6735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Users"
      Height          =   195
      Left            =   6900
      TabIndex        =   12
      Top             =   60
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Nickname"
      Height          =   195
      Left            =   5460
      TabIndex        =   11
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Port"
      Height          =   195
      Left            =   3240
      TabIndex        =   9
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Server"
      Height          =   195
      Left            =   1320
      TabIndex        =   8
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents myClient As TCPClient.clsClient
Attribute myClient.VB_VarHelpID = -1
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Col As New Collection
Private ant As Long
Private mySocket As Long
Private myNick As String
Private PMNick As String

Private Sub cmdSendPM_Click()
Dim Data As String
Data = txtSend.Text
'Remove Pipe, will interfere with the protocol
Data = Replace(Data, "|", "")
txtSend = ""

myClient.Send mySocket, "$SENDPM|" & PMNick & "|" & Data
WriteData "<PM FROM: " & myNick & ">" & Data
End Sub

Private Sub Command1_Click()
Dim Server As String, Port As Integer
Server = txtServer.Text
Port = txtPort.Text
myNick = txtNickname.Text
myClient.Connect Server, Port
Command1.Enabled = False
Command2.Enabled = True
End Sub

Private Sub Command2_Click()
If mySocket <> 0 Then myClient.CloseSocket mySocket
Command1.Enabled = True
Command2.Enabled = False
End Sub

Private Sub Command3_Click()
txtChat.Text = ""
End Sub

Private Sub Command4_Click()
Dim Data As String
Data = txtSend.Text
'Remove Pipe, will interfere with the protocol
Data = Replace(Data, "|", "")
txtSend.Text = ""
myClient.Send mySocket, Data
End Sub

Private Sub Form_Load()
Set myClient = New TCPClient.clsClient
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set myClient = Nothing
End Sub

Private Sub lstUsers_Click()
PMNick = lstUsers.List(lstUsers.ListIndex)
cmdSendPM.Caption = "PM " & PMNick
End Sub

Private Sub myClient_OnConnection(ByVal sckDescriptor As Long)
mySocket = sckDescriptor
myClient.Send sckDescriptor, "$NICK" & myNick
End Sub

Private Sub myClient_OnConnectionClose(ByVal sckDescriptor As Long)
lstUsers.Clear
Command1.Enabled = True
Command2.Enabled = False
mySocket = 0
End Sub

Private Sub myClient_OnDataArrival(ByVal sckDescriptor As Long, Data As String)
Dim splitData As Variant, i As Integer, blnWriteChat As Boolean, Temp As String
On Error GoTo errHandler

If Mid(Data, 1, 9) = "$NICKLIST" Then
    lstUsers.Clear
    splitData = Split(Data, "|")
    For i = 1 To UBound(splitData)
        lstUsers.AddItem splitData(i)
    Next
Else
    'blnWriteChat = True
    WriteData Data
End If

'If blnWriteChat Then
'    WriteData Data
''    If LenB(txtChat.Text) > 32500 Then
''        txtChat.Text = Right(txtChat.Text, 32500 - LenB(Data)) & Data & vbCrLf
''    Else
''        txtChat.Text = txtChat.Text & Data & vbCrLf
''    End If
''
''    txtChat.SelStart = Len(txtChat.Text)
'End If
errHandler:
End Sub

Public Sub WriteData(Data As String)
If LenB(txtChat.Text) > 32500 Then
    txtChat.Text = Right(txtChat.Text, 32500 - LenB(Data)) & Data & vbCrLf
Else
    txtChat.Text = txtChat.Text & Data & vbCrLf
End If

txtChat.SelStart = Len(txtChat.Text)
End Sub

Private Sub txtNickname_Change()
txtNickname.Text = Replace(txtNickname.Text, "|", "")
txtNickname.SelStart = Len(txtNickname)
End Sub
