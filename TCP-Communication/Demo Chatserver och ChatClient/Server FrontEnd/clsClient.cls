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
Attribute VB_Exposed = False
Option Explicit

Private m_lngSocket         As Long
Private mvarNickName        As String

Public Property Get SocketHandle() As Long
    SocketHandle = m_lngSocket
End Property

Public Property Let SocketHandle(lngsocket As Long)
    m_lngSocket = lngsocket
End Property

Public Property Get NickName() As String
    NickName = mvarNickName
End Property

Public Property Let NickName(ByVal strNickname As String)
    mvarNickName = strNickname
End Property

