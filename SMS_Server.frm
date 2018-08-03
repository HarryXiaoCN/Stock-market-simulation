VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form 服务器 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "主持人端"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6495
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   6495
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "开始游戏"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   6255
   End
   Begin VB.Timer OlTime 
      Interval        =   100
      Left            =   5160
      Top             =   0
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   5640
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   8888
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   3180
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   6255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "玩家列表："
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1200
   End
End
Attribute VB_Name = "服务器"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
游戏初始化
主持人.Show
End Sub
Private Sub Form_Load()
Local_State = 1
Winsock1(0).Listen
End Sub
Private Sub Form_Unload(Cancel As Integer)
Local_State = 0: TCPNum = 0: 欢迎.Show
End Sub
Private Sub OlTime_Timer()
For i = 1 To TCPNum
    List1.List(i - 1) = "客户机IP：" & Winsock1(i).RemoteHostIP & vbTab & 连接状态反馈(Winsock1(i).State)
Next
End Sub
Private Sub Winsock1_Close(Index As Integer)
Winsock1(Index).Close
End Sub
Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal RequestID As Long)
If Index = 0 Then
    For i = 1 To TCPNum
        If Winsock1(i).State = 0 Then
            Winsock1(i).Accept RequestID
            Exit Sub
        End If
    Next
    TCPNum = TCPNum + 1
    Load Winsock1(TCPNum)
    Winsock1(TCPNum).LocalPort = 8888
    Winsock1(TCPNum).Accept RequestID
End If
End Sub
Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim TCPData As String
Winsock1(Index).GetData TCPData
Server_GetData TCPData
End Sub
