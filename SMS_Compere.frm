VERSION 5.00
Begin VB.Form 主持人 
   Caption         =   "主持人"
   ClientHeight    =   5235
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12975
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   5235
   ScaleWidth      =   12975
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer2 
      Interval        =   10000
      Left            =   5040
      Top             =   4680
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   4440
      Left            =   6240
      TabIndex        =   5
      Top             =   600
      Width           =   6615
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5640
      Top             =   4680
   End
   Begin VB.Frame Frame1 
      Caption         =   "准备阶段"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   5895
      Begin VB.Label 发布大盘消息 
         AutoSize        =   -1  'True
         Caption         =   "发布大盘消息"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   7
         Top             =   360
         Width           =   1710
      End
      Begin VB.Label 发布个股消息 
         AutoSize        =   -1  'True
         Caption         =   "发布个股消息"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   1
         Top             =   360
         Width           =   1710
      End
   End
   Begin VB.Label Label4 
      Caption         =   "消息池："
      Height          =   375
      Left            =   6240
      TabIndex        =   6
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "详情："
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "当前阶段："
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "当前回合："
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "主持人"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
发布个股消息.FontBold = False
发布大盘消息.FontBold = False
End Sub
Private Sub 发布个股消息_Click()
Dim 回合 As Long
On Error GoTo Er:
回合 = Val(InputBox("请输入在多少回合发布一条行情消息(最少下1回合，最多未来5回合内)"))
If 回合 > 当前回合 + 5 Then 回合 = 当前回合 + 5
If 回合 <= 当前回合 Then 回合 = 当前回合 + 1
消息池(0, 回合).存在 = True
消息池(0, 回合).好坏度 = 个股消息性质选取
MsgBox "在第【" & 回合 & "】回合获得一个【" & 消息池(0, 回合).好坏度 & "】消息"
'If MsgBox("是否告知庄家？", vbYesNo) = 6 Then
主持人行情消息发送 "庄家", 回合 & "-" & 消息池(0, 回合).好坏度
'End If
Timer2_Timer
发布个股消息.Enabled = False
发布个股消息.FontBold = False
Exit Sub
Er:
MsgBox "操作失败！"
发布个股消息.FontBold = False
End Sub
Private Sub Timer1_Timer()
Label1.Caption = "当前回合：" & 当前回合
Label2.Caption = "当前阶段：" & 当前阶段
Label3.Caption = "详情：" & 操作详情
If 操作权限 = True Then 操作详情 = "主持人思考中..."
End Sub
Private Sub Timer2_Timer()
Dim i As Long
List1.Clear
For i = 当前回合 To 当前回合 + 5
    If 消息池(0, i).存在 = True Then
        List1.AddItem "在第【" & i & "】回合" & vbTab & "有一个【" & 消息池(0, i).好坏度 & "】消息"
    End If
Next
End Sub
Private Sub 发布大盘消息_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
发布大盘消息.FontBold = True
End Sub
Private Sub 发布个股消息_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
发布个股消息.FontBold = True
End Sub
