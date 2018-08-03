VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "主持人"
   ClientHeight    =   5370
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6090
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
   ScaleHeight     =   5370
   ScaleWidth      =   6090
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "准备阶段"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   5895
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "发布个股消息"
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1440
      End
   End
   Begin VB.Label Label4 
      Caption         =   "详情："
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "当前阶段："
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "当前回合："
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()
Dim 回合 As Long
回合 = Val(InputBox("请输入第几回合发布"))
个股消息发布(回合) = 个股消息性质选取
End Sub
