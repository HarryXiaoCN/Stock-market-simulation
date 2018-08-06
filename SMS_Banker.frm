VERSION 5.00
Begin VB.Form 庄家 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "庄家"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12960
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
   MaxButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   12960
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "准备阶段"
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   5895
      Begin VB.Label 发布个股假消息 
         AutoSize        =   -1  'True
         Caption         =   "发布个股假消息"
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
         Left            =   1920
         TabIndex        =   5
         Top             =   480
         Width           =   1995
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   12240
      Top             =   120
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2550
      Left            =   6240
      TabIndex        =   3
      Top             =   600
      Width           =   6615
   End
   Begin VB.Timer Timer2 
      Interval        =   10000
      Left            =   11640
      Top             =   120
   End
   Begin VB.Frame Frame2 
      Caption         =   "行动阶段"
      Height          =   2735
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   5895
      Begin VB.TextBox 交易量 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1680
         TabIndex        =   17
         Top             =   2040
         Width           =   2415
      End
      Begin VB.TextBox 交易价 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1680
         TabIndex        =   10
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label 交易数量 
         Caption         =   "交易数量："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label 卖出 
         AutoSize        =   -1  'True
         Caption         =   "卖出"
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
         Left            =   4800
         TabIndex        =   13
         Top             =   2040
         Width           =   570
      End
      Begin VB.Label 买入 
         AutoSize        =   -1  'True
         Caption         =   "买入"
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
         Left            =   4800
         TabIndex        =   12
         Top             =   1440
         Width           =   570
      End
      Begin VB.Label 交易价格 
         Caption         =   "交易价格："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label 行动阶段倒计时 
         AutoSize        =   -1  'True
         Caption         =   "行动阶段倒计时：00：00：00"
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
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3840
      End
      Begin VB.Label 当前大盘状态 
         AutoSize        =   -1  'True
         Caption         =   "当前大盘状态："
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
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   1995
      End
   End
   Begin VB.Label 持仓显示 
      Caption         =   "持仓："
      Height          =   375
      Left            =   3240
      TabIndex        =   15
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label 资金显示 
      Caption         =   "现金："
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "当前回合："
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "当前阶段："
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "详情："
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "消息池："
      Height          =   375
      Left            =   6240
      TabIndex        =   6
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "庄家"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label5_Click()

End Sub

Private Sub Timer1_Timer()

End Sub
