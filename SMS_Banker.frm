VERSION 5.00
Begin VB.Form ׯ�� 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ׯ��"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12960
   BeginProperty Font 
      Name            =   "΢���ź�"
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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame1 
      Caption         =   "׼���׶�"
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   5895
      Begin VB.Label �������ɼ���Ϣ 
         AutoSize        =   -1  'True
         Caption         =   "�������ɼ���Ϣ"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
      Caption         =   "�ж��׶�"
      Height          =   2735
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   5895
      Begin VB.TextBox ������ 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "΢���ź�"
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
      Begin VB.TextBox ���׼� 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "΢���ź�"
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
      Begin VB.Label �������� 
         Caption         =   "����������"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
      Begin VB.Label ���� 
         AutoSize        =   -1  'True
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
      Begin VB.Label ���� 
         AutoSize        =   -1  'True
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
      Begin VB.Label ���׼۸� 
         Caption         =   "���׼۸�"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
      Begin VB.Label �ж��׶ε���ʱ 
         AutoSize        =   -1  'True
         Caption         =   "�ж��׶ε���ʱ��00��00��00"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
      Begin VB.Label ��ǰ����״̬ 
         AutoSize        =   -1  'True
         Caption         =   "��ǰ����״̬��"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
   Begin VB.Label �ֲ���ʾ 
      Caption         =   "�ֲ֣�"
      Height          =   375
      Left            =   3240
      TabIndex        =   15
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label �ʽ���ʾ 
      Caption         =   "�ֽ�"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "��ǰ�غϣ�"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "��ǰ�׶Σ�"
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "���飺"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "��Ϣ�أ�"
      Height          =   375
      Left            =   6240
      TabIndex        =   6
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "ׯ��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label5_Click()

End Sub

Private Sub Timer1_Timer()

End Sub
