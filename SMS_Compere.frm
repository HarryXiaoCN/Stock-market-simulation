VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "������"
   ClientHeight    =   5370
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6090
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
   ScaleHeight     =   5370
   ScaleWidth      =   6090
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame1 
      Caption         =   "׼���׶�"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   5895
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "����������Ϣ"
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1440
      End
   End
   Begin VB.Label Label4 
      Caption         =   "���飺"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "��ǰ�׶Σ�"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "��ǰ�غϣ�"
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
Dim �غ� As Long
�غ� = Val(InputBox("������ڼ��غϷ���"))
������Ϣ����(�غ�) = ������Ϣ����ѡȡ
End Sub
