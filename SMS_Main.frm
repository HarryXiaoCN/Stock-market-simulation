VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ӭ��"
   ClientHeight    =   1980
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   6645
   BeginProperty Font 
      Name            =   "΢���ź�"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   6645
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command2 
      Caption         =   "�����г�"
      Height          =   1695
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�����г�"
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Local_State = 0 Then Form3.Show: Unload Me Else MsgBox "���ȹرյ�ǰ���ӣ�"
End Sub

Private Sub Command2_Click()
If Local_State = 0 Then Form4.Show: Unload Me Else MsgBox "���ȹرյ�ǰ���ӣ�"
End Sub
