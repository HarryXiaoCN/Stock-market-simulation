VERSION 5.00
Begin VB.Form ������ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12975
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
   ScaleHeight     =   5145
   ScaleWidth      =   12975
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame2 
      Caption         =   "�ж��׶�"
      Height          =   1935
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   5895
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
         TabIndex        =   10
         Top             =   1080
         Width           =   1995
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
         TabIndex        =   9
         Top             =   480
         Width           =   3840
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   10000
      Left            =   11640
      Top             =   120
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
      Left            =   12240
      Top             =   120
   End
   Begin VB.Frame Frame1 
      Caption         =   "׼���׶�"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   5895
      Begin VB.Label ����������Ϣ 
         AutoSize        =   -1  'True
         Caption         =   "����������Ϣ"
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
         Left            =   3480
         TabIndex        =   7
         Top             =   840
         Width           =   1710
      End
      Begin VB.Label ����������Ϣ 
         AutoSize        =   -1  'True
         Caption         =   "����������Ϣ"
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
         Left            =   480
         TabIndex        =   1
         Top             =   840
         Width           =   1710
      End
   End
   Begin VB.Label Label4 
      Caption         =   "��Ϣ�أ�"
      Height          =   375
      Left            =   6240
      TabIndex        =   6
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "���飺"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "��ǰ�׶Σ�"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "��ǰ�غϣ�"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
����������Ϣ.FontBold = False
����������Ϣ.FontBold = False
End Sub
Private Sub ����������Ϣ_Click()
Dim �غ� As Long
On Error GoTo Er:
�غ� = Val(InputBox("�������ڶ��ٻغϷ���һ��������Ϣ(������1�غϣ����δ��5�غ���)"))
If �غ� > ��ǰ�غ� + 5 Then �غ� = ��ǰ�غ� + 5
If �غ� <= ��ǰ�غ� Then �غ� = ��ǰ�غ� + 1
������Ϣ��(0, �غ�).���� = True
������Ϣ��(0, �غ�).�û��� = ������Ϣ����ѡȡ
MsgBox "�ڵڡ�" & �غ� & "���غϻ��һ����" & ������Ϣ��(0, �غ�).�û��� & "����Ϣ"
'If MsgBox("�Ƿ��֪ׯ�ң�", vbYesNo) = 6 Then
������������Ϣ���� "ׯ��", �غ� & "-" & ������Ϣ��(0, �غ�).�û���
'End If
Timer2_Timer
����������Ϣ.Enabled = False
����������Ϣ.FontBold = False
Exit Sub
Er:
MsgBox "����ʧ�ܣ�"
����������Ϣ.FontBold = False
End Sub
Private Sub ����������Ϣ_Click()
Dim �غ� As Long
On Error GoTo Er:
�غ� = Val(InputBox("�������ڶ��ٻغϷ���һ��������Ϣ(������1�غϣ����δ��5�غ���)"))
If �غ� > ��ǰ�غ� + 5 Then �غ� = ��ǰ�غ� + 5
If �غ� <= ��ǰ�غ� Then �غ� = ��ǰ�غ� + 1
������Ϣ��(0, �غ�).���� = True
������Ϣ��(0, �غ�).�û��� = ������Ϣ����ѡȡ
MsgBox "�ڵڡ�" & �غ� & "���غϻ��һ����" & ������Ϣ��(0, �غ�).�û��� & "����Ϣ"
'If MsgBox("�Ƿ��֪ׯ�ң�", vbYesNo) = 6 Then
������������Ϣ���� "ׯ��", �غ� & "-" & ������Ϣ��(0, �غ�).�û���
'End If
Timer2_Timer
����������Ϣ.Enabled = False
����������Ϣ.FontBold = False
Exit Sub
Er:
MsgBox "����ʧ�ܣ�"
����������Ϣ.FontBold = False
End Sub
Private Sub Timer1_Timer()
Label1.Caption = "��ǰ�غϣ�" & ��ǰ�غ�
Label2.Caption = "��ǰ�׶Σ�" & ��ǰ�׶�
Label3.Caption = "���飺" & ��������
If ����Ȩ�� = True Then �������� = "������˼����..."
If ��ǰ�׶� = "׼���׶�" Then
    Frame1.Enabled = True: ����������Ϣ = True: ����������Ϣ = True
    Frame2.Enabled = False: �ж��׶ε���ʱ.Enabled = False: ��ǰ����״̬.Enabled = False
Else
    Frame1.Enabled = False: ����������Ϣ = False: ����������Ϣ = False
    Frame2.Enabled = True: �ж��׶ε���ʱ.Enabled = True: ��ǰ����״̬.Enabled = True
End If
End Sub
Private Sub Timer2_Timer()
Dim i As Long
List1.Clear
For i = ��ǰ�غ� To ��ǰ�غ� + 5
    If ������Ϣ��(0, i).���� = True Then
        List1.AddItem "�ڵڡ�" & i & "���غ�" & vbTab & "������һ����" & ������Ϣ��(0, i).�û��� & "����Ϣ"
    End If
    If ������Ϣ��(0, i).���� = True Then
        List1.AddItem "�ڵڡ�" & i & "���غ�" & vbTab & "������һ����" & ������Ϣ��(0, i).�û��� & "����Ϣ"
    End If
Next
End Sub
Private Sub ����������Ϣ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
����������Ϣ.FontBold = True
End Sub
Private Sub ����������Ϣ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
����������Ϣ.FontBold = True
End Sub
