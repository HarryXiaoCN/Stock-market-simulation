VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��Ҷ�"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5520
   BeginProperty Font 
      Name            =   "΢���ź�"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   5520
   StartUpPosition =   2  '��Ļ����
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   3495
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   5295
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   5040
      Top             =   5640
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�Ͽ�"
      Height          =   375
      Left            =   2815
      TabIndex        =   5
      Top             =   1080
      Width           =   2600
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Text            =   "8888"
      Top             =   600
      Width           =   3855
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4560
      Top             =   5640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2600
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label4 
      Caption         =   "��ǰ�г�����ң�"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label3 
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   5640
      Width           =   4095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "�������˿ڣ�"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��������ַ��"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1440
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo Er
Winsock1.RemoteHost = Text1.Text
Winsock1.RemotePort = Text2.Text
Winsock1.Connect
Exit Sub
Er:
MsgBox "����ʧ�ܣ�������������������Ƿ�ͨ�����ظ����ӣ�"
End Sub
Private Sub Command2_Click()
Winsock1.Close
End Sub

Private Sub Form_Load()
Local_State = 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
Winsock1.Close
Local_State = 0
Form1.Show
End Sub
Private Sub Timer1_Timer()
If Winsock1.State = 7 Then Label3.Caption = "������;������IP��" & Winsock1.RemoteHostIP Else Label3.Caption = ����״̬����(Winsock1.State)
If Winsock1.State = 9 Then Winsock1.Close
End Sub
Private Sub Winsock1_Close()
Winsock1.Close
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Winsock1.GetData Client_GetData_tcpTemp
If ClientGetDataShow = True Then Form2.Text1.Text = Client_GetData_tcpTemp
Client_GetData
End Sub
