VERSION 5.00
Begin VB.Form 走势图 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "走势图"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12765
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   12765
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   120
      ScaleHeight     =   6345
      ScaleWidth      =   12465
      TabIndex        =   0
      Top             =   600
      Width           =   12495
   End
End
Attribute VB_Name = "走势图"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
