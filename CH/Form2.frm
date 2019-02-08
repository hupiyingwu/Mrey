VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "网络配置"
   ClientHeight    =   3960
   ClientLeft      =   10470
   ClientTop       =   6930
   ClientWidth     =   7680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3960
   ScaleWidth      =   7680
   Begin VB.CommandButton Command1 
      Caption         =   "连接"
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   2280
      TabIndex        =   3
      Text            =   "main"
      Top             =   1800
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   2280
      TabIndex        =   1
      Text            =   "127.0.0.1:105"
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "网络名称："
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "已知节点地址："
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   1320
      Width           =   1815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Form1.main_nood = Text1.Text
Form1.chain_name = Text2.Text
Form1.Show
Unload Me

End Sub
