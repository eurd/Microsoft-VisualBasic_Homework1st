VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   5610
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   735
      Left            =   1440
      TabIndex        =   2
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2160
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "登录密码"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1 = "123" Then
Form1.Hide
Form2.Show
Else
MsgBox "密码错误，请重新输入！", 48, "提示"
Text1.SetFocus
End If
End Sub
