VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4140
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   6450
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2760
      TabIndex        =   3
      Top             =   2760
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2760
      TabIndex        =   2
      Text            =   "15"
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label2 
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "输入一个十进制数:"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Menu qw 
      Caption         =   "进制转换"
      Index           =   12
      NegotiatePosition=   1  'Left
      Begin VB.Menu er 
         Caption         =   "二进制"
         Shortcut        =   ^E
      End
      Begin VB.Menu ba 
         Caption         =   "八进制"
         Shortcut        =   ^O
      End
      Begin VB.Menu sl 
         Caption         =   "十六进制"
         Shortcut        =   ^H
      End
      Begin VB.Menu ed 
         Caption         =   "退出"
      End
   End
   Begin VB.Menu eds 
      Caption         =   "重置"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub eds_Click()
Text1 = ""
Text2 = ""
End Sub
Private Sub er_Click()
Label2.Caption = "转换成二进制为:"
n = Text1.Text
    Do
    r = n Mod 2
    s = CStr(r) & s
    n = n \ 2
    Loop Until n = 0
    Text2.Text = s
End Sub

Private Sub ba_Click()
Label2.Caption = "转换成八进制为:"
Text2.Text = Oct(Val(Text1))
End Sub


Private Sub sl_Click()
Label2.Caption = "转换成十六进制为:"
Text2.Text = Hex(Val(Text1))
End Sub

Private Sub ed_Click()
End
End Sub
