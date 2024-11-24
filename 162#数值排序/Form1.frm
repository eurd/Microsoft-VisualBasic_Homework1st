VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   ScaleHeight     =   3540
   ScaleWidth      =   6765
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "排序"
      Height          =   735
      Left            =   2160
      TabIndex        =   8
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      Height          =   615
      Left            =   5160
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   3720
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   2280
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   5160
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   3720
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "三个数排序后"
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "请输入三个数"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a&, b&, c&
a = Text1: b = Text2: c = Text3
t = IIf(a > b, a, b)
Text4 = IIf(t > c, t, c)
t = IIf(a < b, a, b)
Text6 = IIf(t < c, t, c)
Text5 = a + b + c - Text4 - Text6
End Sub

