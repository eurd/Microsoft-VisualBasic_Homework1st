VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   6195
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "保留两位小数"
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "开平方"
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "产生随机数"
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   3000
      TabIndex        =   2
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "  →"
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Randomize
Text1 = Int(Rnd * 900 + 100)
End Sub

Private Sub Command2_Click()
Text2 = Sqr(Text1)
End Sub

Private Sub Command3_Click()
Text2 = Round(Val(Text2), 2)
End Sub
