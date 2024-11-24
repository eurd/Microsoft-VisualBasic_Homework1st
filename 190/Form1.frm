VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   9720
   StartUpPosition =   3  '窗口缺省
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Index           =   2
      Left            =   960
      TabIndex        =   4
      Top             =   3840
      Width           =   6855
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   3
      Top             =   3120
      Width           =   6855
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   1
      Top             =   2400
      Width           =   6855
   End
   Begin VB.Label Label3 
      Height          =   375
      Index           =   0
      Left            =   7920
      TabIndex        =   9
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label3 
      Height          =   375
      Index           =   2
      Left            =   7920
      TabIndex        =   8
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label3 
      Height          =   375
      Index           =   1
      Left            =   7920
      TabIndex        =   7
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "蓝色"
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "绿色"
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "红色"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label1 
      Height          =   1815
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   6015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub HScroll1_Change(Index As Integer)
HScroll1(Index).Max = 255
Label1.BackColor = RGB(HScroll1(0).Value, HScroll1(1).Value, HScroll1(2).Value)
Label3(Index) = HScroll1(Index).Value
End Sub
