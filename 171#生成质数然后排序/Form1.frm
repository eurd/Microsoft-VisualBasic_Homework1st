VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   3405
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   5565
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "生成"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "排序"
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label4 
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "排序"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label2 
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "生成"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(1 To 10)
Private Sub Command1_Click()
Label2 = ""
Label4 = ""
Randomize
For i = 1 To 10
a(i) = Int(Rnd * 91 + 10)
Label2 = Label2 & a(i) & " "
Next i
End Sub

Private Sub Command2_Click()
For i = 1 To 9
For j = 1 To 10 - i
If a(j) < a(j + 1) Then
t = a(j): a(j) = a(j + 1): a(j + 1) = t
End If
Next j
Next i
For y = 1 To 10
r = r & a(y) & " "
Next y
Label4 = r
End Sub
