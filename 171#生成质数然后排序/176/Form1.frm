VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "p"
   ClientHeight    =   4275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   5775
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture1 
      Height          =   2055
      Left            =   240
      ScaleHeight     =   1995
      ScaleWidth      =   4035
      TabIndex        =   6
      Top             =   720
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "最大值"
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "生成数"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "所在列："
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "所在行："
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "最大值："
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "所有生成数"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(), n, m
Private Sub Command1_Click()
Picture1.Cls
Dim i%, j%
n = InputBox("请输入行数", "行", 5)
m = InputBox("请输入列数", "列", 6)
ReDim a(n, m)
For i = 1 To n
For j = 1 To m
Randomize
a(i, j) = Int(Rnd * 91 + 10)
Picture1.Print ; a(i, j); " ";
Next j
Picture1.Print
Next i

End Sub

Private Sub Command2_Click()
For i = 1 To n
For j = 1 To m
If Max < a(i, j) Then
 Max = a(i, j): h = i: l = j
 End If
 Next j
 Next i
 Label3 = Label3 & Max
 Label4 = Label4 & h
 Label5 = Label5 & l

End Sub
