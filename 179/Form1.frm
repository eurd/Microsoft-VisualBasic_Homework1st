VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   7155
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      Height          =   735
      Left            =   1920
      TabIndex        =   4
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   4320
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "所有数的平均值为："
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "请输入0-100之间的数："
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "输入的数："
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(), n%
Private Sub Command1_Click()
Dim i%, sum%
For i = 1 To n
  sum = sum + a(i)
  Next i
  Label3.Caption = Label3.Caption & Format(sum / n, "0.00") & " "
  
End Sub
Private Sub text1_keypress(keyascii As Integer)
If keyascii = 13 Then
n = n + 1
ReDim Preserve a(n)
a(n) = Val(Text1.Text)
Label1.Caption = Label1.Caption & Text1 & " "
Text1.SetFocus
Text1.SelLength = Len(Text1)
Text1 = ""
End If
End Sub
