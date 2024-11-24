VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   5835
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List2 
      Height          =   2940
      ItemData        =   "Form1.frx":0000
      Left            =   3120
      List            =   "Form1.frx":0002
      TabIndex        =   7
      Top             =   840
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   2760
      ItemData        =   "Form1.frx":0004
      Left            =   0
      List            =   "Form1.frx":001A
      TabIndex        =   6
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "全部移除"
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "移除"
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "全部添加"
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "添加"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "拟选"
      Height          =   615
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "备选"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i&
Private Sub Command1_Click()
Dim j&
 For i = 0 To List1.ListCount - 1
  If List1.Selected(i) Then
   For j = 0 To List2.ListCount - 1
    If List2.List(j) = List1.List(i) Then Exit For
   Next j
    If j > List2.ListCount - 1 Then
     List2.AddItem List1.List(i)
  End If
    End If
 Next i
 If List2.ListCount > 0 Then
   Command3.Enabled = True
   Command4.Enabled = True
 End If
End Sub



Private Sub Command2_Click()
List2.Clear
For i = 0 To List1.ListCount - 1
List2.AddItem List1.List(i)
Next i
  Command1.Enabled = False
  Command2.Enabled = False
  Command3.Enabled = True
  Command4.Enabled = True
End Sub

Private Sub Command3_Click()
Dim a As Boolean
a = True
Do While a And List2.ListCount > 0
For i = List2.ListCount - 1 To 0 Step -1
If List2.Selected(i) Then
List2.RemoveItem List2.ListIndex
a = True
Else
a = False
End If
Next i
Loop
If List2.ListCount > 0 Then
Command3.Enabled = True
  Command4.Enabled = True
  Else
  Command3.Enabled = False
   Command4.Enabled = False
End If
Command1.Enabled = True
   Command2.Enabled = True
End Sub

Private Sub Command4_Click()
List2.Clear
   Command1.Enabled = True
   Command2.Enabled = True
   Command3.Enabled = False
   Command4.Enabled = False
End Sub
