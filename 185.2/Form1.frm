VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   6600
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "添加"
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   2370
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":0016
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   2415
      Left            =   3840
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "购物车"
      Height          =   615
      Left            =   3960
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "选购列表"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i&
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
Text2 = Text2 & "第" & i + 1 & "项: " & List1.List(i) & vbCrLf
End If
Next i

End Sub
     
