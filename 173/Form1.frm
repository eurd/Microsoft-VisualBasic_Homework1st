VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   6975
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "ת��"
      Height          =   615
      Left            =   1920
      TabIndex        =   4
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   1680
      TabIndex        =   3
      Top             =   960
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "�������д���"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "������Сд���"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim b
Text2 = ""
b = Array("��", "Ҽ", "��", "��", "��", "��", "½", "��", "��", "��")
n = Len(Text1)
a = Text1
For i = n To 1 Step -1
m = Val(Mid(a, n - i + 1, 1))
Select Case i
Case 4
Text2 = Text2 & b(m) & "Ǫ"
Case 3
Text2 = Text2 & b(m) & "��"
Case 2
Text2 = Text2 & b(m) & "ʰ"
Case 1
Text2 = Text2 & b(m) & "Ԫ"
End Select
Next i

End Sub
