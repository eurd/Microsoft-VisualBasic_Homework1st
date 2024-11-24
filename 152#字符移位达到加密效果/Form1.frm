VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "解密"
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "加密"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "密文"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "明文"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
For i = 1 To Len(Text1)
a = Mid(Text1, i, 1)
If a < "v" Then
s = s + Chr(Asc(a) + 5)
Else
s = s + Chr(Asc(a) + 5 - 26)
End If
Next i
Text2 = s
End Sub

Private Sub Command2_Click()
For i = 1 To Len(Text2)
a = Mid(Text2, i, 1)
If a > "e" Then
s = s + Chr(Asc(a) - 5)
Else
s = s + Chr(Asc(a) - 5 + 26)
End If
Next i
Text1 = s
End Sub
