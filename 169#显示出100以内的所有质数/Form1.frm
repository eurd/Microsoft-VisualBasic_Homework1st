VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "输出100以内所有质数"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   2295
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "100以内所有质数："
      Height          =   495
      Left            =   480
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
Private Sub Command1_Click()
Dim a&, i%
For i = 3 To 100
For a = 2 To Sqr(i)
If i Mod a = 0 Then Exit For
Next a
If a > Sqr(i) Then
Text1 = Text1 & i & vbCrLf
End If
Next i
End Sub
