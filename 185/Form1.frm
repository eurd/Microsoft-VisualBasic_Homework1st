VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List1 
      Height          =   1500
      ItemData        =   "Form1.frx":0000
      Left            =   1080
      List            =   "Form1.frx":0016
      TabIndex        =   0
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   2760
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub List1_Click()
Label1 = "您的专业是: 第" & List1.ListIndex + 1 & "项" & List1.Text
End Sub
