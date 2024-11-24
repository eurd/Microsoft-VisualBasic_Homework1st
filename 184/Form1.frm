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
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "Form1.frx":0000
      Left            =   960
      List            =   "Form1.frx":0016
      TabIndex        =   0
      Text            =   "请选择专业"
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   2520
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Combo1_Click()
Label1 = "您的专业是: " & Combo1.Text
End Sub
