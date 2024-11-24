VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4875
   ClientLeft      =   225
   ClientTop       =   900
   ClientWidth     =   8820
   HelpContextID   =   1
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   8820
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   1680
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2040
      Width           =   4095
   End
   Begin VB.Menu x 
      Caption         =   "字形(&x)"
      Begin VB.Menu zx 
         Caption         =   "加粗"
         Index           =   0
      End
      Begin VB.Menu zx 
         Caption         =   "倾斜"
         Index           =   1
      End
      Begin VB.Menu zx 
         Caption         =   "下划线"
         Index           =   2
      End
   End
   Begin VB.Menu h 
      Caption         =   "字号(&h)"
      Begin VB.Menu s 
         Caption         =   "24"
         Index           =   0
      End
      Begin VB.Menu s 
         Caption         =   "30"
         Index           =   1
      End
      Begin VB.Menu s 
         Caption         =   "38"
         Index           =   2
      End
   End
   Begin VB.Menu m 
      Caption         =   "0"
      Visible         =   0   'False
      Begin VB.Menu q 
         Caption         =   "黑体"
         HelpContextID   =   2
         Index           =   0
         Shortcut        =   ^H
      End
      Begin VB.Menu q 
         Caption         =   "隶书"
         Index           =   1
         Shortcut        =   ^L
      End
      Begin VB.Menu q 
         Caption         =   "仿宋"
         Index           =   2
         Shortcut        =   ^F
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub q_Click(Index As Integer)
Text1.FontName = q(Index).Caption
End Sub

Private Sub s_Click(Index As Integer)
Select Case Index
Case 0
Text1.FontSize = 24
Case 1
Text1.FontSize = 30
Case 2
Text1.FontSize = 38
End Select
End Sub


Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 Then PopupMenu m
End Sub

Private Sub zx_Click(Index As Integer)
Select Case Index
Case 0
Text1.FontBold = True
Case 1
Text1.FontItalic = True
Case 2
Text1.FontUnderline = True
End Select
End Sub


