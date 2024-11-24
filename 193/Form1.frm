VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3570
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   4320
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2040
      Top             =   2760
   End
   Begin VB.CommandButton Command2 
      Caption         =   "暂停"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      Height          =   495
      Index           =   1
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   735
   End
   Begin VB.Shape Shape2 
      Height          =   495
      Index           =   0
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   360
      Width           =   735
   End
   Begin VB.Shape Shape2 
      Height          =   495
      Index           =   2
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   1800
      Width           =   735
   End
   Begin VB.Shape Shape1 
      Height          =   2175
      Left            =   1920
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t, a, b
Private Sub Command1_Click()
Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
End Sub


Private Sub Form_Activate()
For i = 0 To 2
Shape2(i).FillStyle = 0
Next i
End Sub



Private Sub Timer1_Timer()
t = t + 1
If t > 0 And t < 7 Then
Shape2(0).BorderColor = RGB(255, 0, 0)
Shape2(0).FillColor = RGB(255, 0, 0)
Shape2(2).BorderColor = RGB(0, 0, 0)
Shape2(2).FillColor = RGB(0, 0, 0)
End If
If t > 7 And t < 9 Then
Shape2(0).BorderColor = RGB(0, 0, 0)
Shape2(0).FillColor = RGB(0, 0, 0)
Shape2(1).BorderColor = RGB(255, 255, 0)
Shape2(1).FillColor = RGB(255, 255, 0)
End If
If t > 9 Then
Shape2(1).BorderColor = RGB(0, 0, 0)
Shape2(1).FillColor = RGB(0, 0, 0)
Shape2(2).BorderColor = RGB(0, 255, 0)
Shape2(2).FillColor = RGB(0, 255, 0)
If t > 11 Then t = 0

End If
End Sub
