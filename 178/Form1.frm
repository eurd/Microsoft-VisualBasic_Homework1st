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
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1800
      Top             =   2160
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "»¶Ó­Ñ§Ï°VB!"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Randomize
Form1.BackColor = RGB(Rnd * 256, Rnd * 256, Rnd * 256)
a = Form1.BackColor
Label1.BackColor = a
Label1.FontSize = Int(Rnd * 26 + 12)

End Sub

Private Sub Timer1_Timer()
Randomize
Form1.BackColor = RGB(Rnd * 256, Rnd * 256, Rnd * 256)
a = Form1.BackColor
Label1.BackColor = a
Label1.FontSize = Int(Rnd * 26 + 12)

End Sub
