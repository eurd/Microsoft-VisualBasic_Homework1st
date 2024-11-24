VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12510
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   12510
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   840
      ScaleHeight     =   2835
      ScaleWidth      =   9675
      TabIndex        =   0
      Top             =   840
      Width           =   9735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
Dim i&, j&
Picture1.Print Tab(4);
For i = 1 To 9
Picture1.Print Format(i, "!@@@@@@@");
Next i
Picture1.Print
For i = 1 To 9
Picture1.Print Format(i, "!@@@");
For j = 1 To i
Picture1.Print j & "*" & i&; "="; Format(j * i, "!@@@");
Next j
Picture1.Print
Next i

End Sub

