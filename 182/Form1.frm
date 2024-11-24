VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5340
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   9270
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   2655
      Left            =   6120
      TabIndex        =   3
      Top             =   1920
      Width           =   2415
      Begin VB.OptionButton Option2 
         Caption         =   "36"
         Height          =   375
         Index           =   2
         Left            =   480
         TabIndex        =   12
         Top             =   1680
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         Caption         =   "28"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         Caption         =   "24"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   2775
      Left            =   3240
      TabIndex        =   2
      Top             =   1920
      Width           =   2535
      Begin VB.CheckBox Check1 
         Caption         =   "下划线"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   2295
      End
      Begin VB.CheckBox Check1 
         Caption         =   "倾斜"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   2295
      End
      Begin VB.CheckBox Check1 
         Caption         =   "加粗"
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   7
         Top             =   480
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3015
      Left            =   480
      TabIndex        =   1
      Top             =   1920
      Width           =   2415
      Begin VB.OptionButton Option1 
         Caption         =   "隶书"
         BeginProperty Font 
            Name            =   "隶书"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   600
         TabIndex        =   6
         Top             =   2040
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "仿宋"
         BeginProperty Font 
            Name            =   "华文仿宋"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   600
         TabIndex        =   5
         Top             =   1080
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "黑体"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   600
         TabIndex        =   4
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Label Label1 
      Caption         =   "欢迎"
      Height          =   1095
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   5415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click(Index As Integer)
Select Case Index
Case (0)
Label1.FontBold = Check1(0).Value
Case (1)
Label1.FontItalic = Check1(1).Value
Case (2)
Label1.FontUnderline = Check1(2).Value
End Select

End Sub

Private Sub Option1_Click(Index As Integer)
Label1.FontName = Option1(Index).Caption


End Sub

Private Sub Option2_Click(Index As Integer)
Select Case Index
Case (0)
Label1.FontSize = 24
Case (1)
Label1.FontSize = 28
Case (2)
Label1.FontSize = 32
End Select


End Sub
