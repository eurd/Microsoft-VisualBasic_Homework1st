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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "����ɼ�"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim sum&, max&, min&
a = InputBox("�����뿼����" & n + 1 & "���ɼ�", "�ɼ�", -1)
min = a
max = a
Do While a <> -1
Label1 = Label1 & a & " "
sum = sum + a
If a > max Then max = a
If a < min Then min = a
n = n + 1
a = InputBox("�����뿼����" & n + 1 & "���ɼ�", "�ɼ�", -1)
Loop
MsgBox "���÷�Ϊ" & (sum - min - max) / (n - 2)
End Sub
