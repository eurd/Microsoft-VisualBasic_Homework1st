VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "������ʱ�亯��"
   ClientHeight    =   4425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   6195
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command9 
      Caption         =   "��"
      Height          =   495
      Left            =   2640
      TabIndex        =   9
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "����"
      Height          =   495
      Left            =   1560
      TabIndex        =   8
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "��ǰ���ڼ�"
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "��"
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "�·�"
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "��ǰʱ��"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ʱ"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "���"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ǰ����"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Label1 = Format(Date, "dddddd") '��ǰ����
End Sub

Private Sub Command2_Click()
Label1 = Year(Date) '���
End Sub

Private Sub Command3_Click()
Label1 = Hour(Time) 'ʱ
End Sub

Private Sub Command4_Click()
Label1 = Format(Time, "ttttt am/pm") '��ǰʱ��
End Sub

Private Sub Command5_Click()
Label1 = Month(Date) '�·�
End Sub

Private Sub Command6_Click()
Label1 = Minute(Time) '��
End Sub

Private Sub Command7_Click()
Label1 = Format(Date, "dddd") '��ǰ���ڼ�
End Sub

Private Sub Command8_Click()
Label1 = Day(Date) '����
End Sub

Private Sub Command9_Click()
Label1 = Second(Now) '��
End Sub
