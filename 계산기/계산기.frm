VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "����"
   ClientHeight    =   4710
   ClientLeft      =   3930
   ClientTop       =   2565
   ClientWidth     =   5490
   Icon            =   "����.frx":0000
   LinkTopic       =   "Form1"
   Palette         =   "����.frx":5C12
   ScaleHeight     =   4710
   ScaleWidth      =   5490
   Begin VB.CommandButton Command8 
      Caption         =   "Back Space"
      Height          =   495
      Left            =   1200
      TabIndex        =   17
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "C"
      Height          =   495
      Left            =   3240
      TabIndex        =   16
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "="
      Height          =   2295
      Left            =   4080
      TabIndex        =   15
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "+"
      Height          =   495
      Index           =   1
      Left            =   3240
      TabIndex        =   14
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      Height          =   495
      Index           =   2
      Left            =   3240
      TabIndex        =   13
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "*"
      Height          =   495
      Index           =   3
      Left            =   3240
      TabIndex        =   12
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "/"
      Height          =   495
      Index           =   4
      Left            =   3240
      TabIndex        =   11
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '������ ����
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   240
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   240
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Index           =   9
      Left            =   2040
      TabIndex        =   9
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Index           =   8
      Left            =   1200
      TabIndex        =   8
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Command1"
      Height          =   495
      Index           =   7
      Left            =   360
      TabIndex        =   7
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Index           =   6
      Left            =   2040
      TabIndex        =   6
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Index           =   5
      Left            =   1200
      TabIndex        =   5
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Index           =   4
      Left            =   360
      TabIndex        =   4
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Index           =   3
      Left            =   2040
      TabIndex        =   3
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Index           =   2
      Left            =   1200
      TabIndex        =   2
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Command1"
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Support Corea IT Computer Academy"
      Height          =   495
      Left            =   2160
      TabIndex        =   19
      Top             =   4200
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   " Made By. KJH, JCW, JCM"
      Height          =   495
      Left            =   1800
      TabIndex        =   18
      Top             =   3960
      Width           =   3615
   End
   Begin VB.Menu ���� 
      Caption         =   "����"
      Begin VB.Menu ���� 
         Caption         =   "����"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim �Է�, x, k, c, ��� As Single

Private Sub ����_Click()
End
End Sub

Private Sub Command1_Click(index As Integer)
If Val(Text1.Text) = 0 Or k = 1 Then
If c > 0 Then
Text1.Text = Val(Text1.Text & index)
Else
Text1.Text = index
c = c + 1
End If
Else
Text1.Text = Val(Text1.Text & index)
End If
End Sub

Private Sub Command2_Click(index As Integer)
c = 0
x = index
If k = 0 Then
�Է� = Val(Text1.Text)
Else
    If x = 1 Then
    �Է� = �Է� + Val(Text1.Text)
    Text1.Text = �Է�
    ElseIf x = 2 Then
    �Է� = �Է� - Val(Text1.Text)
    Text1.Text = �Է�
    ElseIf x = 3 Then
    �Է� = �Է� * Val(Text1.Text)
    Text1.Text = �Է�
    ElseIf x = 4 Then
    �Է� = �Է� / Val(Text1.Text)
    Text1.Text = �Է�
    End If
End If
k = 1
End Sub

Private Sub Command6_Click()
k = 0

If x = 1 Then
��� = �Է� + Val(Text1.Text)
Text1.Text = ���
ElseIf x = 2 Then
��� = �Է� - Val(Text1.Text)
Text1.Text = ���
ElseIf x = 3 Then
��� = �Է� * Val(Text1.Text)
Text1.Text = ���
ElseIf x = 4 Then
��� = �Է� / Val(Text1.Text)
Text1.Text = ���
End If

End Sub

Private Sub Command7_Click()
k = 0
Text1.Text = ""
�Է� = 0
��� = 0
c = 0
End Sub

Private Sub Command8_Click()
If Len(Text1.Text) > 0 Then
Text1.Text = Left(Text1, Len(Text1) - 1)
End If
End Sub

Private Sub Form_DblClick()
Form2.Show
End Sub

Private Sub Form_Load()
c = 0
k = 0
For i = 0 To 9
Command1(i).Caption = i

Next
Text1.Text = ""
End Sub



Private Sub Text1_Change()
If Text1.Text = "" Then
Text1.Text = 0
End If
End Sub
