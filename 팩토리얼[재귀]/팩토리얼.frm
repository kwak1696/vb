VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "재귀함수 연습"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4275
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   4275
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox Text3 
      Height          =   2775
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   3
      Text            =   "팩토리얼.frx":0000
      Top             =   1320
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Text            =   "0"
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "팩토리얼 계산"
      Height          =   1095
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    input_data = Val(Text2.Text)
    output_data = hamsu(input_data)
    Text1.Text = "= " & output_data
End Sub

Function hamsu(x)
    If x = 1 Then
        hamsu = 1
    Else
        hamsu = x * hamsu(x - 1)
    End If
End Function
