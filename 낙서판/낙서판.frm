VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "낙서판"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   9300
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   2520
      TabIndex        =   14
      Text            =   "10"
      Top             =   3600
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "색 바꾸기"
      Height          =   855
      Left            =   1080
      TabIndex        =   6
      Top             =   4080
      Width           =   7935
      Begin VB.Label Label2 
         BackColor       =   &H00800080&
         Height          =   375
         Index           =   5
         Left            =   7080
         TabIndex        =   13
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000FFFF&
         Height          =   375
         Index           =   4
         Left            =   5712
         TabIndex        =   12
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label2 
         BackColor       =   &H00008000&
         Height          =   375
         Index           =   3
         Left            =   4344
         TabIndex        =   11
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF0000&
         Height          =   375
         Index           =   2
         Left            =   2976
         TabIndex        =   10
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000FF&
         Height          =   375
         Index           =   1
         Left            =   1608
         TabIndex        =   9
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "저장"
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "복사"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "지우기"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "텍스트 입력"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   3240
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      ScaleHeight     =   2835
      ScaleWidth      =   8835
      TabIndex        =   0
      Top             =   120
      Width           =   8895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   4320
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String
Dim k As Long

Private Sub Command1_Click()
Picture1.Print Text1
End Sub

Private Sub Command2_Click()
Picture1.Cls
Text1.Text = ""
End Sub

Private Sub Command3_Click()
Clipboard.SetData Picture1.Image
MsgBox "복사되었습니다", vbOKOnly, "복사"
End Sub

Private Sub Command4_Click()

a = InputBox("저장할 위치를 선택하세요", "저장", "C:\낙서판.bmp")
SavePicture Picture1.Image, a
MsgBox a & "로 저장되었습니다.", vbOKOnly, "저장"
End Sub


Private Sub Form_Load()
k = 10
Picture1.AutoRedraw = True
End Sub

Private Sub Label1_Click()
Picture1.ForeColor = vbBlack
End Sub

Private Sub Label2_Click(Index As Integer)
Select Case Index
Case Is = 0
Picture1.ForeColor = vbWhite
Case Is = 1
Picture1.ForeColor = vbRed
Case Is = 2
Picture1.ForeColor = vbBlue
Case Is = 3
Picture1.ForeColor = vbGreen
Case Is = 4
Picture1.ForeColor = vbYellow
Case Is = 5
Picture1.ForeColor = &H800080
End Select
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Picture1.Line (x, Y)-(x, Y)

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 Then
Picture1.Line -(x, Y)
ElseIf Button = 2 Then
Picture1.Circle (x, Y), k * 5
End If

End Sub

Private Sub Text2_Change()
If Text2 = "" Then
Text2 = 10
End If
k = Val(Text2)

Picture1.FontSize = k
End Sub
