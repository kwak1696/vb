VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Å©±â °íÁ¤ ´ëÈ­ »óÀÚ
   Caption         =   "Fastest Click"
   ClientHeight    =   2265
   ClientLeft      =   4965
   ClientTop       =   4530
   ClientWidth     =   4185
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4185
   Begin VB.CommandButton Command1 
      Caption         =   "Click Start"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2880
      Top             =   480
   End
   Begin VB.Label Label6 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "10 Sec"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Time Limit"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   1485
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "0"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   1170
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   870
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "0"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   555
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackStyle       =   0  'Åõ¸í
      Caption         =   "High Score"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   2040
      Left            =   120
      Picture         =   "main.frx":1942
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2715
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b
Private Sub Command1_Click()
Timer1.Enabled = True
Command1.Visible = False
Image1.Enabled = True
End Sub

Private Sub Image1_Click()
b = b + 1
Label4 = b
End Sub

Private Sub Timer1_Timer()
a = a + 1
Label6 = Format(10 - a, "0 Sec")
If a = 10 Then
If Val(Label2) < Val(Label4) Then Label2 = Label4
MsgBox "You did " & b & " cliks!!!", vbInformation, "Time Out"
Timer1.Enabled = False
Command1.Visible = True
Image1.Enabled = False
Label6 = "10 Sec"
b = 0
a = 0
End If
End Sub
