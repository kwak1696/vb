VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "»ö»ó/Å©±â Á¶Á¤ Æû"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5760
      TabIndex        =   10
      Text            =   "RGB(0,0,0)"
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   3015
      Left            =   120
      TabIndex        =   9
      Text            =   "ÅØ½ºÆ®"
      Top             =   2400
      Width           =   7695
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   1560
      Width           =   4455
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      Height          =   1455
      Left            =   5880
      ScaleHeight     =   1395
      ScaleWidth      =   1515
      TabIndex        =   6
      Top             =   360
      Width           =   1575
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Index           =   2
      Left            =   1200
      TabIndex        =   2
      Top             =   1120
      Width           =   4455
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Top             =   680
      Width           =   4455
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Å©±â"
      Height          =   495
      Left            =   600
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "ÆÄ¶û"
      Height          =   495
      Index           =   2
      Left            =   600
      TabIndex        =   5
      Top             =   1120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "ÃÊ·Ï"
      Height          =   495
      Index           =   1
      Left            =   600
      TabIndex        =   4
      Top             =   680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "»¡°­"
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
For i = 0 To 2
HScroll1(i).Min = 0
HScroll1(i).Max = 255
HScroll1(i).SmallChange = 1
HScroll1(i).LargeChange = 10
Next
HScroll2.Min = 10
HScroll2.Max = 127
HScroll2.SmallChange = 1
HScroll2.LargeChange = 10
End Sub

Private Sub HScroll1_Change(Index As Integer)
Dim »¡, ÃÊ, ÆÄ
»¡ = HScroll1(0).Value
ÃÊ = HScroll1(1).Value
ÆÄ = HScroll1(2).Value
Picture1.BackColor = RGB(»¡, ÃÊ, ÆÄ)
Text2.ForeColor = RGB(»¡, ÃÊ, ÆÄ)

Label1(0).ForeColor = RGB(»¡, 0, 0)
Label1(1).ForeColor = RGB(0, ÃÊ, 0)
Label1(2).ForeColor = RGB(0, 0, ÆÄ)

¹®ÀÚ = "RGB("
¹®ÀÚ = ¹®ÀÚ & Format(»¡, "##0") & ","
¹®ÀÚ = ¹®ÀÚ & Format(ÃÊ, "##0") & ","
¹®ÀÚ = ¹®ÀÚ & Format(ÆÄ, "##0") & ")"
Text1.Text = ¹®ÀÚ

End Sub

Private Sub HScroll2_Change()
Text2.FontSize = HScroll2.Value
End Sub

