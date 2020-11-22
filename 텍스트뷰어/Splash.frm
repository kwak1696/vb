VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   ClientHeight    =   3495
   ClientLeft      =   1260
   ClientTop       =   2520
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Splash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   120
   End
   Begin VB.Frame Frame1 
      Height          =   3315
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   360
         Picture         =   "Splash.frx":030A
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblWarning 
         Caption         =   "��� : �� ���α׷����� ���� ���ش� å������ �ʽ��ϴ�."
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   2880
         Width           =   6855
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  '������ ����
         AutoSize        =   -1  'True
         Caption         =   "Ver 1.0"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6015
         TabIndex        =   2
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Photo Viewer"
         BeginProperty Font 
            Name            =   "����"
            Size            =   32.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   2520
         TabIndex        =   4
         Top             =   1140
         Width           =   4095
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "G Community"
         BeginProperty Font 
            Name            =   "����"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2355
         TabIndex        =   3
         Top             =   705
         Width           =   2550
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    lblVersion.Caption = "���� " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i
For i = 1 To 0 Step -0.005
    Module1.MakeLayeredWnd Me.hWnd
    SetLayeredWindowAttributes Me.hWnd, 0, 255 * (i), LWA_ALPHA
Next
out.Show
End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub
