VERSION 5.00
Begin VB.Form f_open 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "Photo Viewer"
   ClientHeight    =   2760
   ClientLeft      =   2475
   ClientTop       =   2790
   ClientWidth     =   5685
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   5280
      Top             =   0
   End
   Begin VB.FileListBox File1 
      Height          =   2250
      Hidden          =   -1  'True
      Left            =   2760
      Pattern         =   "*.jpg;*.bmp;*.gif;*.png;*.swf"
      System          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      Height          =   1980
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Directory"
      Top             =   375
      Width           =   2655
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Drive"
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "설명"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   5295
   End
End
Attribute VB_Name = "f_open"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub 폴더저장()
    Open App.Path & "\" & "last.log" For Output As #1
    폴더경로
    Print #1, full
    Close #1
End Sub
Sub 폴더열기()
    'On Error Resume Next
    Open App.Path & "\" & "last.log" For Input As #2
    Line Input #2, full
    Drive1.Drive = Left(Dir1.Path, 2)
    Dir1.Path = full
    Close #2
End Sub
Sub 출력()
On Error Resume Next
    Label1 = full & 파일
    If Right(파일, 3) = "bmp" Then
        out.flash.Visible = False
        out.Web.Visible = False
        out.bmp.Picture = LoadPicture(full & 파일)
    ElseIf Right(파일, 3) = "gif" Then
        out.flash.Visible = False
         out.Web.Visible = True
        out.Web.Navigate full & 파일
    Else
        out.Web.Visible = False
        out.flash.Visible = True
        out.flash.Movie = full & 파일
    End If
        out.Visible = True
End Sub

Private Sub Dir1_Change()
On Error Resume Next
    File1.Path = Dir1.Path
    If File1.ListCount = 0 Then
    Label1 = "그림파일이 없습니다."
    Else
    Label1 = File1.ListCount & "개의 그림파일이 있습니다."
    End If
End Sub

Private Sub Drive1_Change()
On Error Resume Next
    Dir1.Path = Drive1.Drive
End Sub
Sub 폴더경로()
    full = UCase(Left(File1.Path, 1)) & Mid(File1.Path, 2)
    If Right(full, 1) <> "\" Then full = full & "\"
End Sub
Private Sub File1_Click()
On Error Resume Next
    폴더경로
    파일 = File1.FileName
    출력
End Sub

Private Sub Form_Load()
폴더열기
End Sub

Private Sub Form_Unload(Cancel As Integer)
out.Check1.Value = 0
    폴더저장
End Sub
Private Sub Timer1_Timer()
If f_open.File1.ListIndex = -1 Then Exit Sub
If f_open.File1.ListCount = f_open.File1.ListIndex + 1 Then
f_open.File1.ListIndex = 0
Else
f_open.File1.ListIndex = f_open.File1.ListIndex + 1
End If
End Sub

