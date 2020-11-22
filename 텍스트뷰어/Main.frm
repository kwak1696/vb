VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Text Viewer"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6255
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  '수동
   ScaleHeight     =   7680
   ScaleWidth      =   6255
   StartUpPosition =   2  '화면 가운데
   Begin RichTextLib.RichTextBox 내용 
      Height          =   4935
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   8705
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      RightMargin     =   1
      AutoVerbMenu    =   -1  'True
      OLEDragMode     =   0
      OLEDropMode     =   1
      TextRTF         =   $"Main.frx":030A
   End
   Begin VB.CommandButton Command2 
      Height          =   255
      Left            =   5640
      Picture         =   "Main.frx":03A7
      Style           =   1  '그래픽
      TabIndex        =   7
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Left            =   5640
      Picture         =   "Main.frx":19FF9
      Style           =   1  '그래픽
      TabIndex        =   6
      Top             =   1800
      Width           =   495
   End
   Begin VB.FileListBox File1 
      Height          =   2250
      Hidden          =   -1  'True
      Left            =   2880
      Pattern         =   "*.txt;*smi;*.html;*.log;*.c;*.frm"
      System          =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      Height          =   1980
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Directory"
      Top             =   240
      Width           =   2655
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Drive"
      Top             =   0
      Width           =   2655
   End
   Begin MSComctlLib.Slider 불투명도 
      CausesValidation=   0   'False
      Height          =   1335
      Left            =   5640
      TabIndex        =   0
      ToolTipText     =   "불투명도 설정"
      Top             =   0
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   2355
      _Version        =   393216
      Orientation     =   1
      SelStart        =   10
      TickStyle       =   3
      TickFrequency   =   10
      Value           =   10
   End
   Begin VB.Label 불투명도_표시 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "10"
      Height          =   255
      Left            =   5520
      TabIndex        =   5
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "설명"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   5295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub 내용_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
    퀵
    End If
End Sub
Sub 퀵()
    If 불투명도.Value = 10 Then
    불투명도.Value = 0
    Else
    불투명도.Value = 10
    End If
End Sub

Private Sub 내용_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1 = Data.Files(1)
    출력 (Data.Files(1))
End Sub

Private Sub 불투명도_Change()
    If 불투명도 = 0 Then 내용.SetFocus
    불투명도_표시 = 불투명도.Value
    불투명도설정 (불투명도.Value / 10)
End Sub

Private Sub 불투명도_Scroll()
    불투명도_표시 = 불투명도.Value
    불투명도설정 (불투명도.Value / 10)
End Sub

Sub 불투명도설정(X As Double)
    Module1.MakeLayeredWnd Me.hWnd
    SetLayeredWindowAttributes Me.hWnd, 0, 255 * (X), LWA_ALPHA
End Sub

Private Sub Command1_Click()
If File1.ListCount = 0 Then Exit Sub
폴더저장
Open App.Path & "\" & "book.log" For Output As #3
Print #3, File1.FileName
Print #3, 내용.SelStart
Close #3
End Sub

Private Sub Command2_Click()
    Open App.Path & "\" & "book.log" For Input As #4
    Line Input #4, 파일이름
    Line Input #4, 북마크
    폴더열기
    출력 (full & 파일이름)
    내용.SetFocus
    내용.SelStart = 북마크
    Close #4
End Sub

Private Sub Dir1_Change()
On Error Resume Next
    File1.Path = Dir1.Path
    Label1 = File1.ListCount & "개의 텍스트파일이 있습니다."
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
    폴더경로
    출력 (full & File1.FileName)
End Sub

Private Sub Form_Activate()
Dir1.Path = "C:\"
폴더열기
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    출력 (Data.Files(1))
End Sub

Sub 출력(full As String)
On Error Resume Next
    Label1 = full
    내용.LoadFile full
End Sub

Private Sub Form_Resize()
    내용.Width = Me.ScaleWidth - 250
End Sub

Private Sub Form_Unload(Cancel As Integer)
    폴더저장
End Sub

Sub 폴더저장()
    Open App.Path & "\" & "last.log" For Output As #1
    폴더경로
    Print #1, full
    Close #1
End Sub
Sub 폴더열기()
    On Error Resume Next
    Open App.Path & "\" & "last.log" For Input As #2
    Line Input #2, full
    Dir1.Path = full
    Drive1.Drive = Left(Dir1.Path, 3)
    Close #2
End Sub
