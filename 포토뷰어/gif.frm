VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9b.ocx"
Begin VB.Form out 
   BackColor       =   &H00FFFFFF&
   Caption         =   "뷰어"
   ClientHeight    =   5820
   ClientLeft      =   8145
   ClientTop       =   2925
   ClientWidth     =   6780
   Icon            =   "gif.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5820
   ScaleWidth      =   6780
   Visible         =   0   'False
   Begin VB.TextBox Text1 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   2760
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "1.0"
      Top             =   0
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "파일 열기"
      Height          =   200
      Left            =   120
      TabIndex        =   3
      Top             =   0
      UseMaskColor    =   -1  'True
      Value           =   1  '확인
      Width           =   1215
   End
   Begin VB.HScrollBar 불투명도 
      Height          =   1095
      LargeChange     =   2
      Left            =   4320
      Max             =   10
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4560
      Value           =   10
      Visible         =   0   'False
      Width           =   1215
   End
   Begin SHDocVwCtl.WebBrowser Web 
      Height          =   5535
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   6495
      ExtentX         =   11456
      ExtentY         =   9763
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Timer tmrKey 
      Interval        =   1
      Left            =   6480
      Top             =   0
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flash 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   6495
      _cx             =   11456
      _cy             =   9763
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "자동 넘김"
      Height          =   200
      Left            =   1680
      TabIndex        =   5
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.Image bmp 
      Height          =   5535
      Left            =   120
      Stretch         =   -1  'True
      Top             =   240
      Width           =   6495
   End
End
Attribute VB_Name = "out"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim loaded
Private Scroll As Boolean
Private Shift As Boolean
Private Caps As Boolean
Private KeyResult As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Sub AddKey(Key As String)
Text1 = Text1 & Key
Text1.SelStart = Len(Text1)
End Sub

Private Sub Check1_Click()
If Check1 = 1 Then
        f_open.Show
Else
    Unload f_open
End If
End Sub


Private Sub Check2_Click()
With f_open
If Check2 = 0 Then
.Timer1.Enabled = False
.Timer1.Interval = 0
Else
.Timer1.Enabled = True
.Timer1.Interval = .Timer1.Interval + Val(Text1.Text) * 1000
End If
End With
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
tmrKey.Enabled = True
End Sub

Private Sub Form_Load()
f_open.Show
End Sub
Private Sub File1_KeyPress(KeyAscii As Integer)
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
Private Sub 불투명도_Change()
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

Private Sub Form_Resize()
If Me.ScaleHeight - 240 > 0 And Me.ScaleWidth - 240 > 0 Then
    flash.Height = Me.ScaleHeight - 240
    flash.Width = Me.ScaleWidth - 240
    Web.Height = Me.ScaleHeight - 240
    Web.Width = Me.ScaleWidth - 240
    bmp.Height = Me.ScaleHeight - 240
    bmp.Width = Me.ScaleWidth - 240
End If
End Sub

Private Sub Text1_Change()
Check2 = 0
End Sub

Private Sub tmrKey_Timer()
On Error Resume Next

KeyResult = GetAsyncKeyState(37)
    If KeyResult = -32767 Then
        f_open.File1.ListIndex = f_open.File1.ListIndex - 1
        Exit Sub
    End If
    
KeyResult = GetAsyncKeyState(39)
    If KeyResult = -32767 Then
        f_open.File1.ListIndex = f_open.File1.ListIndex + 1
    Exit Sub
End If

KeyResult = GetAsyncKeyState(13)
    If KeyResult = -32767 Then
    퀵
    Exit Sub
    End If

KeyResult = GetAsyncKeyState(221)
    If KeyResult = -32767 Then
    불투명도.Value = 불투명도.Value + 1
    End If
    
KeyResult = GetAsyncKeyState(219) '219
    If KeyResult = -32767 Then
    불투명도.Value = 불투명도.Value - 1
    Exit Sub
    End If
End Sub
