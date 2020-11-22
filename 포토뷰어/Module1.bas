Attribute VB_Name = "Module1"
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
Public Const LWA_ALPHA = &H2
'불투명도 조절 선언
Public full As String
Public 파일 As String
Public 체크 As Integer
Public Function MakeLayeredWnd(hWnd As Long) As Long
    Dim WndStyle As Long
    WndStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
    WndStyle = WndStyle Or WS_EX_LAYERED
    MakeLayeredWnd = SetWindowLong(hWnd, GWL_EXSTYLE, WndStyle)
End Function


