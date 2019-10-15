VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   3135
   ClientLeft      =   1425
   ClientTop       =   -615
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   Begin VB.Timer Timer2 
      Left            =   3120
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   3000
      Top             =   960
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   -360
      Picture         =   "悬浮窗.frx":0000
      ScaleHeight     =   2535
      ScaleWidth      =   2895
      TabIndex        =   0
      Top             =   -480
      Width           =   2895
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long '定义API
Private Const HWND_TOPMOST = -1 '窗体最上层常数
Private Const SWP_NOMOVE = &H2 '窗体不移动（保持当前位置（x和y设定将被忽略））
Private Const SWP_NOSIZE = &H1 '窗体大小不变（保持当前大小（cx和cy会被忽略））
 Private oShadow As New aShadow
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1


Private Sub Form_Load()
With oShadow
    If .Shadow(Me) Then
        .Depth = 0 '阴影宽度
        .Color = RGB(0, 0, 0) '阴影颜色
        .Transparency = 0 '阴影色深
    End If
 End With
Dim rtn As Long

 rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
 rtn = rtn Or WS_EX_LAYERED
 SetWindowLong hwnd, GWL_EXSTYLE, rtn
 SetLayeredWindowAttributes hwnd, vbWhite, 0, LWA_COLORKEY
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Form5.Timer1.Enabled = True
    Form5.Timer1.Interval = 5
    Form5.Timer2.Enabled = True
    Form5.Timer2.Interval = 5
    Variable
End Sub

Private Sub Picture1_Click()
Form1.Show
Form5.Hide
End Sub

Private Sub Picture1_Mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0
End If
End Sub

Private Sub Timer1_Timer()
    GetFormSide
    GetMouseSide
End Sub

Private Sub Timer2_Timer()

    MouseState
    MoveUpForm
    MoveDownForm
End Sub
