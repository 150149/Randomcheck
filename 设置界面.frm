VERSION 5.00
Object = "{A5804A5B-13E1-4641-9440-19656D6B4A8E}#1.0#0"; "P控件集.ocx"
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "第一次设置程序"
   ClientHeight    =   3135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "华文琥珀"
      Size            =   9
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   2  '屏幕中心
   Begin P控件集.PButton PButton1 
      Height          =   1215
      Left            =   3480
      TabIndex        =   7
      Top             =   840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   2143
      Color_Back      =   33023
      Color_Back_Down =   33023
      Color_Begin     =   33023
      Color_End       =   16576
      Color_Text      =   16777215
      Text            =   "确定"
      Style_Border    =   0
      Color_Border    =   33023
      Can_Text_Move   =   0   'False
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      MaxLength       =   3
      TabIndex        =   1
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      MaxLength       =   3
      TabIndex        =   0
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H000080FF&
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H000080FF&
      Caption         =   "▁▁"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label3 
      BackColor       =   &H000080FF&
      Caption         =   "第一次设置程序"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label Label2 
      Caption         =   "随机数最小值"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "随机数最大值"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
 Private oShadow As New aShadow
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Const MAX_PATH As Integer = 260
Const TH32CS_SNAPPROCESS As Long = 2&
Private Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
   th32ParentProcessID As Long
   pcPriClassBase As Long
   dwFlags As Long
    szExeFile As String * 1024
    End Type
Private Declare Function FINDWINDOW Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long


Private Sub Form_Load()

With oShadow
    If .Shadow(Me) Then
        .Depth = 20 '阴影宽度
        .Color = RGB(0, 0, 0) '阴影颜色
        .Transparency = 50 '阴影色深
    End If
 End With
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.BackColor = &H80FF&
Label4.BackColor = &H80FF&
Form1.Hide
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0
End If
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.BackColor = &H80FF&
Label4.BackColor = &H80FF&
End Sub

Private Sub Label4_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.BackColor = &H40C0&
Label5.BackColor = &H80FF&
End Sub

Private Sub Label5_Click()
Unload Me
End
End Sub


Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.BackColor = &H40C0&
Label4.BackColor = &H80FF&
End Sub

Private Sub PButton1_Click()
If Text1.Text = "" Then
MsgBox "最大值不能为空"
Else
If Text2.Text = "" Then
MsgBox "最小值不能为空"
Else
If Text1.Text < Text2.Text Then
MsgBox "最小值不能大于最大值"
Else
If IsNumeric(Text1.Text) And IsNumeric(Text2.Text) Then
Dim write1 As Long

Form1.max = Text1.Text
Form1.min = Text2.Text
Open "max" For Output As #11
Print #11, Form1.max
Close #11
Open "min" For Output As #11
Print #11, Form1.min
Close #11
Open "start" For Output As #11
Print #11, "1"
Close #11

Dim sla As Integer
Dim w6 As Long
For sla = Form1.min To Form1.max
w6 = WritePrivateProfileString("name", "name" & sla, "", App.Path & "\setting.ini")
Next

Form1.Show
Form2.Hide
Else
MsgBox "只能为数字"
End If
End If
End If
End If
End Sub
