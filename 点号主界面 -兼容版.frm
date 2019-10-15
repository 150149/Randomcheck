VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "随机点号器 by150149"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   45
   ClientWidth     =   13215
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "点号主界面 -兼容版.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   13215
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Caption         =   "点号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8520
      TabIndex        =   10
      Top             =   5280
      Width           =   3375
   End
   Begin VB.Timer Timer1 
      Left            =   9480
      Top             =   480
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Caption         =   "点号模式"
      Height          =   975
      Left            =   5880
      TabIndex        =   5
      Top             =   5160
      Width           =   2415
      Begin VB.OptionButton Option2 
         BackColor       =   &H8000000B&
         Caption         =   "多次点号(手动停止)"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000B&
         Caption         =   "单次点号(自动停止）"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000B&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000080FF&
      X1              =   0
      X2              =   0
      Y1              =   360
      Y2              =   6240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000080FF&
      X1              =   0
      X2              =   13200
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      X1              =   13200
      X2              =   13200
      Y1              =   360
      Y2              =   6240
   End
   Begin VB.Label Label7 
      BackColor       =   &H000080FF&
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   12720
      TabIndex        =   8
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   150
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   3960
      TabIndex        =   4
      Top             =   960
      Width           =   9255
   End
   Begin VB.Label Label4 
      BackColor       =   &H000080FF&
      Caption         =   "随机点号--兼容模式"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   13215
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000B&
      Caption         =   "150149制作者  QQ：1802796278                 微信： w150149"
      Height          =   375
      Left            =   9960
      TabIndex        =   2
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000B&
      Caption         =   "点号器1.52"
      Height          =   255
      Left            =   12000
      TabIndex        =   1
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   200.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3525
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3690
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Num As Long
Public start As String
Public whitelist1 As String
Public whitelist2 As String
Public whitelist3 As String
Public whitelist4 As String
Public whitelist5 As String
Public w1 As Integer
Public w2 As Integer
Public w3 As Integer
Public w4 As Integer
Public w5 As Integer
Public shengyin As String
Public max, min As String
Public n1, n2, n3, n4, n5, n6, n7, n8, n9, n10, n11, n12, n13, n14, n15, n16, n17, n18, n19, n20, n21, n22, n23, n24, n25, n26, n27, n28, n29, n30, n31, n32, n33, n34, n35, n36, n37, n38, n39, n40, n41, n42, n43, n44, n45, n46, n47, n48, n49, n50, n51, n52, n53, n54, n55, n56, n57, n58, n59, n60 As String
Private oShadow As New aShadow
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
    
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
Public moshi As String

Private Sub Command1_Click()
If moshi = "1" Then
Debug.Print "单次点号"

Randomize
Dim m, n, a As Integer
m = CInt(max)
n = CInt(min)
If IsNumeric(whitelist1) Then
w1 = CInt(whitelist1)
Else
End If

If IsNumeric(whitelist2) Then
w2 = CInt(whitelist2)
Else
End If

If IsNumeric(whitelist3) Then
w3 = CInt(whitelist3)
Else
End If

If IsNumeric(whitelist4) Then
w4 = CInt(whitelist4)
Else
End If

If IsNumeric(whitelist5) Then
w5 = CInt(whitelist5)
Else
End If




a = Int(Rnd * (m - n + 1)) + n ' 内层循环。
If a = w1 Then ' 如果条件成立。
         a = Int(Rnd * (m - n + 1)) + n
           ' 退出内层循环。
ElseIf a = w2 Then
a = Int(Rnd * (m - n + 1)) + n
ElseIf a = w3 Then
a = Int(Rnd * (m - n + 1)) + n
           ' 退出内层循环。
ElseIf a = w4 Then
a = Int(Rnd * (m - n + 1)) + n
ElseIf a = w5 Then
a = Int(Rnd * (m - n + 1)) + n
End If
     Label1.Caption = a
     Dim mz As String
     mz = String(10, 0)
     Dim read_OK As Long
     read_OK = GetPrivateProfileString("name", "name" & a, "", mz, 10, App.Path & "\setting.ini")
     Label5.Caption = mz
     

     
ElseIf moshi = "2" And Command1.Caption = "点号(多次抽号)" Then
Timer1.Enabled = True
Timer1.Interval = 50
Debug.Print "多次点号开始"

Command1.Caption = "停止点号"
ElseIf moshi = "2" And Command1.Caption = "停止点号" Then
Timer1.Enabled = False
Command1.Caption = "点号(多次抽号)"
Debug.Print "多次点号结束"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
MsgBox "该程序已运行"
End
End If

With oShadow
    If .Shadow(Me) Then
        .Depth = 20 '阴影宽度
        .Color = RGB(0, 0, 0) '阴影颜色
        .Transparency = 50 '阴影色深
    End If
 End With
 
Dim read_OK, r2, r3, r4 As Long
    whitelist1 = String(10, 0)
    whitelist2 = String(10, 0)
    whitelist3 = String(10, 0)
    whitelist4 = String(10, 0)
    whitelist5 = String(10, 0)
    read_OK = GetPrivateProfileString("setting", "whitelist1", "0", whitelist1, 256, App.Path & "\setting.ini")
    read_OK = GetPrivateProfileString("setting", "whitelist2", "0", whitelist2, 256, App.Path & "\setting.ini")
    read_OK = GetPrivateProfileString("setting", "whitelist3", "0", whitelist3, 256, App.Path & "\setting.ini")
    read_OK = GetPrivateProfileString("setting", "whitelist4", "0", whitelist4, 256, App.Path & "\setting.ini")
    read_OK = GetPrivateProfileString("setting", "whitelist5", "0", whitelist5, 256, App.Path & "\setting.ini")
    Debug.Print ("白名单1读取为" & whitelist1)
    Debug.Print ("白名单2读取为" & whitelist2)
    Debug.Print ("白名单3读取为" & whitelist3)
    Debug.Print ("白名单4读取为" & whitelist4)
    Debug.Print ("白名单5读取为" & whitelist5)

If Dir(App.Path & "\max") = "" Then
Open "max" For Output As #11
Print #11, "60"
Close #11
End If
If Dir(App.Path & "\min") = "" Then
Open "min" For Output As #11
Print #11, "1"
Close #11
End If
If Dir(App.Path & "\start") = "" Then
Open "start" For Output As #11
Print #11, "0"
Close #11
End If

Open App.Path & "\start" For Input As #5
Line Input #5, start
Close #5
    Open App.Path & "\max" For Input As #5
Line Input #5, max
Close #5
Open App.Path & "\min" For Input As #5
Line Input #5, min
Close #5

    If start = "0" Then
    max = InputBox("请输入随机数最大值", "兼容模式-设置最大值", "60")
    min = InputBox("请输入随机数最小值", "兼容模式-设置最小值", "1")
    Dim sla As Integer
Dim w6 As Long
For sla = min To max
w6 = WritePrivateProfileString("name", "name" & sla, "", App.Path & "\setting.ini")
Next
Open "max" For Output As #11
Print #11, max
Close #11
Open "min" For Output As #11
Print #11, min
Close #11
Open "start" For Output As #11
Print #11, "1"
Close #11
    Else
Randomize
Dim m, n, a As Integer
m = CInt(max)
n = CInt(min)

If IsNumeric(whitelist1) Then
w1 = CInt(whitelist1)
Else
End If

If IsNumeric(whitelist2) Then
w2 = CInt(whitelist2)
Else
End If

If IsNumeric(whitelist3) Then
w3 = CInt(whitelist3)
Else
End If

If IsNumeric(whitelist4) Then
w4 = CInt(whitelist4)
Else
End If

If IsNumeric(whitelist5) Then
w5 = CInt(whitelist5)
Else
End If




a = Int(Rnd * (m - n + 1)) + n ' 内层循环。
If a = w1 Then ' 如果条件成立。
         a = Int(Rnd * (m - n + 1)) + n
           ' 退出内层循环。
ElseIf a = w2 Then
a = Int(Rnd * (m - n + 1)) + n
ElseIf a = w3 Then
a = Int(Rnd * (m - n + 1)) + n
           ' 退出内层循环。
ElseIf a = w4 Then
a = Int(Rnd * (m - n + 1)) + n
ElseIf a = w5 Then
a = Int(Rnd * (m - n + 1)) + n
End If

     Label1.Caption = a
     Dim mz As String
     Dim rd As Long
     mz = String(10, 0)
     rd = GetPrivateProfileString("name", "name" & a, "", mz, 10, App.Path & "\setting.ini")
     Debug.Print ("name" & a)
     Label5.Caption = mz
     
     End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0
End If
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0
End If
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0
End If
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0
End If
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0
End If
End Sub

Private Sub Label7_Click()
Unload Me
End
End Sub

Private Sub Option1_Click()
moshi = "1"
Command1.Caption = "点号(单次抽号)"
Debug.Print "模式选择为1"
Dim jlms As Long
jlms = WritePrivateProfileString("setting", "mode", "1", App.Path & "\setting.ini")
End Sub

Private Sub Option2_Click()
moshi = "2"
Command1.Caption = "点号(多次抽号)"
Debug.Print "模式选择为2"
Dim jlms As Long
jlms = WritePrivateProfileString("setting", "mode", "2", App.Path & "\setting.ini")
End Sub

Private Sub Timer1_Timer()

Randomize
Dim m, n, a As Integer
m = CInt(max)
n = CInt(min)
If IsNumeric(whitelist1) Then
w1 = CInt(whitelist1)
Else
End If

If IsNumeric(whitelist2) Then
w2 = CInt(whitelist2)
Else
End If

If IsNumeric(whitelist3) Then
w3 = CInt(whitelist3)
Else
End If

If IsNumeric(whitelist4) Then
w4 = CInt(whitelist4)
Else
End If

If IsNumeric(whitelist5) Then
w5 = CInt(whitelist5)
Else
End If




a = Int(Rnd * (m - n + 1)) + n ' 内层循环。
If a = w1 Then ' 如果条件成立。
         a = Int(Rnd * (m - n + 1)) + n
           ' 退出内层循环。
ElseIf a = w2 Then
a = Int(Rnd * (m - n + 1)) + n
ElseIf a = w3 Then
a = Int(Rnd * (m - n + 1)) + n
           ' 退出内层循环。
ElseIf a = w4 Then
a = Int(Rnd * (m - n + 1)) + n
ElseIf a = w5 Then
a = Int(Rnd * (m - n + 1)) + n
End If
     Label1.Caption = a
     Dim mz As String
     mz = String(10, 0)
     Dim read_OK As Long
     read_OK = GetPrivateProfileString("name", "name" & a, "", mz, 10, App.Path & "\setting.ini")
     Label5.Caption = mz
     
     Debug.Print ("a=" & a)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.BackColor = &H80FF&
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.BackColor = &H40C0&
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.BackColor = &H80FF&
End Sub
