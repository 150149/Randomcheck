VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "兼容性警告"
   ClientHeight    =   1665
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   2745
   LinkTopic       =   "Form3"
   ScaleHeight     =   1665
   ScaleWidth      =   2745
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox Check1 
      Caption         =   "不再提醒"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "兼容模式"
      Height          =   615
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "仍然继续"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "缺少Comdlg32.ocx，可能导致未知错误的产生"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public showa As String
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
    

Private Sub Check1_Click()
showa = "1"
End Sub

Private Sub Command1_Click()
If showa = "1" Then
Dim w5 As Long
w5 = WritePrivateProfileString("setting", "NoShowAgainst", "1", App.Path & "\setting.ini")
Else
End If
Form1.Show
Form3.Hide
End Sub

Private Sub Command2_Click()
If showa = "1" Then
Dim w5 As Long
w5 = WritePrivateProfileString("setting", "NoShowAgainst", "4", App.Path & "\setting.ini")
Else
End If
Form4.Show
Form3.Hide
End Sub

Private Sub Form_Load()

If App.PrevInstance = True Then
MsgBox "该程序已运行"
End
End If

If Dir(App.Path & "\p.ocx") = "" Then
Else
Shell "regsvr32 /s p.ocx", vbHide
End If

If Dir(App.Path & "\setting.ini") = "" Then
Open "setting.ini" For Output As #11
Print #11, ""
Close #11
Dim write1, w2, w3, w4, w5, w6, w7, w8 As Long
write1 = WritePrivateProfileString("setting", "VoiceRate", "0", App.Path & "\setting.ini")
w2 = WritePrivateProfileString("setting", "Voice", "0", App.Path & "\setting.ini")
w3 = WritePrivateProfileString("setting", "VoiceForm", "0", App.Path & "\setting.ini")
'w4 = WritePrivateProfileString("setting", "whitelist4", "0", App.Path & "\setting.ini")
'w5 = WritePrivateProfileString("setting", "whitelist5", "0", App.Path & "\setting.ini")

End If

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

Dim start As String
Open App.Path & "\start" For Input As #5
Line Input #5, start
Close #5

If Dir("c:\windows\system32\Comdlg32.ocx") = "" Then
    Dim no As String
    Dim read_OK As Long
    no = String(10, 0)
    read_OK = GetPrivateProfileString("setting", "NoShowAgainst", "0", no, 256, App.Path & "\setting.ini")
        If no = 1 And start = 0 Then
            Form2.Show
            Debug.Print "no=1 start = 0"
            Form3.Hide
            End If
        If no = 4 And start = 0 Then
            Form4.Show
            Debug.Print "no=1"
            Form3.Hide
            End If
        If no = 1 And start = 1 Then
            Form1.Show
            Form3.Hide
            End If
        If no = 4 And start = 1 Then
            Form4.Show
            Form3.Hide
            End If
        If no = 0 And start = 0 Then
            Form2.Show
            Form3.Hide
            End If
        If no = 0 And start = 1 Then
            Form1.Show
            Form3.Hide
            End If
    Else
        If start = 0 Then
        Form2.Show
        Form3.Hide
        Else
        If start = 1 Then
        Form1.Show
        Form3.Hide
        End If
        End If
End If

End Sub
