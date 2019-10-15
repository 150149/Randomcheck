Attribute VB_Name = "移动窗体"
Public bFORM As Boolean '窗体是否隐藏
Public bMOUSE As Boolean '鼠标是否在窗体内

Public Sub Variable() '变量赋值
    bFORM = False
    bMOUSE = True
End Sub

Public Sub MouseState() '检测鼠标是否在窗体内
    If Point.X < ret.Left Or Point.X > ret.Right Or Point.Y < ret.Top Or Point.Y > ret.Bottom Then
        bMOUSE = False
    Else
        bMOUSE = True
    End If
End Sub

Public Sub MoveUpForm() '当窗体贴近屏幕顶部上移窗体
    If ret.Top < 10 And bFORM = False And bMOUSE = False Then
        Do
            If Form5.Top < -Form5.Height + 100 Then
                Exit Do
            Else
                Form5.Top = Form5.Top - 50
            End If
        Loop
        bFORM = True
    End If
End Sub

Public Sub MoveDownForm() '当窗体隐藏且鼠标靠近时窗体显现
    If bFORM = True And bMOUSE = True Then

            Form5.Top = Form5.Top + 3000

        bFORM = False
    End If
End Sub
