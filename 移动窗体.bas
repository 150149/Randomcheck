Attribute VB_Name = "�ƶ�����"
Public bFORM As Boolean '�����Ƿ�����
Public bMOUSE As Boolean '����Ƿ��ڴ�����

Public Sub Variable() '������ֵ
    bFORM = False
    bMOUSE = True
End Sub

Public Sub MouseState() '�������Ƿ��ڴ�����
    If Point.X < ret.Left Or Point.X > ret.Right Or Point.Y < ret.Top Or Point.Y > ret.Bottom Then
        bMOUSE = False
    Else
        bMOUSE = True
    End If
End Sub

Public Sub MoveUpForm() '������������Ļ�������ƴ���
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

Public Sub MoveDownForm() '��������������꿿��ʱ��������
    If bFORM = True And bMOUSE = True Then

            Form5.Top = Form5.Top + 3000

        bFORM = False
    End If
End Sub
