Attribute VB_Name = "�����ȡ"
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public ret As RECT

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Type POINTAPI
    X As Long
    Y As Long
End Type
Public Point As POINTAPI

Public Sub GetFormSide()
    GetWindowRect Form5.hwnd, ret '��ȡForm1�ı�����(ret.Left, ret.Top, ret.Right, ret.Bottom)
End Sub

Public Sub GetMouseSide()
    GetCursorPos Point '��ȡ���XY����(Point.X, Point.Y)
End Sub



