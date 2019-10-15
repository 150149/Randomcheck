Attribute VB_Name = "坐标获取"
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
    GetWindowRect Form5.hwnd, ret '获取Form1四边坐标(ret.Left, ret.Top, ret.Right, ret.Bottom)
End Sub

Public Sub GetMouseSide()
    GetCursorPos Point '获取鼠标XY坐标(Point.X, Point.Y)
End Sub



