Imports System.Drawing
Public Class SimpleDraw
    Public Sub DrawPoint(ByVal hwnd As Integer, ByVal x As Integer, ByVal y As Integer, ByVal radium As Integer)
        DrawPoint(hwnd, x, y, radium, Color.Black)
    End Sub
    Public Sub DrawPoint(ByVal hwnd As Integer, ByVal x As Integer, ByVal y As Integer, ByVal radium As Integer,
                         ByVal color As Color)
        Dim bm = Graphics.FromHwnd(hwnd)
        Dim brush = New SolidBrush(color)
        bm.FillEllipse(brush, x - radium, y - radium, radium, radium)
        bm = Nothing
    End Sub
End Class
