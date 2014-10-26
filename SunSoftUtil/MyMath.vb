Public Class MyMath
    Public Shared Function GetRndInt() As Integer
        Randomize()
        Return Int(Rnd(1) * 240)
    End Function
End Class
