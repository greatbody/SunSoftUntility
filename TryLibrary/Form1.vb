Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        SunSoftUtility.FileSystem.LocateObject("E:\GitCode\SunSoftLibrary\TryLibrary\TryLibrary.sln")
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        SunSoftUtility.FileSystem.ExcuteURL("E:\快盘\电脑编程\【数据库】图文优秀文章数据库\文章数据库1.0.1.9.exe")
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim i As Int16 = (&H201)
        MsgBox(SunSoftUtility.MyByte.DataToBinString(i))
    End Sub
End Class
