Imports SunSoftUtility
Imports SunSoftUtility.MyInternet

Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        SunSoftUtility.FileSystem.LocateObject("E:\GitCode\SunSoftLibrary\TryLibrary\TryLibrary.sln")
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        SunSoftUtility.FileSystem.ExcuteUrl("E:\快盘\电脑编程\【数据库】图文优秀文章数据库\文章数据库1.0.1.9.exe")
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim i As Int16 = (&H201)
        MsgBox(SunSoftUtility.MyByte.DataToBinString(i))
    End Sub

    Private Sub btnSubmitCookie_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSubmitCookie.Click
        Dim res As String
        Dim k As New MyXmlHttp("http://localhost/xiehui/index.asp?action=login")
        res = k.SimplePost("Sunusername=admin&Sunpassword=sunwinkeyi&B1=%CC%E1%BD%BB")
        Debug.Print(res)
        k.Reset("http://localhost/xiehui/member_edit.asp")
        res = k.SimplePost("")
        Debug.Print(res)
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim res As String
        Dim k As New MyXmlHttp("http://www.baidu.com")
        res = k.PostData("")
        Debug.Print(res)
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim r As New MyWebRequest
        Dim ts As String = r.GetPage("http://www.vote8.cn/v/polyphoto")
        TextBox1.AppendText(ts)
    End Sub
End Class
