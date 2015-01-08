Imports SunSoftUtility
Imports SunSoftUtility.MyInternet
Imports System.Xml

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

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        MsgBox(2 + 8 * 6 / 4 Mod 2)
    End Sub

    Public Function AddElement(ByRef oXmlDom As XmlDocument, ByRef oParent As XmlElement, ByVal xmlName As String) As XmlElement
        Dim nXmlElement As XmlElement = oXmlDom.CreateElement(xmlName)
        If oParent Is Nothing Then
            oXmlDom.AppendChild(nXmlElement)
        Else
            oParent.AppendChild(nXmlElement)
        End If
        Return nXmlElement
    End Function

    Public Function AddAttribute(ByRef oXmlDom As XmlDocument, ByRef oXmlElement As XmlElement, ByVal AttrName As String, ByVal AttrValue As String) As XmlAttribute
        Dim nAttr As XmlAttribute = oXmlDom.CreateAttribute(AttrName)
        If Not String.IsNullOrEmpty(AttrValue) Then
            nAttr.Value = AttrValue
        End If
        oXmlElement.Attributes.Append(nAttr)
        Return nAttr
    End Function

    Public Function AddText(ByRef oXmlDom As XmlDocument, ByRef oXmlElement As XmlElement, ByVal strText As String) As XmlNode
        Dim nTextNode As XmlNode = oXmlDom.CreateTextNode(strText)
        oXmlElement.AppendChild(nTextNode)
        Return nTextNode
    End Function

    Public Shared Function AddElementWithData(ByRef oXmlDom As XmlDocument, ByRef oParent As XmlElement, ByVal sXmlName As String, ByVal sData As String) As XmlElement
        Dim nXmlElement As XmlElement = oXmlDom.CreateElement(sXmlName)
        nXmlElement.InnerText = sData
        If Not oParent Is Nothing Then
            oParent.AppendChild(nXmlElement)
        Else
            oXmlDom.AppendChild(nXmlElement)
        End If
        Return nXmlElement
    End Function


    Public Function AddNode(ByRef oXmlDom As XmlDocument, ByVal NodeName As String) As XmlNode
        Dim nNode As XmlNode = oXmlDom.CreateNode(XmlNodeType.Element, NodeName, "")
        oXmlDom.AppendChild(nNode)
        Return nNode
    End Function

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim l As New XmlDocument
        Dim oParentElement As XmlElement = AddElement(l, Nothing, "data")
        Dim oRowElement As XmlElement = AddElement(l, oParentElement, "row")
        AddAttribute(l, oRowElement, "name", "sunrui")
        AddElementWithData(l, oRowElement, "jan", "0.001")
        Dim s As String = l.OuterXml
        l.Save("C:\test.xml")
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Using tick As New MyTimer
            tick.InitTimer()
            tick.StartTimer()
            Dim i As New MyXml
            Dim o As New XmlDocument
            Threading.Thread.Sleep(50)
            i.CreateByPath(o, "/love/say/me")
            Dim k As String = o.OuterXml
            MsgBox(tick.StopAndGet)
        End Using
    End Sub
End Class
