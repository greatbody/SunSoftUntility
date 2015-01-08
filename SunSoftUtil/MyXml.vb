Imports System.Xml
Imports System.Text
Imports System.IO
Public Class MyXml
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="strXsltPath"></param>
    ''' <param name="XmlCode"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetXml(ByVal strXsltPath As String, ByVal XmlCode As String) As String
        Dim xmlDoc As New XmlDocument
        Dim xslt As New Xml.Xsl.XslCompiledTransform
        Dim writer As XmlWriter
        Dim output As New StringBuilder
        Dim ms As New MemoryStream
        writer = XmlWriter.Create(output)
        xmlDoc.LoadXml(XmlCode)
        xslt.Load(strXsltPath)
        xslt.Transform(xmlDoc, Nothing, writer, Nothing)
        Return(output.ToString)
    End Function
    ''' <summary>
    ''' 为XmlDom添加元素
    ''' </summary>
    ''' <param name="oXmlDom">XmlDocument实例对象</param>
    ''' <param name="oParent">父节点对象，可以填Nothing（C#填null）</param>
    ''' <param name="xmlName">Xml元素名</param>
    ''' <returns>Xml元素对象</returns>
    ''' <remarks></remarks>
    Public Function AddElement(ByRef oXmlDom As XmlDocument, ByRef oParent As XmlElement, ByVal xmlName As String) As XmlElement
        Dim nXmlElement As XmlElement = oXmlDom.CreateElement(xmlName)
        If oParent Is Nothing Then
            oXmlDom.AppendChild(nXmlElement)
        Else
            oParent.AppendChild(nXmlElement)
        End If
        Return nXmlElement
    End Function
    ''' <summary>
    ''' 增加一个无属性元素，同时设置其内容
    ''' </summary>
    ''' <param name="oXmlDom"></param>
    ''' <param name="oParent"></param>
    ''' <param name="sXmlName"></param>
    ''' <param name="sData"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AddElementWithData(ByRef oXmlDom As XmlDocument, ByRef oParent As XmlElement, ByVal sXmlName As String, ByVal sData As String) As XmlElement
        Dim nXmlElement As XmlElement = oXmlDom.CreateElement(sXmlName)
        nXmlElement.InnerText = sData
        If Not oParent Is Nothing Then
            oParent.AppendChild(nXmlElement)
        Else
            oXmlDom.AppendChild(nXmlElement)
        End If
        Return nXmlElement
    End Function
    ''' <summary>
    ''' 为Xml元素添加属性
    ''' </summary>
    ''' <param name="oXmlDom">XmlDocument实例对象</param>
    ''' <param name="oXmlElement">被添加属性的Xml元素对象</param>
    ''' <param name="AttrName">属性名</param>
    ''' <param name="AttrValue">属性值</param>
    ''' <returns>返回新的属性对象，同时属性一附加给Xml元素</returns>
    ''' <remarks></remarks>
    Public Function AddAttribute(ByRef oXmlDom As XmlDocument, ByRef oXmlElement As XmlElement, ByVal AttrName As String, ByVal AttrValue As String) As XmlAttribute
        Dim nAttr As XmlAttribute = oXmlDom.CreateAttribute(AttrName)
        If Not String.IsNullOrEmpty(AttrValue) Then
            nAttr.Value = AttrValue
        End If
        oXmlElement.Attributes.Append(nAttr)
        Return nAttr
    End Function
    ''' <summary>
    ''' 为Xml元素增加文字节点
    ''' </summary>
    ''' <param name="oXmlDom">XmlDocument实例对象</param>
    ''' <param name="oXmlElement">Xml元素</param>
    ''' <param name="strText">文本内容</param>
    ''' <returns>返回新加的文本节点</returns>
    ''' <remarks></remarks>
    Public Function AddText(ByRef oXmlDom As XmlDocument, ByRef oXmlElement As XmlElement, ByVal strText As String) As XmlNode
        Dim nTextNode As XmlNode = oXmlDom.CreateTextNode(strText)
        oXmlElement.AppendChild(nTextNode)
        Return nTextNode
    End Function

    Public Function CreateByPath(ByRef oXmlDom As XmlDocument, ByVal XmlPath As String) As XmlNode
        Dim PathElement() As String = Split(XmlPath, "/")
        Dim CompPath As New StringBuilder
        Dim EachElement As XmlNode
        Dim CurrentElement As XmlNode
        CurrentElement = oXmlDom.SelectSingleNode("/")
        If PathElement.Length > 0 Then
            For i As Integer = 0 To PathElement.Length - 1
                CompPath.Append("/")
                CompPath.Append(PathElement(i))
                EachElement = oXmlDom.SelectSingleNode(CompPath.ToString())
                If EachElement Is Nothing Then
                    EachElement = oXmlDom.CreateElement(PathElement(i))
                    CurrentElement.AppendChild(EachElement)
                    CurrentElement = EachElement
                Else
                    CurrentElement = EachElement
                End If
            Next
            Return CurrentElement
        Else
            Dim element As XmlElement = oXmlDom.CreateElement(XmlPath)
            Return element
        End If
    End Function
End Class
