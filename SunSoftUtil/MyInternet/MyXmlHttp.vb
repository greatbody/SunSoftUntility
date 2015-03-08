Imports System.Net
Imports System.IO
Namespace MyInternet
    Public Class MyXmlHttp
        Dim rtCharset As String '网页属性返回的charset类型
        Dim wsCharset As String '网页代码声称的charset类型“charset="utf-8"”等等类似，ws=Web Says
        Dim httpReq As System.Net.HttpWebRequest
        Dim httpResp As System.Net.HttpWebResponse
        Dim httpURL As System.Uri
        Dim myCookie As CookieCollection
        Public WriteOnly Property UserAgent() As String
            Set(ByVal value As String)
                httpReq.UserAgent = value
            End Set
        End Property

        Public WriteOnly Property Accept() As String
            Set(ByVal value As String)
                httpReq.Accept = value
            End Set
        End Property

        Public WriteOnly Property ContentType() As String
            Set(ByVal value As String)
                httpReq.ContentType = value
            End Set
        End Property

        Public WriteOnly Property Referer() As String
            Set(ByVal value As String)
                httpReq.Referer = value
            End Set
        End Property

        Public WriteOnly Property KeepAlive() As Boolean
            Set(ByVal value As Boolean)
                httpReq.KeepAlive = value
            End Set
        End Property

        Public Property Method() As String
            Set(ByVal value As String)
                httpReq.Method = value
            End Set
            Get
                Return httpReq.Method
            End Get
        End Property
        Public Sub New(ByVal url As String)
            httpURL = New System.Uri(url)
            httpReq = CType(WebRequest.Create(httpURL), HttpWebRequest)
        End Sub

        Public Sub Reset(ByVal url As String)
            httpURL = New System.Uri(url)
            httpReq = CType(WebRequest.Create(httpURL), HttpWebRequest)
        End Sub

        ''' <summary>
        ''' 支持自定义配置的参数的Post
        ''' </summary>
        ''' <param name="data">Post的数据</param>
        ''' <returns>服务器返回的网页代码</returns>
        ''' <remarks></remarks>
        Public Function PostData(ByVal data As String) As String  '获取源码
            Dim ContentStr As String = ""
            Try
                httpReq.Method = "POST"
                'httpReq.Timeout = 100
                If String.IsNullOrEmpty(data) = False Then
                    Dim bs() As Byte = System.Text.Encoding.GetEncoding("UTF-8").GetBytes(data)
                    httpReq.GetRequestStream.Write(bs, 0, bs.Length)
                End If

                If myCookie Is Nothing Then
                Else
                    httpReq.CookieContainer = New CookieContainer
                    httpReq.CookieContainer.Add(myCookie)
                End If
                httpResp = CType(httpReq.GetResponse(), HttpWebResponse)
                rtCharset = httpResp.CharacterSet '获取网页属性体现的charset
                wsCharset = httpResp.ContentEncoding
                Dim readers As StreamReader = New StreamReader(httpResp.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
                Dim ke As IO.Stream = httpResp.GetResponseStream
                ContentStr = readers.ReadToEnd
                wsCharset = RegExp.GetFirstMatch(LCase(ContentStr), "charset=[""0-9a-zA-Z]*")
                myCookie = httpResp.Cookies
                readers.Close()
                httpResp.Close()
                Return ContentStr
            Catch ex As Exception
                Return "error"
            End Try
        End Function
        ''' <summary>
        ''' 简易Post函数，内部设置了常用的参数
        ''' </summary>
        ''' <param name="data">需要提交的数据</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function SimplePost(ByVal data As String) As String
            Dim ContentStr As String = ""
            Try
                httpReq.Method = "POST"
                'httpReq.Timeout = 100
                httpReq.IfModifiedSince = CDate("1900-09-21")
                httpReq.Headers.Add("Accept-Charset", "UTF-8")
                If String.IsNullOrEmpty(data) = False Then
                    Dim bs() As Byte = System.Text.Encoding.GetEncoding("UTF-8").GetBytes(data)
                    httpReq.GetRequestStream.Write(bs, 0, bs.Length)
                End If

                If myCookie Is Nothing Then
                Else
                    httpReq.CookieContainer = New CookieContainer
                    httpReq.CookieContainer.Add(myCookie)
                End If
                httpReq.KeepAlive = True
                httpResp = CType(httpReq.GetResponse(), HttpWebResponse)
                rtCharset = httpResp.CharacterSet '获取网页属性体现的charset
                Dim readers As StreamReader = New StreamReader(httpResp.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
                ContentStr = readers.ReadToEnd
                readers.Close()
                httpResp.Close()
                Return ContentStr
            Catch ex As Exception
                Return "error"
            End Try

        End Function

        ''' <summary>
        ''' 加头数据
        ''' </summary>
        ''' <param name="KeyName">数据名</param>
        ''' <param name="ValueData">数据类型</param>
        ''' <remarks></remarks>
        Public Sub AddHeader(ByVal KeyName As String, ByVal ValueData As String)
            Select Case LCase(KeyName)
                Case "connection"
                    If LCase(ValueData) = "keep-alive" Then
                        httpReq.KeepAlive = True
                    Else
                        httpReq.KeepAlive = False
                    End If
                Case "user-agent"
                    httpReq.UserAgent = ValueData
                Case "accept"
                    httpReq.Accept = ValueData
                Case "content-length"
                    httpReq.ContentLength = CLng(ValueData)
                Case "content-type"
                    httpReq.ContentType = ValueData
                Case "expect"
                    httpReq.Expect = ValueData
                Case "date"
                Case "host"
                Case "if-modified-since"
                    httpReq.IfModifiedSince = CDate(ValueData)
                Case "refer"
                    httpReq.Referer = ValueData
                Case "transfer-encoding"
                    httpReq.TransferEncoding = ValueData
                Case Else
                    httpReq.Headers.Add(KeyName, ValueData)
            End Select
        End Sub
    End Class
End Namespace