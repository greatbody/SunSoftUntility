Imports System.Net
Imports System.IO
Imports System.IO.Compression

Public Class MyWebRequest
    Public CookieContainer As New CookieContainer
    Public CookieCollection As New CookieCollection

    Private _isIpFake As Boolean
    Public Property IsFakeIp() As Boolean
        Get
            Return _isIpFake
        End Get
        Set(ByVal value As Boolean)
            _isIpFake = value
        End Set
    End Property

    Public Sub ClearCookie()
        CookieCollection = New CookieCollection()
        CookieContainer = New CookieContainer()
    End Sub

    Public Function GetPage(ByVal url As String) As String
        Dim httpReq As HttpWebRequest
        Dim httpResp As HttpWebResponse
        Dim uriLoc As New Uri(url)
        'Get Ready
        httpReq = HttpWebRequest.Create(uriLoc)

        httpReq.CookieContainer = New CookieContainer()
        httpReq.CookieContainer.Add(CookieCollection)

        httpReq.Method = "GET"
        httpReq.Headers.Add("Cache-Control", "max-age=0") ': 
        httpReq.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
        httpReq.Headers.Add("Accept-Encoding", "gzip,deflate")
        httpReq.Headers.Add("Accept-Language", "zh-CN,zh;q=0.8") 'Accept-Language: zh-CN,zh;q=0.8
        httpReq.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/34.0.1847.137 Safari/537.36 LBBROWSER"

        If _isIpFake Then
            httpReq.Headers.Add("x-forward-for", MyMath.GetRndInt().ToString() & "." & MyMath.GetRndInt().ToString() & "." & MyMath.GetRndInt().ToString() & "." & MyMath.GetRndInt().ToString())
        End If

        httpReq.KeepAlive = True
        'Have Response
        httpResp = CType(httpReq.GetResponse(), HttpWebResponse)
        'cache Cookies
        CookieCollection.Add(httpResp.Cookies)
        '返回网页源代码
        Return GetResponseText(httpResp)
    End Function
    ''' <summary>
    ''' 返回响应的二进制字节数组
    ''' </summary>
    ''' <param name="webResponse">返回的响应流</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetResponseBytes(ByRef webResponse As HttpWebResponse) As Byte()
        Dim binStream As Stream
        Dim respBin() As Byte
        Dim tmpInt As Integer

        If Not webResponse Is Nothing Then
            Select Case webResponse.ContentEncoding.ToLower()
                Case "gzip"
                    binStream = New GZipStream(webResponse.GetResponseStream(), CompressionMode.Decompress)
                Case "deflate"
                    binStream = New DeflateStream(webResponse.GetResponseStream(), CompressionMode.Decompress)
                Case Else
                    binStream = webResponse.GetResponseStream()
            End Select
            '定义一个内存流
            Dim mems As New MemoryStream
            '开始循环读取内存流
            tmpInt = binStream.ReadByte()
            '如果读取的是一个正整数，则表示为数据
            While tmpInt > 0
                '将字节数据（传入的int自动转为byte）写入内存流
                mems.WriteByte(tmpInt)
                '获取下一个字节（也有可能是结束标志）
                tmpInt = binStream.ReadByte()
            End While
            '关闭二进制流
            binStream.Close()
            '重置字节数组大小
            ReDim respBin(mems.Length - 1)
            '将内存流的指针拨回到0
            mems.Position = 0
            '将内存流中的数据读入到字节数组
            mems.Read(respBin, 0, respBin.Length - 1)
            '关闭内存流
            mems.Close()
            Return respBin
        End If
        Return Nothing
    End Function
    ''' <summary>
    ''' 获取服务器响应的文本
    ''' </summary>
    ''' <param name="webResponse">返回的响应流</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetResponseText(ByRef webResponse As HttpWebResponse) As String
        Dim textStream As StreamReader
        Dim cacheText As String
        Select Case webResponse.ContentEncoding.ToLower()
            Case "gzip"
                Dim respStream As GZipStream = New GZipStream(webResponse.GetResponseStream(), CompressionMode.Decompress)
                textStream = New StreamReader(respStream, Text.Encoding.GetEncoding(webResponse.CharacterSet))
                cacheText = textStream.ReadToEnd()
            Case "deflate"
                Dim respStream As DeflateStream = New DeflateStream(webResponse.GetResponseStream(), CompressionMode.Decompress)
                textStream = New StreamReader(respStream, Text.Encoding.GetEncoding(webResponse.CharacterSet))
                cacheText = textStream.ReadToEnd()
            Case Else
                textStream = New StreamReader(webResponse.GetResponseStream(), Text.Encoding.GetEncoding(webResponse.CharacterSet))
                cacheText = textStream.ReadToEnd()
        End Select
        Return cacheText
    End Function

    Public Function PostData(ByVal url As String, ByVal data As String) As String
        Dim httpReq As HttpWebRequest
        Dim httpResp As HttpWebResponse
        Dim uriLoc As New Uri(url)
        'Get Ready
        httpReq = HttpWebRequest.Create(uriLoc)
        '
        httpReq.CookieContainer = New CookieContainer()
        httpReq.CookieContainer.Add(CookieCollection)
        '
        httpReq.Method = "POST"
        httpReq.Headers.Add("Cache-Control", "max-age=0") ': 
        httpReq.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
        httpReq.Headers.Add("Accept-Encoding", "gzip,deflate")
        httpReq.Headers.Add("Accept-Language", "zh-CN,zh;q=0.8") 'Accept-Language: zh-CN,zh;q=0.8
        httpReq.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/34.0.1847.137 Safari/537.36 LBBROWSER"
        If _isIpFake Then
            httpReq.Headers.Add("x-forward-for", MyMath.GetRndInt().ToString() & "." & MyMath.GetRndInt().ToString() & "." & MyMath.GetRndInt().ToString() & "." & MyMath.GetRndInt().ToString())
        End If
        httpReq.ContentType = "application/x-www-form-urlencoded"
        httpReq.ContentLength = Text.Encoding.UTF8.GetByteCount(data)

        httpReq.KeepAlive = True
        '
        Dim bs() As Byte = Text.Encoding.GetEncoding("UTF-8").GetBytes(data)
        httpReq.GetRequestStream.Write(bs, 0, bs.Length)
        'Have Response
        Try
            httpResp = CType(httpReq.GetResponse(), HttpWebResponse)
        Catch ex As Exception
            Return ""
        End Try

        'cache Cookies
        CookieCollection.Add(httpResp.Cookies)
        '返回网页源代码
        Return GetResponseText(httpResp)
    End Function

    Public Sub DownloadToFile(ByVal url As String, ByVal filePath As String)
        'define data cache
        Dim webdata() As Byte
        'define stream
        Dim webStream As Stream
        'define request and response and uri
        Dim httpReq As HttpWebRequest
        Dim httpResp As HttpWebResponse
        Dim uri As New Uri(url)
        'create request
        httpReq = WebRequest.Create(uri)
        'request setting
        httpReq.Method = "GET"
        httpReq.KeepAlive = True
        If _isIpFake Then
            httpReq.Headers.Add("x-forward-for", MyMath.GetRndInt().ToString() & "." & MyMath.GetRndInt().ToString() & "." & MyMath.GetRndInt().ToString() & "." & MyMath.GetRndInt().ToString())
        End If
        httpReq.Referer = url
        'cookie
        httpReq.CookieContainer = New CookieContainer()
        httpReq.CookieContainer.Add(CookieCollection)
        'send request and get the response
        Try
            'send and get response
            httpResp = CType(httpReq.GetResponse(), HttpWebResponse)
            'ready to save to file
            ReDim webdata(httpResp.ContentLength - 1)
            webStream = httpResp.GetResponseStream()
            webStream.Read(webdata, 0, webdata.Length)
            webStream.Close()
            CookieCollection.Add(httpResp.Cookies)
            If webdata.Length > 0 Then
                If IO.File.Exists(filePath) Then
                    IO.File.Delete(filePath)
                End If
                Dim file As New FileStream(filePath, FileMode.Create, FileAccess.Write)
                file.Write(webdata, 0, webdata.Length)
                file.Close()
            End If
            httpResp.Close()
        Catch ex As Exception

        End Try
    End Sub
End Class
