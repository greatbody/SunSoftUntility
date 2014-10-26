Imports System.Net
Imports System.Net.Cache
Imports System.IO
Imports System.IO.Compression

Public Class MyWebRequest
    Public CookieContainer As New CookieContainer
    Public CookieCollection As New CookieCollection

    Public Sub ClearCookie()
        CookieCollection = New CookieCollection()
        CookieContainer = New CookieContainer()
    End Sub

    Public Function GetPage(ByVal Url As String) As String
        Dim httpReq As HttpWebRequest
        Dim httpResp As HttpWebResponse
        Dim UriLoc As New Uri(Url)
        'Get Ready
        httpReq = HttpWebRequest.Create(UriLoc)

        httpReq.CookieContainer = New CookieContainer()
        httpReq.CookieContainer.Add(CookieCollection)

        httpReq.Method = "GET"
        httpReq.Headers.Add("Cache-Control", "max-age=0") ': 
        httpReq.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
        httpReq.Headers.Add("Accept-Encoding", "gzip,deflate")
        httpReq.Headers.Add("Accept-Language", "zh-CN,zh;q=0.8") 'Accept-Language: zh-CN,zh;q=0.8
        httpReq.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/34.0.1847.137 Safari/537.36 LBBROWSER"
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
    ''' <param name="WebResponse">返回的响应流</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetResponseBytes(ByRef WebResponse As HttpWebResponse) As Byte()
        Dim BinStream As Stream
        Dim RespBin() As Byte
        Dim TmpInt As Integer

        If Not WebResponse Is Nothing Then
            Select Case WebResponse.ContentEncoding.ToLower()
                Case "gzip"
                    BinStream = New GZipStream(WebResponse.GetResponseStream(), CompressionMode.Decompress)
                Case "deflate"
                    BinStream = New DeflateStream(WebResponse.GetResponseStream(), CompressionMode.Decompress)
                Case Else
                    BinStream = WebResponse.GetResponseStream()
            End Select
            '定义一个内存流
            Dim Mems As New MemoryStream
            '开始循环读取内存流
            TmpInt = BinStream.ReadByte()
            '如果读取的是一个正整数，则表示为数据
            While TmpInt > 0
                '将字节数据（传入的int自动转为byte）写入内存流
                Mems.WriteByte(TmpInt)
                '获取下一个字节（也有可能是结束标志）
                TmpInt = BinStream.ReadByte()
            End While
            '关闭二进制流
            BinStream.Close()
            '重置字节数组大小
            ReDim RespBin(Mems.Length - 1)
            '将内存流的指针拨回到0
            Mems.Position = 0
            '将内存流中的数据读入到字节数组
            Mems.Read(RespBin, 0, RespBin.Length - 1)
            '关闭内存流
            Mems.Close()
            Return RespBin
        End If
        Return Nothing
    End Function
    ''' <summary>
    ''' 获取服务器响应的文本
    ''' </summary>
    ''' <param name="WebResponse">返回的响应流</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetResponseText(ByRef WebResponse As HttpWebResponse) As String
        Dim TextStream As StreamReader
        Dim CacheText As String
        Select Case WebResponse.ContentEncoding.ToLower()
            Case "gzip"
                Dim RespStream As IO.Compression.GZipStream = New GZipStream(WebResponse.GetResponseStream(), CompressionMode.Decompress)
                TextStream = New StreamReader(RespStream, Text.Encoding.GetEncoding(WebResponse.CharacterSet))
                CacheText = TextStream.ReadToEnd()
            Case "deflate"
                Dim RespStream As IO.Compression.DeflateStream = New DeflateStream(WebResponse.GetResponseStream(), CompressionMode.Decompress)
                TextStream = New StreamReader(RespStream, Text.Encoding.GetEncoding(WebResponse.CharacterSet))
                CacheText = TextStream.ReadToEnd()
            Case Else
                TextStream = New StreamReader(WebResponse.GetResponseStream(), Text.Encoding.GetEncoding(WebResponse.CharacterSet))
                CacheText = TextStream.ReadToEnd()
        End Select
        Return CacheText
    End Function

    Public Function PostData(ByVal Url As String, ByVal Data As String) As String
        Dim httpReq As HttpWebRequest
        Dim httpResp As HttpWebResponse
        Dim UriLoc As New Uri(Url)
        'Get Ready
        httpReq = HttpWebRequest.Create(UriLoc)
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

        httpReq.ContentType = "application/x-www-form-urlencoded"
        httpReq.ContentLength = Text.Encoding.UTF8.GetByteCount(Data)

        httpReq.KeepAlive = True
        '
        Dim bs() As Byte = System.Text.Encoding.GetEncoding("UTF-8").GetBytes(Data)
        httpReq.GetRequestStream.Write(bs, 0, bs.Length)
        'Have Response
        httpResp = CType(httpReq.GetResponse(), HttpWebResponse)
        'cache Cookies
        CookieCollection.Add(httpResp.Cookies)
        '返回网页源代码
        Return GetResponseText(httpResp)
    End Function

    Public Sub DownloadToFile(ByVal Url As String, ByVal FilePath As String)
        'define data cache
        Dim webdata() As Byte
        'define stream
        Dim WebStream As Stream
        'define request and response and uri
        Dim httpReq As HttpWebRequest
        Dim httpResp As Net.HttpWebResponse
        Dim Uri As New Uri(Url)
        'create request
        httpReq = WebRequest.Create(Uri)
        'request setting
        httpReq.Method = "GET"
        httpReq.KeepAlive = True
        httpReq.Referer = Url
        'cookie
        httpReq.CookieContainer = New CookieContainer()
        httpReq.CookieContainer.Add(CookieCollection)
        'send request and get the response
        Try
            'send and get response
            httpResp = CType(httpReq.GetResponse(), HttpWebResponse)
            'ready to save to file
            ReDim webdata(httpResp.ContentLength - 1)
            WebStream = httpResp.GetResponseStream()
            WebStream.Read(webdata, 0, webdata.Length)
            WebStream.Close()
            CookieCollection.Add(httpResp.Cookies)
            If webdata.Length > 0 Then
                If IO.File.Exists(FilePath) Then
                    IO.File.Delete(FilePath)
                End If
                Dim file As New IO.FileStream(FilePath, FileMode.Create, FileAccess.Write)
                file.Write(webdata, 0, webdata.Length)
                file.Close()
            End If
            httpResp.Close()
        Catch ex As Exception

        End Try
    End Sub
End Class
