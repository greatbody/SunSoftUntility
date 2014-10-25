Imports System.Net
Imports System.Net.Cache
Imports System.IO
Imports System.IO.Compression

Public Class MyWebRequest
    Private _cookieContainer As CookieContainer
    Private _cookieCollection As CookieCollection

    Public Sub ClearCookie()
        _cookieCollection = Nothing
        _cookieContainer = Nothing
    End Sub

    Public Function GetPage(ByVal Url As String) As String
        Dim httpReq As HttpWebRequest
        Dim httpResp As HttpWebResponse
        Dim UriLoc As New Uri(Url)
        'Get Ready
        httpReq = HttpWebRequest.Create(UriLoc)
        httpReq.CookieContainer = _cookieContainer
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
        _cookieCollection = httpResp.Cookies
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
End Class
