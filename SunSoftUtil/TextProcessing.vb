Imports System.Text
Public Class TextProcessing
    ''' <summary>
    ''' 以默认编码格式读取文本到字符串的代码
    ''' </summary>
    ''' <param name="FilePath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ReadFile(ByVal FilePath As String) As String
        Dim StringRead As String = ""
        Try
            If IO.File.Exists(FilePath) = True Then
                Dim myFileReader As New IO.StreamReader(FilePath, System.Text.Encoding.Default)
                StringRead = myFileReader.ReadToEnd
                Return StringRead
            End If
            Return StringRead
        Catch ex As Exception
            Return StringRead
        End Try
    End Function
    ''' <summary>
    ''' 以指定编码格式读取文件文本到字符串
    ''' </summary>
    ''' <param name="FilePath">文件路径</param>
    ''' <param name="encoding">文本编码类型</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ReadFile(ByVal FilePath As String, ByVal encoding As System.Text.Encoding) As String
        Dim StringRead As String = ""
        Try
            If IO.File.Exists(FilePath) = True Then
                Dim myFileReader As New IO.StreamReader(FilePath, encoding)
                StringRead = myFileReader.ReadToEnd
                myFileReader.Dispose()
                Return StringRead
            End If
            Return StringRead
        Catch ex As Exception
            Return StringRead
        End Try
    End Function
    ''' <summary>
    ''' 以默认编码写入字符串到文本文件，覆盖原内容
    ''' </summary>
    ''' <param name="FilePath">文件路径</param>
    ''' <param name="Content">文本内容</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function WriteToFile(ByVal FilePath As String, ByVal Content As String) As Boolean
        Try
            Dim myWriter As New IO.StreamWriter(FilePath, False, System.Text.Encoding.Default)
            myWriter.Write(Content)
            myWriter.Dispose()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    ''' <summary>
    ''' 以指定编码写入字符串到文本文件，覆盖原内容
    ''' </summary>
    ''' <param name="FilePath">文件路径</param>
    ''' <param name="Content">文本内容</param>
    ''' <param name="encoding">文本编码类型</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function WriteToFile(ByVal FilePath As String, ByVal Content As String, ByVal encoding As System.Text.Encoding) As Boolean
        Try
            Dim myWriter As New IO.StreamWriter(FilePath, False, encoding)
            myWriter.Write(Content)
            myWriter.Dispose()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    ''' <summary>
    ''' 以默认编码追加文本到文件末尾，不加行结束符
    ''' </summary>
    ''' <param name="FilePath">文件路径</param>
    ''' <param name="Content">文本内容</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function AddStringToFile(ByVal FilePath As String, ByVal Content As String) As Boolean
        Try
            Dim myWriter As New IO.StreamWriter(FilePath, True, System.Text.Encoding.Default)
            myWriter.Write(Content)
            myWriter.Dispose()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    ''' <summary>
    ''' 以指定编码追加文本到文件末尾，不加行结束符
    ''' </summary>
    ''' <param name="FilePath">文件路径</param>
    ''' <param name="Content">文本内容</param>
    ''' <param name="encoding">文本编码类型</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function AddStringToFile(ByVal FilePath As String, ByVal Content As String, ByVal encoding As System.Text.Encoding) As Boolean
        Try
            Dim myWriter As New IO.StreamWriter(FilePath, True, encoding)
            myWriter.Write(Content)
            myWriter.Dispose()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    ''' <summary>
    ''' 追加一行字符串到文件末尾
    ''' </summary>
    ''' <param name="FilePath">文件路径</param>
    ''' <param name="Content">文本内容</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertString(ByVal FilePath As String, ByVal Content As String) As Boolean
        Try
            Dim myWriter As New IO.StreamWriter(FilePath, True)
            myWriter.WriteLine(Content)
            myWriter.Dispose()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    ''' <summary>
    ''' 追加一行字符串以指定编码格式写入到文件末尾
    ''' </summary>
    ''' <param name="FilePath">文件路径</param>
    ''' <param name="Content">文本内容</param>
    ''' <param name="encoding">文本编码类型</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertString(ByVal FilePath As String, ByVal Content As String, ByVal encoding As System.Text.Encoding) As Boolean
        Try
            Dim myWriter As New IO.StreamWriter(FilePath, True, encoding)
            myWriter.WriteLine(Content)
            myWriter.Dispose()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Shared Function AutoReadFile(ByVal FilePath As String) As String
        Try
            If IO.File.Exists(FilePath) = True Then
                Dim newStr As New IO.FileStream(FilePath, IO.FileMode.Open, IO.FileAccess.Read)
                Dim tmpByte() As Byte
                ReDim tmpByte(newStr.Length - 1)
                newStr.Read(tmpByte, 0, newStr.Length)
                Return ByteToText(tmpByte)
            Else
                Return ""
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Shared Function ReturnEncoding(ByVal tB() As Byte) As System.Text.Encoding
        Dim tB1 As Byte, tB2 As Byte, tB3 As Byte, tB4 As Byte
        If tB.Length < 2 Then Return Nothing
        tB1 = tB(0)
        tB2 = tB(1)
        If tB.Length >= 3 Then tB3 = tB(2)
        If tB.Length >= 4 Then tB4 = tB(3)
        If (tB1 = &HFE AndAlso tB2 = &HFF) Then Return System.Text.Encoding.BigEndianUnicode
        If (tB1 = &HFF AndAlso tB2 = &HFE AndAlso tB3 <> &HFF) Then Return System.Text.Encoding.Unicode
        If (tB1 = &HEF AndAlso tB2 = &HBB AndAlso tB3 = &HBF) Then Return System.Text.Encoding.UTF8
        Return System.Text.Encoding.Default
    End Function

    Private Shared Function ByteToText(ByVal mByte() As Byte) As String
        Dim tE As System.Text.Encoding = ReturnEncoding(mByte)
        Return tE.GetString(mByte)
    End Function

    Private Shared Function ReadFileByte(ByVal FilePath As String) As Byte()
        Dim tmpByte() As Byte
        ReDim tmpByte(0)
        If IO.File.Exists(FilePath) = True Then
            Dim newStr As New IO.FileStream(FilePath, IO.FileMode.Open, IO.FileAccess.Read)
            ReDim tmpByte(newStr.Length - 1)
            newStr.Read(tmpByte, 0, newStr.Length)
            Return tmpByte
        End If
        Return tmpByte
    End Function
End Class
