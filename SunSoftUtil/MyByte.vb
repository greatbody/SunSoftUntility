Public Class MyByte
    Public Shared Function MyAnd(ByVal aByte As Byte, ByVal bByte As Byte)
        Return ""
        '留空，后续补上
    End Function
    ''' <summary>
    ''' 返回指定Short（Int16）类型的数据的二进制字符串
    ''' </summary>
    ''' <param name="data">Short（Int16）</param>
    ''' <returns>Short（Int16）数据的字符串形式</returns>
    ''' <remarks></remarks>
    Public Shared Function DataToBinString(ByVal data As Int16) As String
        Return DataToBinString(GetByte(data, 1)) & DataToBinString(GetByte(data, 0))
    End Function
    ''' <summary>
    ''' 返回指定Integer（Int32）类型的数据的二进制字符串
    ''' </summary>
    ''' <param name="data">Int32数据</param>
    ''' <returns>Integer（Int32）数据的字符串形式</returns>
    ''' <remarks></remarks>
    Public Shared Function DataToBinString(ByVal data As Int32) As String
        Return DataToBinString(GetByte(data, 3)) & DataToBinString(GetByte(data, 2)) & DataToBinString(GetByte(data, 1)) & DataToBinString(GetByte(data, 0))
    End Function
    ''' <summary>
    ''' 返回指定Long（Int64）类型的数据的二进制字符串
    ''' </summary>
    ''' <param name="data">Int64(Long)数据</param>
    ''' <returns>Long（Int64）数据的字符串形式</returns>
    ''' <remarks></remarks>
    Public Shared Function DataToBinString(ByVal data As Int64) As String
        Return DataToBinString(GetByte(data, 7)) & DataToBinString(GetByte(data, 6)) & _
               DataToBinString(GetByte(data, 5)) & DataToBinString(GetByte(data, 4)) & _
               DataToBinString(GetByte(data, 3)) & DataToBinString(GetByte(data, 2)) & _
               DataToBinString(GetByte(data, 1)) & DataToBinString(GetByte(data, 0))
    End Function
    ''' <summary>
    ''' 返回指定字节类型的数据的二进制字符串
    ''' </summary>
    ''' <param name="data">byte数据</param>
    ''' <returns>byte数据的字符串形式</returns>
    ''' <remarks></remarks>
    Public Shared Function DataToBinString(ByVal data As Byte) As String
        Dim rOut As String
        rOut = ""
        Do While data > 0
            rOut = data Mod 2 & rOut
            data = data \ 2
        Loop
        If rOut.Length < 8 Then
            rOut = TextProcessing.RepeatString("0", 8 - Len(rOut)) & rOut
        End If
        Return rOut
    End Function
    ''' <summary>
    ''' 将数字对象的（从低到高位）第index-1个字节提取出来
    ''' </summary>
    ''' <param name="data">数字对象</param>
    ''' <param name="index">从0开始的偏移量，0表示取第一个字节</param>
    ''' <returns>取得的字节</returns>
    ''' <remarks></remarks>
    Private Shared Function GetByte(ByVal data As Object, ByVal index As Integer) As Byte
        If data.GetType Is GetType(Byte) Then
        ElseIf data.GetType Is GetType(Int16) Then
        ElseIf data.GetType Is GetType(Int32) Then
        ElseIf data.GetType Is GetType(Int64) Then
        Else
            Return CByte(0)
        End If
        If index > Len(data) Or index < 0 Then
            Return CByte(0)
        Else
            data = data >> (index * 8)
            Return CByte((&HFF And data))
        End If
    End Function
End Class
