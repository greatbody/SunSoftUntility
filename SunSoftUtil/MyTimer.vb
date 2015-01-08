Public Class MyTimer
    Implements IDisposable
    Private _stopWatch As New System.Diagnostics.Stopwatch
    ''' <summary>
    ''' 停止计时器，并重置计数数据
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub InitTimer()
        _stopWatch.Reset()
    End Sub
    ''' <summary>
    ''' 启动计时器
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub StartTimer()
        _stopWatch.Start()
    End Sub
    ''' <summary>
    ''' 停止计时器
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub StopTimer()
        _stopWatch.Stop()
    End Sub
    ''' <summary>
    ''' 停止计时器，并返回经历毫秒数
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function StopAndGet() As Long
        _stopWatch.Stop()
        Return _stopWatch.ElapsedMilliseconds
    End Function
    ''' <summary>
    ''' 获取经历的毫秒数
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function UnstopGet() As Long
        Return _stopWatch.ElapsedMilliseconds
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' 检测冗余的调用

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 释放托管状态(托管对象)。
                _stopWatch = Nothing
            End If

            ' TODO: 释放非托管资源(非托管对象)并重写下面的 Finalize()。
            ' TODO: 将大型字段设置为 null。
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: 仅当上面的 Dispose(ByVal disposing As Boolean)具有释放非托管资源的代码时重写 Finalize()。
    'Protected Overrides Sub Finalize()
    '    ' 不要更改此代码。请将清理代码放入上面的 Dispose(ByVal disposing As Boolean)中。
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' Visual Basic 添加此代码是为了正确实现可处置模式。
    Public Sub Dispose() Implements IDisposable.Dispose
        ' 不要更改此代码。请将清理代码放入上面的 Dispose(ByVal disposing As Boolean)中。
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
