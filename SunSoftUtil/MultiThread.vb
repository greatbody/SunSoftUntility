Public Class MultiThread
    ''' <summary>
    ''' 检测当前程序进程是否是程序的第二次运行
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' 
    Public Function preInstanceCheck() As Boolean
        'This uses the Process class to check for the name of the current applications process and ‘see whether or not there a
        're more than 1x instance loaded.‘The end result of this code is similar to Visual Basic 6.0′s App.Previnstance feature.
        Dim appName As String = Process.GetCurrentProcess.ProcessName   '获取当前进程的进程名
        Dim sameProcessTotal As Integer = Process.GetProcessesByName(appName).Length    '获取进程中同名的进程
        If sameProcessTotal > 1 Then
            'MessageBox.Show("A previous instance of this application is already open!", " App.PreInstance Detected!", _
            'MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return True
        End If
        appName = Nothing
        sameProcessTotal = Nothing
        Return False
    End Function

End Class
