Imports System.Windows.Forms

Public Class FileSystem
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
    Sub New()

    End Sub
    ''' <summary>
    ''' 打开资源管理器，选中制定文件或目录
    ''' </summary>
    ''' <param name="Path">文件或目录路径</param>
    ''' <remarks></remarks>
    Public Shared Sub LocateObject(ByVal Path As String) 'ok at 12-08-01[RE]【移至成功】  '【完成：2012-09-06】
        '功能：在资源管理其中选中一个文件或文件夹
        Shell("explorer.exe /n ,/select ," & Replace(Path, "\\", "\"), vbNormalFocus)
    End Sub
    ''' <summary>
    ''' 判断一个文件是否为不含扩展名的文件
    ''' 因为有的路径是目录，有的路径是不含扩展名的文件
    ''' </summary>
    ''' <param name="FilePath">文件路径</param>
    ''' <returns>逻辑值，表示路径所指是否为不含扩展名的文件</returns>
    ''' <remarks></remarks>
    Public Shared Function IsNoExFile(ByVal FilePath As String) As Boolean  'ok at 12-05-13[RE]  '【完成：2012-09-06】
        '功能：判断路径到底是不是文件，因为可能有的文件是没有扩展名的
        If GetAttr(FilePath) <> vbDirectory And InStr(1, Mid(FilePath, InStrRev(FilePath, "\") + 1), ".") = 0 Then
            '此路径不为目录，且没有扩展名
            Return True
        Else
            Return False
        End If
    End Function
    ''' <summary>
    ''' 运行指定路径，不论是目录还是文件、程序
    ''' </summary>
    ''' <param name="UrlPath"></param>
    ''' <remarks></remarks>
    Public Shared Sub ExcuteUrl(ByVal UrlPath As String) 'ok at 12-09-06[RE]  '【完成：2012-09-06】
        '程序功能：打开文件，且设置文件的父文件夹为其当前运行文件夹路径
        '2012-09-06 改造成为过程
        Dim lngReturn As Long
        If IsNoExFile(UrlPath) = False Then
            lngReturn = ShellExecute(0, "open", UrlPath, "", Mid(UrlPath, 1, InStrRev(UrlPath, "\") - 1), 1)
        Else
            lngReturn = Shell("rundll32.exe   shell32.dll   OpenAs_RunDLL   " & UrlPath)
        End If
        If lngReturn = 31 Then
            lngReturn = Shell("rundll32.exe   shell32.dll   OpenAs_RunDLL   " & UrlPath)
        End If
    End Sub
    Public Shared Function SelectFolder(ByVal Describe As String, Optional ByVal ShowNewFolder As Boolean = True) As String
        Using nOpen As New System.Windows.Forms.FolderBrowserDialog()
            nOpen.Description = Describe
            nOpen.ShowNewFolderButton = ShowNewFolder
            nOpen.ShowDialog()
            Return nOpen.SelectedPath
        End Using
    End Function
    Public Shared Function SelectFiles(ByVal Describe As String) As String
        Dim nOpen As New System.Windows.Forms.SaveFileDialog
        '不检查文件是否存在
        nOpen.CheckFileExists = False
        '检查路径是否存在，并提示
        nOpen.CheckPathExists = True
        '设置文件选择对话框的标题
        nOpen.Title = Describe
        '如果指定不存在的文件，系统提示是否创建
        nOpen.CreatePrompt = False
        nOpen.Filter = "All Files(*.*)|*.*"
        Dim diaResult As DialogResult = nOpen.ShowDialog()
        If diaResult = DialogResult.OK Then
            Return nOpen.FileName
        Else
            Return Nothing
        End If
    End Function
End Class
