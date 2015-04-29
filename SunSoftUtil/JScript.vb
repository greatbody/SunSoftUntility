Imports System.Windows.Forms

Public Class JScript
    Private FunctionList As List(Of String)
    Private CurseIndex As Integer = 0

    Public Function GetNextFunction() As String
        CurseIndex += 1
        If FunctionList.Count = 0 Then
            Return ""
        End If
        If CurseIndex > FunctionList.Count Then
            CurseIndex = 1
        End If
        Return FunctionList.Item(CurseIndex)
    End Function

    Public Function GetFunctions() As String()
        Dim OutFunctions() As String
        If FunctionList.Count = 0 Then
            ReDim OutFunctions(0)
            Return OutFunctions
        Else
            Return FunctionList.ToArray()
        End If
    End Function

    Public Sub New()
        FunctionList = New List(Of String)
    End Sub
    Public Sub New(ByVal JScriptContent As String)
        FunctionList = New List(Of String)
        '获取函数列表
        Dim sFunction(0) As String
        Dim indexR As Integer = 0
        GetTopicInfo(JScriptContent, "function ", "(", sFunction)
        If sFunction.Length = 0 Or (sFunction.Length = 1 And sFunction(0) = "") Then
            '没有内容，或者空内容
            'do nothing
            Exit Sub
        End If
        '
        For indexR = 0 To sFunction.Length - 1
            FunctionList.Add(sFunction(indexR))
        Next
    End Sub
    Public Sub LoadJScript(ByVal JScriptContent As String)
        FunctionList = New List(Of String)
        '获取函数列表
        Dim sFunction(0) As String
        Dim indexR As Integer = 0
        GetTopicInfo(JScriptContent, "function ", "(", sFunction)
        If sFunction.Length = 0 Or (sFunction.Length = 1 And sFunction(0) = "") Then
            '没有内容，或者空内容
            'do nothing
            Exit Sub
        End If
        '
        For indexR = 0 To sFunction.Length - 1
            FunctionList.Add(sFunction(indexR))
        Next
    End Sub

    Public Sub FillListBox(ByRef RefListBox As ListBox)
        RefListBox.Items.Clear()
        If FunctionList.Count = 0 Then
            Exit Sub
        End If
        For indexs As Integer = 0 To FunctionList.Count - 1
            RefListBox.Items.Add(FunctionList.Item(indexs))
        Next
    End Sub

    Public Sub LoadFile(ByVal FilePath As String)
        Dim sFunction(0) As String
        Dim indexR As Integer = 0
        Dim JavaScript As String = SunSoftUtility.TextProcessing.AutoReadFile(FilePath)
        GetTopicInfo(JavaScript, "function ", "(", sFunction)
        If sFunction.Length = 0 Or (sFunction.Length = 1 And sFunction(0) = "") Then
            '没有内容，或者空内容
            'do nothing
            Exit Sub
        End If
        '
        For indexR = 0 To sFunction.Length - 1
            FunctionList.Add(sFunction(indexR))
        Next
    End Sub

    ''' <summary>
    ''' 将分割出的信息返回数组[完成，测试通过]
    ''' </summary>
    ''' <param name="code">文本</param>
    ''' <param name="divBegin">开始段</param>
    ''' <param name="divEnd">结束段</param>
    ''' <param name="pInfo">输出的数组引用</param>
    ''' <remarks></remarks>
    Private Sub GetTopicInfo(ByVal code As String, ByVal divBegin As String, ByVal divEnd As String, ByRef pInfo() As String)
        Dim beGins As Integer = 1
        Dim lgStart As Integer = 0
        Dim lgEnd As Integer = 0
        Dim nums As Integer = 0
        Dim lens As Integer = 0
        lens = Len(divBegin)
        Do While InStr(beGins, code, divBegin) > 0
            If InStr(beGins, code, divBegin) = 0 Then
                'can not find enother pointer
                'exit is ok
                Exit Sub
            End If
            lgStart = InStr(beGins, code, divBegin) + CInt(lens)  '获取开头位置
            lgEnd = InStr(lgStart, code, divEnd)  '获取尾部
            If lgEnd = 0 Then
                Exit Sub
            End If

            'Save To pInfo
            ReDim Preserve pInfo(nums)
            pInfo(nums) = Mid(code, lgStart, lgEnd - lgStart)
            beGins = lgEnd + 1
            nums = nums + 1
        Loop
    End Sub
End Class
