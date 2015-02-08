Imports System.Windows.Forms
Imports System.Runtime.InteropServices

Public Class RegHotKey
    '引入系统API
    <DllImport("user32.dll")> _
    Private Shared Function RegisterHotKey(ByVal hWnd As IntPtr, ByVal id As Integer, ByVal modifiers As Integer, ByVal vk As Keys) As Boolean
    End Function
    <DllImport("user32.dll")> _
    Private Shared Function UnregisterHotKey(ByVal hWnd As IntPtr, ByVal id As Integer) As Boolean
    End Function


    Private _keyId As Integer = 10
    '区分不同的快捷键
    Private ReadOnly _keyMap As New Dictionary(Of Integer, HotKeyCallBackHanlder)()
    '每一个key对于一个处理函数
    Public Delegate Sub HotKeyCallBackHanlder()

    '组合控制键
    Public Enum HotkeyModifiers
        Alt = 1
        Control = 2
        Shift = 4
        Win = 8
    End Enum

    '注册快捷键
    ''' <summary>
    ''' 注册快捷键
    ''' </summary>
    ''' <param name="hWnd">窗口句柄</param>
    ''' <param name="modifiers"></param>
    ''' <param name="vk">虚拟按键</param>
    ''' <param name="callBack">回调函数</param>
    ''' <remarks></remarks>
    Public Sub Regist(ByVal hWnd As IntPtr, ByVal modifiers As Integer, ByVal vk As Keys, ByVal callBack As HotKeyCallBackHanlder)
        Dim id As Integer = System.Math.Max(System.Threading.Interlocked.Increment(_keyId), _keyId - 1)
        If Not RegisterHotKey(hWnd, id, modifiers, vk) Then
            Throw New Exception("注册失败！")
        End If
        If _keyMap.ContainsKey(id) Then
            _keyMap(id) = callBack
        Else
            _keyMap.Add(id, callBack)
        End If
    End Sub

    ' 注销快捷键
    Public Sub UnRegist(ByVal hWnd As IntPtr, ByVal callBack As HotKeyCallBackHanlder)
        For Each handleValue As KeyValuePair(Of Integer, HotKeyCallBackHanlder) In _keyMap
            If handleValue.Value.Equals(callBack) Then
                UnregisterHotKey(hWnd, handleValue.Key)
                Return
            End If
        Next
    End Sub

    ' 快捷键消息处理
    Public Sub ProcessHotKey(ByVal m As Message)
        If m.Msg = &H312 Then
            Dim id As Integer = m.WParam.ToInt32()
            Dim callback As HotKeyCallBackHanlder
            If _keyMap.ContainsKey(id) Then
                callback = _keyMap.Item(id)
                callback()
            End If
        End If
    End Sub
End Class
