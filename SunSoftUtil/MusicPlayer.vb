Imports System.Threading
Imports WMPLib

Public Class MusicPlayer
#Region "私有变量"
    Private _mAudio As WMPLib.WindowsMediaPlayer
#End Region
#Region "构造函数"
    Public Sub New(ByVal filePath As String)
        _mAudio = New WindowsMediaPlayer()

    End Sub

    Public Sub ChangeMusic(ByVal filePath As String)
        If Not _mAudio Is Nothing Then
            _mAudio = Nothing
        End If
        _mAudio = New Audio(filePath)
    End Sub

    Public Sub SetVolume(ByVal newVolume As Integer)
        _mAudio.Volume = newVolume
    End Sub

    Public Sub Play()
        _mAudio.controls.stop()
    End Sub

    Public Sub PlayAt(ByVal OrderDate As Date)
        Call (New Thread(AddressOf mWaitToPlay)).Start(OrderDate)
    End Sub

    Public Sub mWaitToPlay(ByVal OrderDate As Date)
        While True
            Thread.Sleep(500)
            If OrderDate <= Now Then
                Play()
                Exit Sub
            End If
        End While
    End Sub
#Region "内部方法"

#End Region

#End Region
#Region "方法"

#End Region

End Class
