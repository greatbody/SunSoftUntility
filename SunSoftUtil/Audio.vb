Imports System.Security.Cryptography
Imports System.Text
Imports System.Threading

Public Class Audio
    Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrRetumString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
    Private _mFilePath As String = ""
    Private _hashName As String = ""
    ''' <summary>
    ''' 创建音乐播放对象
    ''' </summary>
    ''' <param name="filePath">文件路径</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal filePath As String)
        _mFilePath = filePath
        _hashName = BitConverter.ToString(CType(CryptoConfig.CreateFromName("MD5"), HashAlgorithm).ComputeHash((New UnicodeEncoding).GetBytes(Trim(filePath))))
        mciSendString("open """ & _mFilePath & """ alias " & _hashName, 0&, 0, 0)
    End Sub



#Region "操作"
    ''' <summary>
    ''' 将当前文件切到指定路径
    ''' </summary>
    ''' <param name="filePath">文件路径</param>
    ''' <remarks></remarks>
    Public Sub ChangeTo(ByVal filePath As String)
        _mFilePath = filePath
        _hashName = BitConverter.ToString(CType(CryptoConfig.CreateFromName("MD5"), HashAlgorithm).ComputeHash((New UnicodeEncoding).GetBytes(Trim(filePath))))
        mciSendString("open """ & _mFilePath & """ alias " & _hashName, 0&, 0, 0)
    End Sub
    ''' <summary>
    ''' 播放当前文件
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Play()
        If String.IsNullOrEmpty(_mFilePath) Then
            Throw New Exception("请指定文件路径！")
        End If
        ChangeTo(_mFilePath)
        mciSendString("play " & _hashName, 0&, 0, 0)
    End Sub
    ''' <summary>
    ''' 暂停当前文件
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Pause()
        If String.IsNullOrEmpty(_mFilePath) Then
            Throw New Exception("请指定文件路径！")
        End If
        ChangeTo(_mFilePath)
        mciSendString("pause " & _hashName, 0&, 0, 0)
    End Sub
    ''' <summary>
    ''' 移动到指定时刻
    ''' </summary>
    ''' <param name="PosText"></param>
    ''' <remarks></remarks>
    Public Sub MoveTo(ByVal PosText As String)
        CurrentPos = PosText
    End Sub
#End Region

#Region "属性"
    ''' <summary>
    ''' 音量
    ''' </summary>
    ''' <remarks></remarks>
    Private _mVolume As Integer
    ''' <summary>
    ''' 设置、获取音量，范围：0-1000
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Volume() As Integer
        Get
            Dim volumeStr As String = Space(255)
            ChangeTo(_mFilePath)
            mciSendString("status " & _hashName & " volume", volumeStr, 255, 0)
            _mVolume = CInt(Trim(volumeStr))
            Return _mVolume

        End Get
        Set(ByVal value As Integer)
            _mVolume = value
            ChangeTo(_mFilePath)
            mciSendString("setaudio " & _hashName & " volume to " & _mVolume.ToString(), 0&, 0, 0)
        End Set
    End Property

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Private _mPosition As Integer
    ''' <summary>
    ''' 当前播放毫秒数
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Position() As Integer
        Get
            Dim positionStr As String = Space(255)
            ChangeTo(_mFilePath)
            mciSendString("status " & _hashName & " position", positionStr, 255, 0)
            _mPosition = CInt(Trim(positionStr))
            Return _mPosition
        End Get
        Set(ByVal value As Integer)
            _mPosition = value
        End Set
    End Property

    Private _mPositionString As String
    Public Property PositionString() As String
        Get
            Dim positionStr As String = Space(255)
            ChangeTo(_mFilePath)
            mciSendString("status " & _hashName & " position", positionStr, 255, 0)
            _mPositionString = CInt(Trim(positionStr))
            Return _mPositionString
        End Get
        Set(ByVal value As String)
            _mPositionString = value
        End Set
    End Property

    Public ReadOnly Property Length() As Integer
        Get
            Dim lengthStr As String = Space(255)
            ChangeTo(_mFilePath)
            mciSendString("status " & _hashName & " length", lengthStr, 255, 0)
            Return CInt(Trim(lengthStr))
        End Get
    End Property

    Public ReadOnly Property LengthString() As String
        Get
            Dim mMilliom As Integer = Length
            Dim mSec As Integer
            Dim mMin As Integer
            Dim mHour As Integer
            '秒
            mSec = mMilliom \ 1000
            mMilliom = mMilliom - 1000 * mSec
            '分
            mMin = mSec \ 60
            mSec = mSec - 60 * mMin
            '时
            mHour = mMin \ 60
            mMin = mMin - 60 * mHour
            Return Format(mHour, "00") & ":" & Format(mMin, "00") & ":" & Format(mSec, "00") & "." & mMilliom
        End Get
    End Property

    Public Property CurrentPos() As String
        Get
            Dim mMilliom As Integer = Position
            Dim mSec As Integer
            Dim mMin As Integer
            Dim mHour As Integer
            '秒
            mSec = mMilliom \ 1000
            mMilliom = mMilliom - 1000 * mSec
            '分
            mMin = mSec \ 60
            mSec = mSec - 60 * mMin
            '时
            mHour = mMin \ 60
            mMin = mMin - 60 * mHour
            Return Format(mHour, "00") & ":" & Format(mMin, "00") & ":" & Format(mSec, "00") & "." & mMilliom
        End Get
        Set(ByVal value As String)
            Dim timeSet() As String = Split(value, ":")
            Dim mMillion As Integer = 0
            If timeSet.Length > 0 Then
                For i As Integer = timeSet.Length - 1 To 0 Step -1
                    If i = timeSet.Length - 1 Then
                        mMillion += CDec(timeSet(i))
                    Else
                        mMillion += CDec(timeSet(i)) * 60
                    End If
                Next
                mMillion = mMillion * 1000
            End If
            ChangeTo(_mFilePath)
            mciSendString("seek " & _hashName & " to " & mMillion.ToString(), 0&, 0, 0)
        End Set
    End Property
#End Region
End Class
