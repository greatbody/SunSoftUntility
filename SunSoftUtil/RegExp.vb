Imports System.Text.RegularExpressions
Public Class RegExp
    Inherits System.Text.RegularExpressions.Regex
    Private _regPre As System.Text.RegularExpressions.Regex
    Private _regStr As String = ""
    Public Property RegString() As String
        Get
            Return _regStr
        End Get
        Set(ByVal value As String)
            _regStr = value
        End Set
    End Property

    Public Sub New()
        _regPre = New System.Text.RegularExpressions.Regex
    End Sub
    Public Sub New(ByVal regExpression As String)
        _regPre = New Regex(regExpression)
    End Sub
    'Public Function ToArray() As String()
    '    'Dim regExString() As String

    'End Function
    Public Shared Function GetMatches(ByVal strData As String, ByVal strRegExp As String) As String()
        Dim TmpRegEx As New Regex(strRegExp)
        Dim TmpCollect As MatchCollection = TmpRegEx.Matches(strData)
        Dim TmpMatch As Match
        Dim iRet() As String
        Dim iCount As Integer = 0
        If TmpCollect.Count > 0 Then
            ReDim iRet(TmpCollect.Count - 1)
        Else
            ReDim iRet(0)
            iRet(0) = ""
            Return iRet
        End If
        For Each TmpMatch In TmpCollect
            iRet(iCount) = TmpMatch.Value
            iCount += 1
        Next
        Return iRet
    End Function
    Public Shared Function GetFirstMatch(ByVal strData As String, ByVal strRegExp As String) As String
        Dim TmpRegEx As New Regex(strRegExp)
        Dim TmpCollect As MatchCollection = TmpRegEx.Matches(strData)
        If TmpCollect.Count <= 0 Then
            Return ""
        End If
        Return TmpCollect.Item(0).Value
    End Function
End Class
