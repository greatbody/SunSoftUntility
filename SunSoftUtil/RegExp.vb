Imports System.Text.RegularExpressions
Public Class RegExp
    Inherits Text.RegularExpressions.Regex
    Private _regPre As Text.RegularExpressions.Regex
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
        _regPre = New Text.RegularExpressions.Regex
    End Sub
    Public Sub New(ByVal regExpression As String)
        _regPre = New Regex(regExpression)
    End Sub
    'Public Function ToArray() As String()
    '    'Dim regExString() As String

    'End Function
    Public Shared Function GetMatches(ByVal strData As String, ByVal strRegExp As String) As String()
        Dim tmpRegEx As New Regex(strRegExp)
        Dim tmpCollect As MatchCollection = tmpRegEx.Matches(strData)
        Dim tmpMatch As Match
        Dim iRet() As String
        Dim iCount As Integer = 0
        If tmpCollect.Count > 0 Then
            ReDim iRet(tmpCollect.Count - 1)
        Else
            ReDim iRet(0)
            iRet(0) = ""
            Return iRet
        End If
        For Each tmpMatch In tmpCollect
            iRet(iCount) = tmpMatch.Value
            iCount += 1
        Next
        Return iRet
    End Function
    Public Shared Function GetFirstMatch(ByVal strData As String, ByVal strRegExp As String) As String
        Dim tmpRegEx As New Regex(strRegExp)
        Dim tmpCollect As MatchCollection = tmpRegEx.Matches(strData)
        If tmpCollect.Count <= 0 Then
            Return ""
        End If
        Return tmpCollect.Item(0).Value
    End Function
End Class
