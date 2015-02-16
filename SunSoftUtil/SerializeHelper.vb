Imports System.Runtime.Serialization.Formatters.Binary
Imports System.IO

Public Class SerializeHelper
    Public Shared Function DeserializeFromFile(ByVal filePath As String) As Object
        Using fo As New FileStream(filePath, FileMode.Open)
            Dim deserizer As New BinaryFormatter
            Return deserizer.Deserialize(fo)
        End Using
    End Function

    Public Shared Sub SerializeObjToFile(ByVal obj As Object, ByVal filePath As String)
        Using fo As New FileStream(filePath, FileMode.Create)
            Dim serizer As New BinaryFormatter
            serizer.Serialize(fo, obj)
        End Using
    End Sub
End Class
