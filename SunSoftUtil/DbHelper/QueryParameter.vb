Namespace DbHelper
    Public Class QueryParameter
        Private ReadOnly _value As Object '这个变量是通用变量，这个类的作用就是统一的容纳各个类型的变量
        Public Sub New(ByVal val As Object)
            Me._value = val
        End Sub

        Public Shared Narrowing Operator CType(ByVal value As String) As QueryParameter
            Return New QueryParameter(value)
        End Operator

        Public Shared Widening Operator CType(ByVal value As Boolean) As QueryParameter
            Return New QueryParameter(value)
        End Operator

        Public Shared Widening Operator CType(ByVal value As Byte) As QueryParameter
            Return New QueryParameter(value)
        End Operator

        Public Shared Widening Operator CType(ByVal value As DateTime) As QueryParameter
            Return New QueryParameter(value)
        End Operator

        Public Shared Widening Operator CType(ByVal value As DBNull) As QueryParameter
            Return New QueryParameter(value)
        End Operator

        Public Shared Widening Operator CType(ByVal value As Decimal) As QueryParameter
            Return New QueryParameter(value)
        End Operator

        Public Shared Widening Operator CType(ByVal value As Double) As QueryParameter
            Return New QueryParameter(value)
        End Operator

        Public Shared Widening Operator CType(ByVal value As Guid) As QueryParameter
            Return New QueryParameter(value)
        End Operator

        Public Shared Widening Operator CType(ByVal value As Short) As QueryParameter
            Return New QueryParameter(value)
        End Operator

        Public Shared Widening Operator CType(ByVal value As Integer) As QueryParameter
            Return New QueryParameter(value)
        End Operator

        Public Shared Widening Operator CType(ByVal value As Long) As QueryParameter
            Return New QueryParameter(value)
        End Operator

        Public Shared Widening Operator CType(ByVal value As Single) As QueryParameter
            Return New QueryParameter(value)
        End Operator

        Public Shared Widening Operator CType(ByVal value As Byte()) As QueryParameter
            Return New QueryParameter(value)
        End Operator


        ' Properties
        Public ReadOnly Property Value As Object
            Get
                Return Me._value
            End Get
        End Property
    End Class
End Namespace