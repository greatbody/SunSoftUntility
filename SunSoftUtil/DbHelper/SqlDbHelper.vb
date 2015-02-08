Imports System.Text
Imports System.Data.SqlClient

Namespace DbHelper
    Public Class SqlDbHelper
        Private Shared _uniConnString As String
        Private _conn As SqlConnection
        Private _connString As String = ""
        Private ReadOnly _command As SqlCommand
        Private _sqlBuilder As StringBuilder
        Private _count As Integer = 0
        Public Property ConnectionString() As String
            Get
                Return _connString
            End Get
            Set(ByVal value As String)
                _connString = value
            End Set
        End Property
        Public Shared Sub SetUnsafeConnString(ByVal connString As String)
            _uniConnString = connString
        End Sub
        Public Sub New()
            '构造函数仅仅负责初始化连接字符串、命令文本对象、连接对象
            _command = New SqlCommand
            _sqlBuilder = New StringBuilder("")
            If String.IsNullOrEmpty(_uniConnString) = True Then
                _connString = ""
            Else
                '静态连接字符串存在内容
                _connString = _uniConnString
            End If
        End Sub
        Public Sub New(ByVal connString As String)
            _connString = connString
            _command = New SqlCommand
            _sqlBuilder = New StringBuilder("")
        End Sub

        Private Sub CreateConn(ByVal connString As String)
            '改动：2014年3月20日22:05:34:
            '功能：将打开连接的功能关闭, 仅在执行的时候打开, 执行后关闭
            If String.IsNullOrEmpty(connString) = True Then
                Throw New Exception("请配置初始化的连接字符串，否则请提供连接字符串")
            End If
            _connString = connString

            _conn = New SqlConnection(_connString)
            _command.Connection = _conn
            _conn.Open()
        End Sub

        Public Function Create(ByVal connString As String)
            '返回一个对象自身
            Return New SqlDbHelper(connString)
        End Function
        Public Shared Function Create() As SqlDbHelper
            Return New SqlDbHelper()
        End Function

        Public Function ExcuteNonQuery() As Integer
            Return ExcuteNonQuery(Me._command)
        End Function

        Public Function ExcuteNonQuery(ByVal sqlComm As SqlClient.SqlCommand) As Integer
            CreateConn(_connString)
            sqlComm.CommandText = _sqlBuilder.ToString
            Return sqlComm.ExecuteNonQuery
            _conn.Close()
        End Function

        Public Function ExecuteScalar(Of T)() As T
            Return ExecuteScalar(Of T)(Me._command)
        End Function
        Public Function ExecuteScalar(Of T)(ByVal sqlComm As SqlClient.SqlCommand) As T
            '更新：改为在执行中打开连接，执行后关闭
            Dim obj As Object
            CreateConn(_connString)
            sqlComm.CommandText = _sqlBuilder.ToString
            obj = sqlComm.ExecuteScalar
            _conn.Close()
            If ((obj Is Nothing) OrElse DBNull.Value.Equals(obj)) Then
                Return CType(Nothing, T)
            End If
            If (obj.GetType Is GetType(T)) Then
                Return DirectCast(obj, T)
            End If
            Return DirectCast(Convert.ChangeType(obj, GetType(T)), T)
        End Function

        Public Function FillDataSet() As System.Data.DataSet
            Return FillDataSet(Me._command)
        End Function

        Public Function FillDataTable() As System.Data.DataTable
            Return FillDataTable(Me._command)
        End Function

        Public Shared Function From(ByVal sqlStr As String) As SqlDbHelper
            Return New SqlDbHelper(sqlStr)
        End Function

        Public Function FillDataTable(ByVal sqlComm As SqlClient.SqlCommand) As System.Data.DataTable
            '更新：改为在执行中打开连接，执行后关闭
            Dim _DataTable As New DataTable("_sunsoft")

            sqlComm.CommandText = _sqlBuilder.ToString
            CreateConn(_connString)
            Using dbReader As Data.SqlClient.SqlDataReader = sqlComm.ExecuteReader
                _DataTable.Load(dbReader)
            End Using
            _conn.Close()

            Return _DataTable
        End Function

        Public Function FillDataSet(ByVal sqlComm As SqlCommand) As DataSet
            '更新：改为在执行中打开连接，执行后关闭
            Dim dataSet As New DataSet
            sqlComm.CommandText = _sqlBuilder.ToString
            CreateConn(_connString)
            Dim dataAdpt As SqlDataAdapter = New SqlDataAdapter(sqlComm)
            dataAdpt.Fill(dataSet)
            _conn.Close()
            Dim sCount As Integer
            For sCount = 0 To dataSet.Tables.Count - 1
                dataSet.Tables.Item(sCount).TableName = ("_sunsoft" & sCount.ToString)
            Next sCount
            Return dataSet
        End Function


        '''后面的代码用于运算符重载相关的代码
        Public Shared Operator +(ByVal meDefault As SqlDbHelper, ByVal strp As String) As SqlDbHelper
            '完成:运算符重载
            meDefault._addSqlText(strp)
            Return meDefault
        End Operator

        Private Sub _addSqlText(ByVal paramStr As String)
            _sqlBuilder.Append(paramStr)
        End Sub
        Public Shared Operator +(ByVal meDefault As SqlDbHelper, ByVal strp As QueryParameter) As SqlDbHelper
            meDefault._addParam(strp)
            Return meDefault
        End Operator
        ''' <summary>
        ''' 添加参数，重载，参数类型
        ''' </summary>
        ''' <param name="p">直接数据类型</param>
        ''' <remarks></remarks>
        Private Sub _addParam(ByVal p As QueryParameter)
            _sqlBuilder.Append("@p" & _count.ToString)
            _command.Parameters.AddWithValue("@p" & _count.ToString, p.Value)
            _count += 1
        End Sub

        Public Sub Dispose()
            _conn.Dispose()
            _connString = Nothing
            _command.Dispose()
            _sqlBuilder = Nothing
            _count = Nothing
        End Sub
    End Class
End Namespace

