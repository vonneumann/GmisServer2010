Public Class DataBaseErr
    Inherits System.Exception

    '并发冲突错误
    Public Const UpdateCommandErr As String = "另一用户已对源数据进行修改,请对该应用重新处理!"

    Private Err As String

    Public Property ErrMessage()
        Get
            Return Err
        End Get
        Set(ByVal Value)

        End Set
    End Property

End Class
