Public Class DataBaseErr
    Inherits System.Exception

    '������ͻ����
    Public Const UpdateCommandErr As String = "��һ�û��Ѷ�Դ���ݽ����޸�,��Ը�Ӧ�����´���!"

    Private Err As String

    Public Property ErrMessage()
        Get
            Return Err
        End Get
        Set(ByVal Value)

        End Set
    End Property

End Class
