Public Class BusinessErr
    Inherits System.Exception

    Private Err As String
    Public Property ErrMessage()
        Get
            Return Err
        End Get
        Set(ByVal Value)

        End Set
    End Property
End Class
