Namespace CompuMaster.Reflection

    Public Class InvokeException
        Inherits System.Exception

        Public Sub New(message As String, innerException As Exception)
            MyBase.New(message, innerException)
        End Sub

    End Class

End Namespace