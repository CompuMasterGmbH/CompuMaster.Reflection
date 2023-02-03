Option Explicit On
Option Strict On

Namespace CompuMaster.Reflection

    ''' <summary>
    ''' Provides simplified InvokeMember for public members of an instance object
    ''' </summary>
    Friend NotInheritable Class ReflectionTools

        Public Shared Function ConvertArgumentsToObjectArray(Of T)(value As T) As Object()
            Dim Args As Object()
            If value IsNot Nothing Then
                Dim ArgsList As New List(Of Object)
                ArgsList.Add(value)
                Args = ArgsList.ToArray
            Else
                Args = Nothing
            End If
            Return Args
        End Function

        Public Shared Function ConvertArgumentsToObjectArray(Of T)(values As T()) As Object()
            Dim Args As Object()
            If values IsNot Nothing Then
                Dim ArgsList As New List(Of Object)
                For MyCounter As Integer = 0 To values.Count - 1
                    ArgsList.Add(values(MyCounter))
                Next
                Args = ArgsList.ToArray
            Else
                Args = Nothing
            End If
            Return Args
        End Function

    End Class

End Namespace