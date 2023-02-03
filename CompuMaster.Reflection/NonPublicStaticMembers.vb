Option Explicit On
Option Strict On

Namespace CompuMaster.Reflection

    ''' <summary>
    ''' Provides simplified InvokeMember for public members of an instance object
    ''' </summary>
    Public NotInheritable Class NonPublicStaticMembers

        Private Const MEMBERFILTER As System.Reflection.BindingFlags = System.Reflection.BindingFlags.Static Or System.Reflection.BindingFlags.NonPublic
        Private Const EXCEPTION_FILTERINFO As String = "non-public instance"

        Public Shared Function GetMembers(Of TSearchedMembers As System.Reflection.MemberInfo)(objtype As Type, expectedType As Type) As List(Of TSearchedMembers)
            Return ReflectionWorker.GetMembers(Of TSearchedMembers)(MEMBERFILTER, EXCEPTION_FILTERINFO, objtype, expectedType)
        End Function

        Public Shared Function GetMembers(Of TSearchedMembers As System.Reflection.MemberInfo)(objtype As Type, expectedType As Type, expectedName As String) As TSearchedMembers
            Return ReflectionWorker.GetMembers(Of TSearchedMembers)(MEMBERFILTER, EXCEPTION_FILTERINFO, objtype, expectedType, expectedName)
        End Function

        Public Shared Function GetMembers(Of TSearchedMembers As System.Reflection.MemberInfo)(objtype As Type, expectedName As String) As TSearchedMembers
            Return ReflectionWorker.GetMembers(Of TSearchedMembers)(MEMBERFILTER, EXCEPTION_FILTERINFO, objtype, expectedName)
        End Function

        Public Shared Function GetMembers(objtype As Type) As System.Reflection.MemberInfo()
            Return ReflectionWorker.GetMembers(MEMBERFILTER, EXCEPTION_FILTERINFO, objtype)
        End Function

        Public Shared Function InvokeFunction(Of T)(objType As Type, name As String, ParamArray values As Object()) As T
            Return ReflectionWorker.InvokeFunction(Of T)(MEMBERFILTER, EXCEPTION_FILTERINFO, Nothing, objType, name, values)
        End Function

        Public Shared Sub InvokeMethod(objType As Type, name As String, ParamArray values As Object())
            ReflectionWorker.InvokeMethod(MEMBERFILTER, EXCEPTION_FILTERINFO, Nothing, objType, name, values)
        End Sub

        Public Shared Function InvokePropertyGet(Of T)(objtype As Type, name As String) As T
            Return ReflectionWorker.InvokePropertyGet(Of T)(MEMBERFILTER, EXCEPTION_FILTERINFO, Nothing, objtype, name)
        End Function

        Public Shared Function InvokePropertyGet(Of T)(objtype As Type, name As String, propertyArrayItem As Object) As T
            Return ReflectionWorker.InvokePropertyGet(Of T)(MEMBERFILTER, EXCEPTION_FILTERINFO, Nothing, objtype, name, propertyArrayItem)
        End Function

        Public Shared Sub InvokePropertySet(Of T)(objType As Type, name As String, value As T)
            ReflectionWorker.InvokePropertySet(Of T)(MEMBERFILTER, EXCEPTION_FILTERINFO, Nothing, objType, name, value)
        End Sub

        Public Shared Sub InvokePropertySet(Of T)(objType As Type, name As String, values As T())
            ReflectionWorker.InvokePropertySet(Of T)(MEMBERFILTER, EXCEPTION_FILTERINFO, Nothing, objType, name, values)
        End Sub

        Public Shared Function InvokeFieldGet(Of T)(objtype As Type, name As String) As T
            Return ReflectionWorker.InvokeFieldGet(Of T)(MEMBERFILTER, EXCEPTION_FILTERINFO, Nothing, objtype, name)
        End Function

        Public Shared Sub InvokeFieldSet(Of T)(objType As Type, name As String, value As T)
            ReflectionWorker.InvokeFieldSet(Of T)(MEMBERFILTER, EXCEPTION_FILTERINFO, Nothing, objType, name, value)
        End Sub

        Public Shared Sub InvokeFieldSet(Of T)(objType As Type, name As String, values As T())
            ReflectionWorker.InvokeFieldSet(Of T)(MEMBERFILTER, EXCEPTION_FILTERINFO, Nothing, objType, name, values)
        End Sub

        Public Shared Function InvokeConstructor(Of T)(ParamArray args As Object()) As T
            Dim M As System.Reflection.BindingFlags = MEMBERFILTER And Not System.Reflection.BindingFlags.Instance
            Return ReflectionWorker.InvokeConstructor(Of T)(MEMBERFILTER, EXCEPTION_FILTERINFO, args)
        End Function

    End Class

End Namespace